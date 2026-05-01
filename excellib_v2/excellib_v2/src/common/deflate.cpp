// DEFLATE解凍器 (RFC 1951) + CRC32 — 外部依存なし
// zlibを使わずに、RFC 1951 で定義された DEFLATE アルゴリズムを自前実装したファイル

// 自前の deflate.hpp で宣言した関数の実装
#include "deflate.hpp"
// std::array: 固定長配列型
#include <array>
// std::runtime_error などの例外クラス
#include <stdexcept>
// std::fill: 配列・コンテナを指定値で埋めるアルゴリズム
#include <algorithm>

// 内部実装の名前空間
namespace excellib::detail {

// ─── CRC32 (IEEE 802.3 / zlib互換) ────────────────────────────────────────
// CRC32 は「データが壊れていないか確認するためのチェックサム」を計算するアルゴリズム
// ZIP ファイルはエントリごとに CRC32 を記録し、解凍後に検証する

// crc32_table(): CRC32 の計算に使うルックアップテーブルを返す関数
// static で「このファイル内だけで使う」関数にしている
// const std::array<uint32_t, 256>& を返すことでコピーを避ける
static const std::array<uint32_t, 256>& crc32_table() {
    // static const auto tbl: プログラム起動時に1回だけ初期化される静的変数
    static const auto tbl = []() {
        // テーブルの要素は 256 個 (バイト値 0〜255 に対応)
        std::array<uint32_t, 256> t{};
        // n は 0〜255 の各バイト値
        for (uint32_t n = 0; n < 256; ++n) {
            uint32_t c = n;    // c = 現在の CRC 値 (初期値はバイト値そのまま)
            // 8ビット分のシフト演算で CRC を計算 (IEEE 802.3 多項式)
            for (int k = 0; k < 8; ++k)
                // 最下位ビットが 1 なら XOR する (0xEDB88320 は生成多項式)
                // c >> 1 は右に1ビットシフト
                c = (c & 1u) ? (0xEDB88320u ^ (c >> 1)) : (c >> 1);
            // 計算結果をテーブルに格納
            t[n] = c;
        }
        // ラムダ関数の戻り値 = 完成したテーブル
        return t;
    }();    // ラムダをその場で呼び出して初期化
    // 完成したテーブルへの参照を返す
    return tbl;
}

// crc32_compute(): バイト列の CRC32 チェックサムを計算する
uint32_t crc32_compute(const uint8_t* data, size_t len) {
    // テーブルへの参照を取得
    const auto& t = crc32_table();
    // CRC の初期値は 0xFFFFFFFF (CRC32 の仕様)
    uint32_t crc = 0xFFFFFFFFu;
    // バイト列を1バイトずつ処理
    for (size_t i = 0; i < len; ++i)
        // (crc ^ data[i]) & 0xFF でテーブルのインデックスを計算
        // テーブルの値と crc の上位ビットを XOR してローリング計算
        crc = t[(crc ^ data[i]) & 0xFFu] ^ (crc >> 8);
    // 最後に 0xFFFFFFFF と XOR して最終 CRC 値を得る (CRC32 の仕様)
    return crc ^ 0xFFFFFFFFu;
}

// ─── ビット読み取り器 (LSB優先 / RFC 1951) ────────────────────────────────
// DEFLATE データはビット単位で読む必要があるためこのクラスが必要
// LSB (Least Significant Bit) 優先: 各バイトの最下位ビットを先に読む

struct BitReader {
    // src: 現在読み取り中のバイト列へのポインタ (読むたびに進める)
    const uint8_t* src;
    // end: バイト列の終端 (ここ以降は読んではいけない)
    const uint8_t* end;
    // buf: 現在バッファに溜まっているビット列 (最大32ビット)
    uint32_t buf  = 0;
    // bits: バッファ内の有効ビット数
    int      bits = 0;

    // コンストラクタ: 読み取り先バイト列のポインタと長さを受け取る
    BitReader(const uint8_t* s, size_t n) : src(s), end(s + n) {}

    // refill(): バッファが空に近いときにバイト列からビットを補充する
    void refill() {
        // バッファの有効ビットが 24 以下かつ残りデータがある間は補充する
        while (bits <= 24 && src < end) {
            // *src++ でバイトを読みつつポインタを進める
            // <<bits でシフトして既存ビットの上に積み重ねる
            buf |= uint32_t(*src++) << bits;
            bits += 8;    // 8ビット追加したのでカウントを増やす
        }
    }

    // peek(n): バッファから n ビット先読みする (バッファは変化しない)
    uint32_t peek(int n) { refill(); return buf & ((1u << n) - 1u); }

    // read(n): バッファから n ビット読み取り、読み取った分を消費する
    uint32_t read(int n) {
        uint32_t v = peek(n);    // n ビット先読み
        buf >>= n; bits -= n;    // 読んだ分をバッファから取り除く
        return v;
    }

    // align(): 現在バイト内の残りビットを捨ててバイト境界に揃える
    // 非圧縮ブロックの前に呼ぶ必要がある
    void align() { int r = bits & 7; buf >>= r; bits -= r; }

    // read_u16_aligned(): バイト境界に揃えてから 16ビット符号なし整数を読む
    // バイト境界に揃えてから 2バイト読んで LE (リトルエンディアン) で解釈
    uint16_t read_u16_aligned() {
        align();    // まずバイト境界に揃える
        uint8_t lo = uint8_t(read(8));    // 下位バイトを読む
        uint8_t hi = uint8_t(read(8));    // 上位バイトを読む
        // リトルエンディアン: 下位バイトが先に来る
        return uint16_t(lo) | (uint16_t(hi) << 8);
    }
};

// ─── 正規ハフマン復号器 ───────────────────────────────────────────────────
// ハフマン符号: 出現頻度の高いシンボルを短いビット列で表すことで圧縮する符号化方式

struct HuffDecoder {
    // FB: 高速テーブルのビット幅 (9ビット以下の符号を高速に復号)
    static constexpr int FB = 9; // 高速テーブルのビット幅

    // fast_sym[]: 9ビット以下の符号に対応するシンボル (0xFFFF = 該当なし)
    uint16_t fast_sym[1 << FB]; // 復号シンボル (0xFFFF = 該当なし)
    // fast_len[]: 各エントリの実際の符号長 (何ビット使ったか)
    uint8_t  fast_len[1 << FB]; // 符号長

    // OverflowEntry: FB ビットを超える長い符号を格納する構造体
    struct OverflowEntry { uint32_t code; uint8_t len; uint16_t sym; };
    // overflow: FB より長い符号のリスト
    std::vector<OverflowEntry> overflow; // FB超えの符号

    // build(): 符号長リストからハフマンデコーダを構築する
    // lens: 各シンボルの符号長を格納した配列
    // n: シンボルの総数
    void build(const uint8_t* lens, int n) {
        // 符号長ごとのカウント (何個のシンボルがその符号長を持つか)
        int cnt[16] = {};
        for (int i = 0; i < n; ++i)
            if (lens[i] > 0) cnt[lens[i]]++;    // 符号長 0 はそのシンボルが使われないことを意味する

        // 各符号長の先頭正規符号を計算 (RFC 1951 の正規ハフマン符号生成アルゴリズム)
        uint32_t code = 0;
        uint32_t next[16] = {};    // 各符号長の次に使う符号値
        for (int l = 1; l <= 15; ++l) {
            next[l] = code;    // 今の符号値を記録
            code += cnt[l];    // その符号長のシンボル数を加算
            code <<= 1;        // 次の符号長へ: 左シフトで符号空間を広げる
        }

        // 高速テーブルを「未登録」で初期化
        std::fill(fast_sym, fast_sym + (1 << FB), uint16_t(0xFFFF));
        std::fill(fast_len, fast_len + (1 << FB), uint8_t(0));
        overflow.clear();    // オーバーフローリストをクリア

        // 各シンボルに符号を割り当てて高速テーブルまたはオーバーフローリストに登録
        for (int sym = 0; sym < n; ++sym) {
            int l = lens[sym];
            if (!l) continue;    // 符号長 0 のシンボルはスキップ
            uint32_t c = next[l]++;    // このシンボルの符号値を取得し、次の符号値をインクリメント

            if (l <= FB) {
                // FB ビット以下の符号: 高速テーブルに登録
                // LSB優先読み出し用にビット反転してテーブルへ
                // DEFLATE は LSB (最下位ビット) から読むため、ビット順を逆にする
                uint32_t rev = 0;
                for (int k = 0; k < l; ++k) rev = (rev << 1) | ((c >> k) & 1u);
                // 符号が l ビットのとき、9ビットテーブルの残り (9-l) ビット分を展開して登録
                for (int ext = 0; ext < (1 << (FB - l)); ++ext) {
                    // idx = 反転した符号 + 残りビット (任意の値でよい)
                    uint32_t idx = rev | (uint32_t(ext) << l);
                    fast_sym[idx] = uint16_t(sym);    // シンボルを記録
                    fast_len[idx] = uint8_t(l);        // 符号長を記録
                }
            } else {
                // FB より長い符号: オーバーフローリストに追加
                overflow.push_back({c, uint8_t(l), uint16_t(sym)});
            }
        }
    }

    // decode(): BitReader からビット列を読んでシンボルを復号する
    int decode(BitReader& br) {
        // 高速パス: FB ビット以内で確定する場合
        br.refill();    // バッファを補充
        uint32_t idx = br.buf & ((1u << FB) - 1u);    // 下位 FB ビットをインデックスとして使う
        if (fast_sym[idx] != 0xFFFF) {
            // 高速テーブルにヒット: 符号長分だけバッファを消費してシンボルを返す
            int l = fast_len[idx];
            br.buf >>= l; br.bits -= l;
            return fast_sym[idx];
        }

        // 低速パス: FB ビットで確定しない場合 (長い符号)
        // 1ビットずつ読んでMSB優先で符号を積み上げる
        uint32_t code = 0;
        for (int l = 1; l <= 15; ++l) {
            br.refill();
            // LSB から1ビット読んで符号の下位に追加 (MSB優先符号に変換)
            code = (code << 1) | (br.buf & 1u);
            br.buf >>= 1; br.bits--;
            // オーバーフローリストで一致する符号を探す
            for (auto& e : overflow)
                if (e.len == uint8_t(l) && e.code == code)
                    return e.sym;
        }
        // 15ビットを読んでも一致しない = 不正なデータ
        throw std::runtime_error("DEFLATE: 無効なハフマン符号");
    }
};

// ─── RFC 1951 静的テーブル ────────────────────────────────────────────────
// DEFLATE の仕様書 (RFC 1951) で定義されている定数テーブル

// LENGTH_BASE: 長さ符号 (257〜285) に対応する実際の長さの基本値
// バックリファレンスで「何バイト戻って何バイトコピーするか」の「何バイト」部分
static const uint16_t LENGTH_BASE[29] = {
    3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,
    35,43,51,59,67,83,99,115,131,163,195,227,258
};
// LENGTH_EXTRA: 各長さ符号に対する追加ビット数 (基本値に加算する値を読むビット数)
static const uint8_t LENGTH_EXTRA[29] = {
    0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0
};
// DIST_BASE: 距離符号 (0〜29) に対応する実際の距離の基本値
// バックリファレンスで「何バイト前に戻るか」の基本値
static const uint16_t DIST_BASE[30] = {
    1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,
    257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577
};
// DIST_EXTRA: 各距離符号に対する追加ビット数
static const uint8_t DIST_EXTRA[30] = {
    0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13
};
// 符号長アルファベットの読み出し順 (RFC 1951 §3.2.7 で定義されている順序)
// 動的ハフマンブロックで符号長ハフマン木を構築するときに使う
static const int CLCL_ORDER[19] = {
    16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15
};

// ─── 固定ハフマン木 (RFC 1951 §3.2.6) ─────────────────────────────────────
// DEFLATE には固定ハフマン符号と動的ハフマン符号の2種類がある
// 固定ハフマン符号: RFC 1951 で事前に定義されたシンボル→符号長のマッピング

// build_fixed(): RFC 1951 の固定ハフマン木をデコーダに組み込む
// lit: リテラル/長さ用デコーダ  dist: 距離用デコーダ
static void build_fixed(HuffDecoder& lit, HuffDecoder& dist) {
    // 固定ハフマン符号のリテラル/長さ符号の符号長テーブル (RFC 1951 仕様)
    uint8_t ll[288];
    for (int i = 0;   i <= 143; ++i) ll[i] = 8;    // 0〜143: 8ビット
    for (int i = 144; i <= 255; ++i) ll[i] = 9;    // 144〜255: 9ビット
    for (int i = 256; i <= 279; ++i) ll[i] = 7;    // 256〜279: 7ビット (終端・長さ符号)
    for (int i = 280; i <= 287; ++i) ll[i] = 8;    // 280〜287: 8ビット
    // 符号長テーブルからデコーダを構築
    lit.build(ll, 288);

    // 距離符号はすべて 5ビット固定
    uint8_t dl[30];
    for (int i = 0; i < 30; ++i) dl[i] = 5;
    dist.build(dl, 30);
}

// ─── ブロック復号 ─────────────────────────────────────────────────────────
// decode_block(): ハフマン符号化されたブロックを復号して出力バッファに追加する
// br: ビット読み取り器  lit: リテラル/長さデコーダ  dist: 距離デコーダ
// out: 復号結果を追加するバッファ
static void decode_block(BitReader& br, HuffDecoder& lit, HuffDecoder& dist,
                         std::vector<uint8_t>& out) {
    // 256 (終端シンボル) が来るまで無限ループ
    for (;;) {
        // リテラル/長さシンボルを1つ復号
        int sym = lit.decode(br);
        if (sym < 256) {
            // 0〜255: リテラルバイト (そのまま出力に追加)
            out.push_back(uint8_t(sym));
        } else if (sym == 256) {
            // 256: ブロック終端 (ループを抜ける)
            break; // ブロック終端
        } else {
            // 257〜285: 長さ符号 (バックリファレンスの長さを表す)
            int li = sym - 257;    // 長さテーブルのインデックス (0〜28)
            if (li < 0 || li >= 29) throw std::runtime_error("DEFLATE: 不正な長さ符号");
            // 基本長 + 追加ビット分の長さを計算
            uint32_t len = LENGTH_BASE[li] + br.read(LENGTH_EXTRA[li]);

            // 距離符号を復号 (「何バイト前から」コピーするかを表す)
            int di = dist.decode(br);
            if (di < 0 || di >= 30) throw std::runtime_error("DEFLATE: 不正な距離符号");
            // 基本距離 + 追加ビット分の距離を計算
            uint32_t d = DIST_BASE[di] + br.read(DIST_EXTRA[di]);

            // 距離が出力済みデータより長い場合は不正データ
            if (d > out.size()) throw std::runtime_error("DEFLATE: 後方参照が開始位置より前");
            // バックリファレンスの起点位置
            size_t pos = out.size() - d;
            // ローカル変数に読んでからpush_backしてreallocationによる参照無効化を防ぐ
            // len バイト分コピーする (pos+k が out の末尾を超える場合はラップアラウンド)
            for (uint32_t k = 0; k < len; ++k) {
                // k % d で距離 d の範囲内でループしてコピーする (ランレングス圧縮に対応)
                uint8_t b = out[pos + k % d];
                out.push_back(b);
            }
        }
    }
}

// ─── メインエントリ ───────────────────────────────────────────────────────

// deflate_decompress(): Raw DEFLATE 圧縮データを解凍する
// src: 圧縮データのポインタ  src_len: 圧縮データのバイト数
// hint: 解凍後のサイズの目安 (0でも動くが指定するとメモリ確保が効率的)
std::vector<uint8_t> deflate_decompress(const uint8_t* src, size_t src_len, size_t hint) {
    // BitReader でビット単位の読み取りを初期化
    BitReader br(src, src_len);
    // 解凍結果を格納するバッファ
    std::vector<uint8_t> out;
    // hint が指定されていれば事前に容量を確保 (メモリの再確保を減らすため)
    if (hint) out.reserve(hint);

    // DEFLATE は複数のブロックで構成される。bfinal=1 のブロックが最後
    for (;;) {
        // BFINAL: このブロックが最終ブロックかどうか (1ビット)
        bool bfinal = br.read(1) != 0;
        // BTYPE: ブロックの圧縮方式 (2ビット)
        int  btype  = int(br.read(2));

        if (btype == 0) {
            // 非圧縮ブロック (BTYPE=00): そのままバイト列をコピー
            // LEN: ブロックのバイト数 (2バイト)
            uint16_t len  = br.read_u16_aligned();
            // NLEN: LEN の 1の補数 (整合性チェック用)
            uint16_t nlen = br.read_u16_aligned();
            // LEN と NLEN の XOR が 0xFFFF でなければデータが壊れている
            if (uint16_t(len ^ nlen) != 0xFFFFu)
                throw std::runtime_error("DEFLATE: LEN/NLEN 不一致");
            // len バイト分をそのまま出力に追加
            for (uint16_t k = 0; k < len; ++k)
                out.push_back(uint8_t(br.read(8)));

        } else if (btype == 1) {
            // 固定ハフマンブロック (BTYPE=01): RFC 1951 の固定ハフマン木を使う
            HuffDecoder llit, ddist;    // リテラル/長さ用と距離用のデコーダ
            build_fixed(llit, ddist);   // 固定ハフマン木をセットアップ
            decode_block(br, llit, ddist, out);    // ブロックを復号

        } else if (btype == 2) {
            // 動的ハフマンブロック (BTYPE=10): ブロック先頭にハフマン木の定義がある

            // HLIT: リテラル/長さ符号の数 - 257
            int hlit  = int(br.read(5)) + 257;
            // HDIST: 距離符号の数 - 1
            int hdist = int(br.read(5)) + 1;
            // HCLEN: 符号長ハフマン木の符号数 - 4
            int hclen = int(br.read(4)) + 4;

            // 符号長ハフマン木の各符号長を CLCL_ORDER の順で読む
            uint8_t cl_lens[19] = {};
            for (int i = 0; i < hclen; ++i)
                cl_lens[CLCL_ORDER[i]] = uint8_t(br.read(3));    // 各符号は 3ビット

            // 符号長ハフマン木を構築 (このデコーダで後続の符号長を読む)
            HuffDecoder cl_dec;
            cl_dec.build(cl_lens, 19);

            // リテラル/長さ符号と距離符号の符号長を読む
            std::vector<uint8_t> all_lens;
            all_lens.reserve(size_t(hlit + hdist));
            // hlit + hdist 個の符号長が読めるまでループ
            while (int(all_lens.size()) < hlit + hdist) {
                int sym = cl_dec.decode(br);    // 符号長シンボルを復号
                if (sym <= 15) {
                    // 0〜15: 実際の符号長 (そのまま追加)
                    all_lens.push_back(uint8_t(sym));
                } else if (sym == 16) {
                    // 16: 直前の符号長を 3〜6 回繰り返す
                    if (all_lens.empty()) throw std::runtime_error("DEFLATE: 反復元の符号長なし");
                    int rep = int(br.read(2)) + 3;    // 追加2ビットで 3〜6 回の繰り返し
                    for (int k = 0; k < rep; ++k) all_lens.push_back(all_lens.back());
                } else if (sym == 17) {
                    // 17: 符号長 0 を 3〜10 回繰り返す
                    int rep = int(br.read(3)) + 3;    // 追加3ビットで 3〜10 回
                    for (int k = 0; k < rep; ++k) all_lens.push_back(0);
                } else if (sym == 18) {
                    // 18: 符号長 0 を 11〜138 回繰り返す
                    int rep = int(br.read(7)) + 11;    // 追加7ビットで 11〜138 回
                    for (int k = 0; k < rep; ++k) all_lens.push_back(0);
                } else {
                    // 19以上は定義されていない不正シンボル
                    throw std::runtime_error("DEFLATE: 不正な符号長シンボル");
                }
            }

            // all_lens の前 hlit 個がリテラル/長さ符号の符号長
            // 後 hdist 個が距離符号の符号長
            HuffDecoder llit, ddist;
            llit.build(all_lens.data(), hlit);
            ddist.build(all_lens.data() + hlit, hdist);
            decode_block(br, llit, ddist, out);    // ブロックを復号

        } else {
            // BTYPE=3 は RFC 1951 で「予約済み・使用禁止」
            throw std::runtime_error("DEFLATE: 予約済みブロック種別(3)");
        }

        // bfinal=true なら最後のブロック → ループ終了
        if (bfinal) break;
    }
    // 解凍結果を返す
    return out;
}

} // namespace excellib::detail
