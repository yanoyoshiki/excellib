// DEFLATE解凍器 (RFC 1951) + CRC32 — 外部依存なし

#include "deflate.hpp"
#include <array>
#include <stdexcept>
#include <algorithm>

namespace excellib::detail {

// ─── CRC32 (IEEE 802.3 / zlib互換) ────────────────────────────────────────

static const std::array<uint32_t, 256>& crc32_table() {
    static const auto tbl = []() {
        std::array<uint32_t, 256> t{};
        for (uint32_t n = 0; n < 256; ++n) {
            uint32_t c = n;
            for (int k = 0; k < 8; ++k)
                c = (c & 1u) ? (0xEDB88320u ^ (c >> 1)) : (c >> 1);
            t[n] = c;
        }
        return t;
    }();
    return tbl;
}

uint32_t crc32_compute(const uint8_t* data, size_t len) {
    const auto& t = crc32_table();
    uint32_t crc = 0xFFFFFFFFu;
    for (size_t i = 0; i < len; ++i)
        crc = t[(crc ^ data[i]) & 0xFFu] ^ (crc >> 8);
    return crc ^ 0xFFFFFFFFu;
}

// ─── ビット読み取り器 (LSB優先 / RFC 1951) ────────────────────────────────

struct BitReader {
    const uint8_t* src;
    const uint8_t* end;
    uint32_t buf  = 0;
    int      bits = 0;

    BitReader(const uint8_t* s, size_t n) : src(s), end(s + n) {}

    void refill() {
        while (bits <= 24 && src < end) {
            buf |= uint32_t(*src++) << bits;
            bits += 8;
        }
    }

    uint32_t peek(int n) { refill(); return buf & ((1u << n) - 1u); }

    uint32_t read(int n) {
        uint32_t v = peek(n);
        buf >>= n; bits -= n;
        return v;
    }

    // 現在バイト内の残りビットを捨ててバイト境界に揃える
    void align() { int r = bits & 7; buf >>= r; bits -= r; }

    uint16_t read_u16_aligned() {
        align();
        uint8_t lo = uint8_t(read(8));
        uint8_t hi = uint8_t(read(8));
        return uint16_t(lo) | (uint16_t(hi) << 8);
    }
};

// ─── 正規ハフマン復号器 ───────────────────────────────────────────────────

struct HuffDecoder {
    static constexpr int FB = 9; // 高速テーブルのビット幅

    uint16_t fast_sym[1 << FB]; // 復号シンボル (0xFFFF = 該当なし)
    uint8_t  fast_len[1 << FB]; // 符号長

    struct OverflowEntry { uint32_t code; uint8_t len; uint16_t sym; };
    std::vector<OverflowEntry> overflow; // FB超えの符号

    void build(const uint8_t* lens, int n) {
        // 符号長ごとのカウント
        int cnt[16] = {};
        for (int i = 0; i < n; ++i)
            if (lens[i] > 0) cnt[lens[i]]++;

        // 各符号長の先頭正規符号を計算
        uint32_t code = 0;
        uint32_t next[16] = {};
        for (int l = 1; l <= 15; ++l) {
            next[l] = code;
            code += cnt[l];
            code <<= 1;
        }

        std::fill(fast_sym, fast_sym + (1 << FB), uint16_t(0xFFFF));
        std::fill(fast_len, fast_len + (1 << FB), uint8_t(0));
        overflow.clear();

        for (int sym = 0; sym < n; ++sym) {
            int l = lens[sym];
            if (!l) continue;
            uint32_t c = next[l]++;

            if (l <= FB) {
                // LSB優先読み出し用にビット反転してテーブルへ
                uint32_t rev = 0;
                for (int k = 0; k < l; ++k) rev = (rev << 1) | ((c >> k) & 1u);
                for (int ext = 0; ext < (1 << (FB - l)); ++ext) {
                    uint32_t idx = rev | (uint32_t(ext) << l);
                    fast_sym[idx] = uint16_t(sym);
                    fast_len[idx] = uint8_t(l);
                }
            } else {
                overflow.push_back({c, uint8_t(l), uint16_t(sym)});
            }
        }
    }

    int decode(BitReader& br) {
        // 高速パス: FB ビット以内で確定
        br.refill();
        uint32_t idx = br.buf & ((1u << FB) - 1u);
        if (fast_sym[idx] != 0xFFFF) {
            int l = fast_len[idx];
            br.buf >>= l; br.bits -= l;
            return fast_sym[idx];
        }

        // 低速パス: 1ビットずつ読んでMSB優先で符号を積み上げる
        uint32_t code = 0;
        for (int l = 1; l <= 15; ++l) {
            br.refill();
            code = (code << 1) | (br.buf & 1u);
            br.buf >>= 1; br.bits--;
            for (auto& e : overflow)
                if (e.len == uint8_t(l) && e.code == code)
                    return e.sym;
        }
        throw std::runtime_error("DEFLATE: 無効なハフマン符号");
    }
};

// ─── RFC 1951 静的テーブル ────────────────────────────────────────────────

static const uint16_t LENGTH_BASE[29] = {
    3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,
    35,43,51,59,67,83,99,115,131,163,195,227,258
};
static const uint8_t LENGTH_EXTRA[29] = {
    0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0
};
static const uint16_t DIST_BASE[30] = {
    1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,
    257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577
};
static const uint8_t DIST_EXTRA[30] = {
    0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13
};
// 符号長アルファベットの読み出し順 (RFC 1951 §3.2.7)
static const int CLCL_ORDER[19] = {
    16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15
};

// ─── 固定ハフマン木 (RFC 1951 §3.2.6) ─────────────────────────────────────

static void build_fixed(HuffDecoder& lit, HuffDecoder& dist) {
    uint8_t ll[288];
    for (int i = 0;   i <= 143; ++i) ll[i] = 8;
    for (int i = 144; i <= 255; ++i) ll[i] = 9;
    for (int i = 256; i <= 279; ++i) ll[i] = 7;
    for (int i = 280; i <= 287; ++i) ll[i] = 8;
    lit.build(ll, 288);

    uint8_t dl[30];
    for (int i = 0; i < 30; ++i) dl[i] = 5;
    dist.build(dl, 30);
}

// ─── ブロック復号 ─────────────────────────────────────────────────────────

static void decode_block(BitReader& br, HuffDecoder& lit, HuffDecoder& dist,
                         std::vector<uint8_t>& out) {
    for (;;) {
        int sym = lit.decode(br);
        if (sym < 256) {
            out.push_back(uint8_t(sym));
        } else if (sym == 256) {
            break; // ブロック終端
        } else {
            int li = sym - 257;
            if (li < 0 || li >= 29) throw std::runtime_error("DEFLATE: 不正な長さ符号");
            uint32_t len = LENGTH_BASE[li] + br.read(LENGTH_EXTRA[li]);

            int di = dist.decode(br);
            if (di < 0 || di >= 30) throw std::runtime_error("DEFLATE: 不正な距離符号");
            uint32_t d = DIST_BASE[di] + br.read(DIST_EXTRA[di]);

            if (d > out.size()) throw std::runtime_error("DEFLATE: 後方参照が開始位置より前");
            size_t pos = out.size() - d;
            // ローカル変数に読んでからpush_backしてreallocationによる参照無効化を防ぐ
            for (uint32_t k = 0; k < len; ++k) {
                uint8_t b = out[pos + k % d];
                out.push_back(b);
            }
        }
    }
}

// ─── メインエントリ ───────────────────────────────────────────────────────

std::vector<uint8_t> deflate_decompress(const uint8_t* src, size_t src_len, size_t hint) {
    BitReader br(src, src_len);
    std::vector<uint8_t> out;
    if (hint) out.reserve(hint);

    for (;;) {
        bool bfinal = br.read(1) != 0;
        int  btype  = int(br.read(2));

        if (btype == 0) {
            // 非圧縮ブロック
            uint16_t len  = br.read_u16_aligned();
            uint16_t nlen = br.read_u16_aligned();
            if (uint16_t(len ^ nlen) != 0xFFFFu)
                throw std::runtime_error("DEFLATE: LEN/NLEN 不一致");
            for (uint16_t k = 0; k < len; ++k)
                out.push_back(uint8_t(br.read(8)));

        } else if (btype == 1) {
            // 固定ハフマン
            HuffDecoder llit, ddist;
            build_fixed(llit, ddist);
            decode_block(br, llit, ddist, out);

        } else if (btype == 2) {
            // 動的ハフマン
            int hlit  = int(br.read(5)) + 257;
            int hdist = int(br.read(5)) + 1;
            int hclen = int(br.read(4)) + 4;

            uint8_t cl_lens[19] = {};
            for (int i = 0; i < hclen; ++i)
                cl_lens[CLCL_ORDER[i]] = uint8_t(br.read(3));

            HuffDecoder cl_dec;
            cl_dec.build(cl_lens, 19);

            std::vector<uint8_t> all_lens;
            all_lens.reserve(size_t(hlit + hdist));
            while (int(all_lens.size()) < hlit + hdist) {
                int sym = cl_dec.decode(br);
                if (sym <= 15) {
                    all_lens.push_back(uint8_t(sym));
                } else if (sym == 16) {
                    if (all_lens.empty()) throw std::runtime_error("DEFLATE: 反復元の符号長なし");
                    int rep = int(br.read(2)) + 3;
                    for (int k = 0; k < rep; ++k) all_lens.push_back(all_lens.back());
                } else if (sym == 17) {
                    int rep = int(br.read(3)) + 3;
                    for (int k = 0; k < rep; ++k) all_lens.push_back(0);
                } else if (sym == 18) {
                    int rep = int(br.read(7)) + 11;
                    for (int k = 0; k < rep; ++k) all_lens.push_back(0);
                } else {
                    throw std::runtime_error("DEFLATE: 不正な符号長シンボル");
                }
            }

            HuffDecoder llit, ddist;
            llit.build(all_lens.data(), hlit);
            ddist.build(all_lens.data() + hlit, hdist);
            decode_block(br, llit, ddist, out);

        } else {
            throw std::runtime_error("DEFLATE: 予約済みブロック種別(3)");
        }

        if (bfinal) break;
    }
    return out;
}

} // namespace excellib::detail
