// WorkbookFactory の実装ファイル
// ファイルのフォーマットを判定して適切なパーサーに処理を委譲する
#include "excellib/workbook.hpp"
#include "excellib/xls_parser.hpp"
#include "excellib/xlsx_parser.hpp"
// std::ifstream: ファイルを読み込むためのストリーム
#include <fstream>
// std::istreambuf_iterator: ストリームからバイト列を読み込むイテレータ
#include <iterator>
// std::memcmp: バイト列を比較する C 標準関数
#include <cstring>

namespace excellib {

// detect_format(): バイト列の先頭を見てファイルフォーマットを判定する
// 拡張子ではなく「マジックバイト」(ファイルの先頭数バイトの固定パターン)で判定
// 戻り値: XLS または XLSX (不明な場合は FormatError を投げる)
static FileFormat detect_format(const std::vector<uint8_t>& data) {
    // 最低 8 バイト必要 (OLE2 マジックが 8 バイト)
    if (data.size() < 8)
        throw ParseError("File too small to detect format (< 8 bytes)");

    // OLE2 マジックバイト: XLS は OLE2 コンテナ形式なのでこの 8 バイトで始まる
    // 0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1 が OLE2 のシグネチャ
    static constexpr uint8_t OLE2[8] = {0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1};
    // memcmp でマジックバイトを比較 (0 なら一致)
    if (std::memcmp(data.data(), OLE2, 8) == 0)
        return FileFormat::XLS;    // OLE2 シグネチャ → XLS ファイル

    // XLSX は ZIP 形式 (PK アーカイブ) なので先頭 4 バイトは 50 4B 03 04 (= "PK\x03\x04")
    if (data[0]==0x50 && data[1]==0x4B && data[2]==0x03 && data[3]==0x04)
        return FileFormat::XLSX;    // ZIP シグネチャ → XLSX ファイル

    // どちらにも一致しない場合はエラー (先頭4バイトをエラーメッセージに含める)
    throw FormatError("Unknown file format. "
                      "First 4 bytes: "
                      + std::to_string(int(data[0])) + " "
                      + std::to_string(int(data[1])) + " "
                      + std::to_string(int(data[2])) + " "
                      + std::to_string(int(data[3])));
}

// WorkbookFactory::open() の実装: ファイルパスからワークブックを開く
std::unique_ptr<Workbook> WorkbookFactory::open(
    const std::string& path, const OpenOptions& opts)
{
    // std::ios::binary: テキストモードの改行変換を防ぐためのバイナリモード
    std::ifstream f(path, std::ios::binary);
    // ファイルを開けなかった場合は IOError を投げる
    if (!f) throw IOError("Cannot open file: " + path);
    // イテレータを使ってファイル全体をバイト列に読み込む
    // istreambuf_iterator<char>(f) はストリームの先頭
    // istreambuf_iterator<char>() はストリームの末尾を表す
    std::vector<uint8_t> data(
        (std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
    // 空ファイルの場合は IOError を投げる
    if (data.empty()) throw IOError("File is empty: " + path);
    // バイト列版の open() に処理を委譲
    return open(data, opts);
}

// WorkbookFactory::open() の実装: バイト列からワークブックを開く
std::unique_ptr<Workbook> WorkbookFactory::open(
    const std::vector<uint8_t>& data, const OpenOptions& opts)
{
    // フォーマットを判定して対応するパーサーに処理を委譲
    switch (detect_format(data)) {
        case FileFormat::XLS: {
            // XLS: BiffRecord (BIFF8 形式) のパーサーを使用
            xls::XlsParser p(opts); return p.parse(data);
        }
        case FileFormat::XLSX: {
            // XLSX: ZIP/XML のパーサーを使用
            xlsx::XlsxParser p(opts); return p.parse(data);
        }
        // default には到達しないはずだが、コンパイラの警告を防ぐために記述
        default: throw FormatError("Unsupported format");
    }
}

// WorkbookFactory::create() の実装: 新規の空ワークブックを作成する
std::unique_ptr<Workbook> WorkbookFactory::create(FileFormat fmt) {
    switch (fmt) {
        // XLSX: XlsxWorkbook の新しいインスタンスを作成して返す
        case FileFormat::XLSX: return std::make_unique<xlsx::XlsxWorkbook>();
        // XLS: XlsWorkbook の新しいインスタンスを作成して返す
        case FileFormat::XLS:  return std::make_unique<xls::XlsWorkbook>();
        // Auto またはそれ以外: デフォルトは XLSX
        default:               return std::make_unique<xlsx::XlsxWorkbook>();
    }
}

} // namespace excellib
