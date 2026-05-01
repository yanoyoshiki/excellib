#pragma once
// uint8_t / size_t など標準的な整数型を使うため
#include <cstdint>
#include <cstddef>
// 動的配列型 std::vector を使うため
#include <vector>

// excellib::detail は外部公開しない内部実装の名前空間
namespace excellib::detail {

// CRC32 (IEEE 802.3 多項式、zlibのcrc32と同一)
// data: チェックサムを計算するバイト列へのポインタ
// len: バイト列の長さ
// 戻り値: 32ビットの CRC チェックサム値
uint32_t crc32_compute(const uint8_t* data, size_t len);

// Raw DEFLATE (RFC 1951) 解凍 — zlib/gzip ヘッダーなし
// src: 圧縮済みバイト列へのポインタ
// src_len: 圧縮済みデータのバイト数
// hint: 展開後サイズの目安 (0 可。指定するとメモリ確保が効率的になる)
// 戻り値: 解凍後のバイト列
std::vector<uint8_t> deflate_decompress(const uint8_t* src, size_t src_len, size_t hint = 0);

} // namespace excellib::detail
