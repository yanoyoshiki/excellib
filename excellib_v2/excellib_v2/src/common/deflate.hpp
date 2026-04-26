#pragma once
#include <cstdint>
#include <cstddef>
#include <vector>

namespace excellib::detail {

// CRC32 (IEEE 802.3 多項式、zlibのcrc32と同一)
uint32_t crc32_compute(const uint8_t* data, size_t len);

// Raw DEFLATE (RFC 1951) 解凍 — zlib/gzip ヘッダーなし
// hint: 展開後サイズの目安 (0 可)
std::vector<uint8_t> deflate_decompress(const uint8_t* src, size_t src_len, size_t hint = 0);

} // namespace excellib::detail
