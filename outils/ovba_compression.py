"""Compression MS-OVBA (§2.4.1) — remplace celle du paquet ms_ovba, qui corrompt certaines données."""
import struct


def _copy_token_params(decompressed_current, chunk_start):
    difference = decompressed_current - chunk_start
    bit_count = max((difference - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    maximum_length = length_mask + 3
    return bit_count, maximum_length


def _compress_chunk(data, start, end):
    out = bytearray()
    pos = start
    while pos < end:
        flag_index = len(out); out.append(0); flags = 0
        for bit in range(8):
            if pos >= end:
                break
            bit_count, max_len = _copy_token_params(pos, start)
            best_len, best_off = 0, 0
            candidate = pos - 1
            while candidate >= start:
                length = 0
                while (pos + length < end and length < max_len
                       and data[candidate + length] == data[pos + length]):
                    length += 1
                if length > best_len:
                    best_len, best_off = length, pos - candidate
                candidate -= 1
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack('<H', token); flags |= 1 << bit; pos += best_len
            else:
                out.append(data[pos]); pos += 1
        out[flag_index] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    data = bytes(data)
    out = bytearray(b'\x01')
    for start in range(0, len(data), 4096):
        end = min(start + 4096, len(data))
        body = _compress_chunk(data, start, end)
        if len(body) > 4096:
            if end - start < 4096:   # cas impossible pour du code source (données incompressibles)
                raise ValueError('dernier bloc incompressible')
            out += struct.pack('<H', 0x3000 | 4095) + data[start:end]   # bloc brut de 4096 octets
            continue
        out += struct.pack('<H', 0xB000 | (len(body) + 2 - 3)) + body
    return bytes(out)
