#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерация 10x10 RGB PNG чистым python (struct/zlib, без PIL) для смоука Task 5."""
import struct
import sys
import zlib
from pathlib import Path


def make_png(path: Path, w: int = 10, h: int = 10) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)  # 8-bit, RGB
    raw = b"".join(b"\x00" + bytes((0x40, 0x80, 0xC0)) * w for _ in range(h))
    path.write_bytes(sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw))
                     + chunk(b"IEND", b""))


if __name__ == "__main__":
    out = Path(sys.argv[1])
    out.parent.mkdir(parents=True, exist_ok=True)
    make_png(out)
    print(f"PNG записан: {out} ({out.stat().st_size} байт)")
