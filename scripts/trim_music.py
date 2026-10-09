#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
音频裁剪 —— 截取音乐某一段（从 X 秒到 Y 秒）。

用法:
    python trim_music.py 输入.mp3 30 120          # 裁 30 到 120 秒
    python trim_music.py 输入.mp3 01:30 02:00     # 分:秒格式
    python trim_music.py 输入.mp3 30 -o 输出.mp3  # 从 30 秒到结尾
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def find_ffmpeg():
    p = shutil.which("ffmpeg")
    if p:
        return p
    c = Path.home() / "Doubao" / "ffmpeg" / "bin" / "ffmpeg.exe"
    return str(c) if c.exists() else None


def parse_time(s):
    s = str(s).strip()
    if ":" in s:
        parts = s.split(":")
        if len(parts) == 3:
            h, m, sec = parts
            return int(h) * 3600 + int(m) * 60 + float(sec)
        if len(parts) == 2:
            m, sec = parts
            return int(m) * 60 + float(sec)
    return float(s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("start")
    ap.add_argument("end", nargs="?")
    ap.add_argument("-o", "--output")
    args = ap.parse_args()

    ffmpeg = find_ffmpeg()
    if not ffmpeg:
        print("[错误] 找不到 ffmpeg")
        return 1

    src = Path(args.input)
    if not src.exists():
        print("[错误] 文件不存在:", src)
        return 1

    out = Path(args.output) if args.output else src.with_name(src.stem + "_裁剪" + src.suffix)
    start = parse_time(args.start)
    end = parse_time(args.end) if args.end else None

    cmd = [ffmpeg, "-y", "-i", str(src), "-ss", str(start)]
    if end is not None:
        cmd += ["-to", str(end)]
    cmd += ["-c", "copy", str(out)]

    print("裁剪: %s  [%s -> %s]" % (src.name, args.start, args.end or "结尾"))
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print("[失败]", r.stderr.strip()[:300])
        return 1
    print("完成:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())