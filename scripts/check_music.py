#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
仪式前音乐核检 —— 检查音乐文件能不能正常播放。

场景：婚礼仪式前逐个检查 BGM，避免现场"音乐放不出来"。

用法:
    python check_music.py <音乐文件夹>
    python check_music.py <音乐文件夹> --deep
    python check_music.py <单个文件.mp3>
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

AUDIO_EXTS = {".mp3", ".m4a", ".wav", ".flac", ".aac", ".ogg", ".wma", ".opus"}


def find_tool(name):
    p = shutil.which(name)
    if p:
        return p
    candidates = [
        Path.home() / "Doubao" / "ffmpeg" / "bin" / (name + ".exe"),
        Path(r"C:\ffmpeg\bin") / (name + ".exe"),
        Path(r"D:\ffmpeg\bin") / (name + ".exe"),
    ]
    for c in candidates:
        if c.exists():
            return str(c)
    return None


def probe(ffprobe, path):
    r = subprocess.run(
        [ffprobe, "-v", "error", "-show_entries",
         "format=duration,format_name", "-of", "json", str(path)],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        return False, 0, r.stderr.strip()[:120]
    try:
        fmt = json.loads(r.stdout).get("format", {})
        return True, float(fmt.get("duration") or 0), fmt.get("format_name", "?")
    except Exception:
        return False, 0, "解析失败"


def deep(ffmpeg, path):
    r = subprocess.run(
        [ffmpeg, "-v", "error", "-i", str(path), "-f", "null", "-"],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    return r.returncode == 0, r.stderr.strip()[:120]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("target")
    ap.add_argument("--deep", action="store_true")
    args = ap.parse_args()

    ffprobe = find_tool("ffprobe")
    ffmpeg = find_tool("ffmpeg")
    if not ffprobe or not ffmpeg:
        print("[错误] 找不到 ffprobe/ffmpeg")
        return 1

    target = Path(args.target)
    files = [target] if target.is_file() else sorted(
        p for p in target.rglob("*") if p.suffix.lower() in AUDIO_EXTS)

    if not files:
        print("没找到音乐文件。支持:", " ".join(sorted(AUDIO_EXTS)))
        return 1

    print("仪式前音乐核检（%d 个文件）\n" % len(files))
    ok, bad, warn = 0, 0, []
    for f in files:
        good, dur, info = probe(ffprobe, f)
        if not good:
            bad += 1
            print("X 打不开  %s\n      %s" % (f.name, info))
            continue
        if dur <= 0.5:
            warn.append(f)
            print("! 疑似异常 %s（%.1f 秒）" % (f.name, dur))
            continue
        if args.deep:
            dg, de = deep(ffmpeg, f)
            if not dg:
                bad += 1
                print("X 解码失败 %s\n      %s" % (f.name, de))
                continue
        ok += 1
        print("OK %-30s %5.0f 秒  %s" % (f.name, dur, info))

    print("\n=== 结果 ===")
    print("  正常 %d · 异常 %d · 疑似 %d" % (ok, bad, len(warn)))
    if bad or warn:
        print("  !! 有问题的文件需替换，不能进仪式流程")
        return 1
    print("  全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())