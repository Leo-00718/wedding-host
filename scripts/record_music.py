#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
音乐录制 —— 把电脑正在播放的声音录下来（loopback 录音）。

场景：某些音乐下载不到，播放时录下来。
原理：录制"系统播放"的声音（loopback），不是麦克风收音。

用法:
    python record_music.py -d 30 -o 输出.wav      # 录 30 秒
    python record_music.py -o 输出.wav            # 手动停止（Ctrl+C）

依赖: pip install soundcard numpy
"""
from __future__ import annotations

import argparse
import sys
import time
import wave

import numpy as np
import soundcard as sc

SAMPLERATE = 48000


def find_loopback_mic():
    """找到 loopback 麦克风（录系统声音用）。"""
    for m in sc.all_microphones(include_loopback=True):
        if getattr(m, "isloopback", False):
            return m
    return None


def list_devices():
    print("可用的录音设备（含 loopback）：")
    for i, m in enumerate(sc.all_microphones(include_loopback=True)):
        tag = "  <- loopback（录系统声音）" if getattr(m, "isloopback", False) else ""
        print("  [%d] %s%s" % (i, m.name, tag))


def save_wav(path, data, sr):
    data = (data * 32767).astype(np.int16)
    with wave.open(str(path), "w") as w:
        w.setnchannels(data.shape[1] if data.ndim > 1 else 1)
        w.setsampwidth(2)
        w.setframerate(sr)
        w.writeframes(data.tobytes())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-d", "--duration", type=float, help="录制时长（秒）；不填则按 Ctrl+C 停止")
    ap.add_argument("-o", "--output", default="录音.wav")
    ap.add_argument("--list", action="store_true", help="列出设备")
    args = ap.parse_args()

    if args.list:
        list_devices()
        return 0

    mic = find_loopback_mic()
    if mic is None:
        print("[错误] 没找到 loopback 设备。")
        list_devices()
        return 1

    print("录制设备: %s" % mic.name)
    print("采样率: %d Hz" % SAMPLERATE)
    print("输出: %s" % args.output)

    if args.duration:
        print("开始录 %s 秒...（请现在就播放音乐）" % args.duration)
        with mic.recorder(samplerate=SAMPLERATE) as rec:
            data = rec.record(numframes=int(SAMPLERATE * args.duration))
        save_wav(args.output, data, SAMPLERATE)
        print("完成:", args.output)
    else:
        print("手动模式：请播放音乐，按 Ctrl+C 停止录音。")
        chunks = []
        try:
            with mic.recorder(samplerate=SAMPLERATE) as rec:
                while True:
                    chunks.append(rec.record(numframes=SAMPLERATE))  # 每秒一块
                    print("  已录 %d 秒" % len(chunks), end="\r")
        except KeyboardInterrupt:
            pass
        if chunks:
            data = np.concatenate(chunks, axis=0)
            save_wav(args.output, data, SAMPLERATE)
            print("\n完成（%d 秒）: %s" % (len(chunks), args.output))
        else:
            print("\n没录到任何内容。")

    return 0


if __name__ == "__main__":
    sys.exit(main())