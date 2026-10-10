# -*- coding: utf-8 -*-
"""可靠测量：按亮度分层提取「人物 / 灰白边 / 背景」，逐列输出三项指标
   不做任何"暗->亮"假设，而是找亮度剖面上的两处分界。"""
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from PIL import Image
import numpy as np
R=r"D:\Documents\ChatGPT\头脑风暴"
def measure(path, x0, x1, ys, ye, tag):
    a=np.array(Image.open(path).convert("RGB")).astype(int); lum=a.mean(2)
    rows=[]
    for x in range(x0,x1+1):
        col=lum[ys:ye,x]
        dark=np.where(col<120)[0]          # 人物
        if len(dark)==0: continue
        y_person=ys+dark.max()             # 人物下边界
        rest=col[dark.max()+1:]
        # 灰度带 = 比"背景基准"暗的部分；背景基准取该列末段中位数
        if len(rest)<20: continue
        bg=np.median(rest[-15:])
        gray=np.where(rest < bg-6)[0]      # 比背景暗 => 灰白边
        if len(gray)==0: continue
        w=gray.max()-gray.min()+1
        rows.append((x, y_person, ys+dark.max()+1+gray.min(), w, bg))
    if not rows:
        print(tag, "无有效列"); return None
    arr=np.array(rows)
    print("="*64)
    print(tag, " 有效列 %d" % len(arr))
    print("  ① 撕裂线(人物下边界) y: %d ~ %d  均值 %.1f" % (arr[:,1].min(),arr[:,1].max(),arr[:,1].mean()))
    print("  ② 灰白边宽度: min=%d max=%d 均值=%.1f 中位=%.1f" % (arr[:,3].min(),arr[:,3].max(),arr[:,3].mean(),np.median(arr[:,3])))
    print("  ③ 背景亮度中位: %.1f" % np.median(arr[:,4]))
    # 灰白边实际亮度
    vals=[]
    for x,y0,gy0,ww,bg in rows[:200]:
        vals.append(lum[int(gy0):int(gy0)+int(ww), int(x)].mean())
    print("  ④ 灰白边亮度均值: %.1f  (背景 %.1f, 差 %.1f)" % (np.mean(vals), np.median(arr[:,4]), np.mean(vals)-np.median(arr[:,4])))
    return arr

REF=R+r"\poster_source_round2\赵丽海报关键源文件_第二轮\artifact-template-poster\assets\reference.png"
import os
MINE=sys.argv[1] if len(sys.argv)>1 else R+r"\poster_build_out\张文学员海报-定稿9.png"
e1=measure(REF, 560, 860, 1300, 1440, "母版 TEMPLATE")
e2=measure(MINE, 480, 900, 1270, 1440, "我的 MINE")

