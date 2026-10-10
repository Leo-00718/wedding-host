# -*- coding: utf-8 -*-
"""学员欢迎/毕业海报生成 —— 固定底板 + 自动抠图 + 自动锚定

用法：
    python scripts/build_poster.py 学员照.jpg --name "张文" --cohort "第35期学员" \
        --qr 二维码.png --phone "13800000000" --out "欢迎海报_张文.png"

要点：
    - 底板 assets/poster/base.png 是唯一版式来源，不改动其固定区域
    - 人像用 rembg(u2net_human_seg) 抠图，RGB 取 matting（无灰边），
      alpha 取干净形状 + 羽化 σ2.2（无锯齿且顺滑）
    - 人物底边自动锚定「蓝笔触高度的 50%」，无需手工测量
    - 姓名/期数/电话/扫码咨询全部本地字体绘制；二维码用原图等比缩放
    - 客户照片与成品只存本地，不上传、不提交公开仓库
"""
import argparse
import json
import os
import sys
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont

SKILL_DIR = Path(__file__).resolve().parent.parent
BASE_IMG  = SKILL_DIR / "assets" / "poster" / "base.png"
TEAR_TPL  = SKILL_DIR / "assets" / "poster" / "tear_template.json"
TEAR_PROF = SKILL_DIR / "assets" / "poster" / "tear_profile.npy"
FONT_DIR  = SKILL_DIR / "assets" / "fonts"
SONG_FONT = r"C:\Windows\Fonts\simsun.ttc"
NAME_FONT = FONT_DIR / "瘦金简体.ttf"       # 姓名专用细体（同字号 = 各字等大）          # 期数/电话/扫码咨询

# ---------- 版式参数（相对底板固定）----------
L = dict(
    person_w        = 425,     # 人物显示宽度
    person_right    = 905,     # 人物右边缘 x
    person_ratio    = 1.0,     # 人物底边 = 蓝笔触顶部 + 笔触高度*ratio（1.0=压在笔触下沿，对齐母版）
    tear_lift       = 2,      # 撕纸线整体上移量（母版实测：撕纸线在人物底边上方 12px）
    tear_white      = (214, 213, 217),   # 母版实测：比背景(234)暗约 20,   # 母版实测：比背景暗约 20 的灰白（不是纯白！）,
    name_x=106, name_top=790, name_band=(790, 961), name_gap=6.5,
    name_size0=127.5, name_right_pad=2,   # 统一字号（=原170×0.75）；总宽超了就整体缩小
    name_color=(42, 42, 42),
    qi_x=130, qi_y=978, qi_size=34, qi_color=(48, 48, 48),
    phone_x=302, phone_y=1481, phone_size=28, phone_color=(48, 48, 48),
    qr_size=90, qr_x=90, qr_y=1422, qr_crop=None,   # qr_crop=None 时自动去除白边
    saoma_y=1529, saoma_size=18, saoma_color=(70, 70, 70),
    vert_x=872, vert_y=336, vert_step=30, vert_size=18, vert_color=(181, 179, 174),
)


def detect_blue_anchor(base_path, ratio):
    """检测底板上最大的蓝色笔触，返回 (人物底边y, (y0, y1, ratio))"""
    a = np.array(Image.open(base_path).convert("RGB")).astype(int)
    r, g, b = a[:, :, 0], a[:, :, 1], a[:, :, 2]
    m = ((b > 110) & (b - r > 45) & (b - g > 40)).astype(np.uint8)
    n, lab, st, _ = cv2.connectedComponentsWithStats(m, 8)
    i = max(range(1, n), key=lambda k: st[k, cv2.CC_STAT_AREA])
    y0 = int(st[i, cv2.CC_STAT_TOP]); h = int(st[i, cv2.CC_STAT_HEIGHT])
    return int(round(y0 + h * ratio)), (y0, y0 + h, ratio)


def cutout(src_path):
    """抠图：RGB 取 matting 修正版，alpha 取干净形状+羽化。返回 RGBA Image"""
    from rembg import remove, new_session
    src = Image.open(src_path).convert("RGBA")
    sess = new_session("u2net_human_seg")
    rgb = np.array(remove(
        src, session=sess, alpha_matting=True,
        alpha_matting_foreground_threshold=240,
        alpha_matting_background_threshold=15,
        alpha_matting_erode_size=8))[:, :, :3].copy()
    a = np.array(remove(src, session=sess))[:, :, 3].astype(np.uint8)
    b = (a > 128).astype(np.uint8) * 255
    b = cv2.morphologyEx(b, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8))
    b = cv2.morphologyEx(b, cv2.MORPH_OPEN,  np.ones((3, 3), np.uint8))
    n, lab, st, _ = cv2.connectedComponentsWithStats((b > 0).astype(np.uint8), 8)
    if n > 1:
        i = 1 + int(np.argmax(st[1:, cv2.CC_STAT_AREA]))
        b = np.where(lab == i, 255, 0).astype(np.uint8)
    soft = cv2.GaussianBlur(b, (0, 0), 2.2)
    rgba = np.dstack([rgb, soft]).astype(np.uint8)
    img = Image.fromarray(rgba, "RGBA")
    al = rgba[:, :, 3]; ys, xs = np.where(al > 2)
    return img.crop((int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1))


def torn(person, anchor_y, white=(214, 213, 217), lift=2, seed=5, pad=6):
    """撕纸：在**锚定绝对位置**生成斜向平滑波浪 + 细密毛边。

    注意：anchor_y 是【人物小图内的局部坐标】（= 海报锚点 - 人物贴图左上角 y）。

    为什么要用绝对位置（关键教训）：
      母版撕纸线固定在 y=1352.8，与"人物底部在哪"无关。
      若改成"相对人物底部"，人物底部起伏一变，斜线就会被 bottom 截平 → 斜度消失。
      所以撕纸线 = anchor_y - lift + 斜率 + 波浪，与人物底部解耦。

    母版实测：斜率 k≈0.12(左上→右下)；撕纸线在锚点上方 12px；
              波形平滑圆润（窗口180低通）；白边 2.5~9px + 细纤维毛。
    """
    a = np.array(person).copy(); h, w = a.shape[:2]
    alpha = a[:, :, 3].copy()
    bottom = np.full(w, -1, int)
    for x in range(w):
        idx = np.where(alpha[:, x] > 10)[0]
        if len(idx): bottom[x] = int(idx.max())
    valid = bottom >= 0

    tpl = json.loads(TEAR_TPL.read_text(encoding="utf-8"))
    _top = np.array(tpl["top_y"], float)
    k, _ = np.polyfit(np.arange(len(_top)), _top, 1)      # ★ 母版斜率

    def lp(p, kk):
        return np.convolve(p, np.ones(kk) / kk, mode="same")

    tile = np.load(TEAR_PROF)
    seg = tile[37:37 + w]
    if len(seg) < w:
        seg = np.concatenate([seg, tile[:w - len(seg)]])
    big = lp(seg, 180); big = big - big.mean(); big = big / (np.std(big) + 1e-9) * 3.4
    rng = np.random.RandomState(seed)
    fine = lp(rng.randn(w), 5); fine = fine / (np.std(fine) + 1e-9) * 1.8
    prof = big + fine
    for i in range(1, w):
        d = prof[i] - prof[i - 1]
        if d > 2.2: prof[i] = prof[i - 1] + 2.2
        elif d < -2.2: prof[i] = prof[i - 1] - 2.2

    px = np.arange(w)
    y_line = (anchor_y - lift) + k * (px - w / 2) + prof   # ★ 绝对位置 + 斜率 + 波浪
    # ★ 不做截断：母版撕纸线本来就跨过人物底部（下方那截白边画在人物外）
    thick = 8.0 + np.clip((prof - prof.min()) / (np.ptp(prof) + 1e-9), 0, 1) * 8.0
    fib = (rng.rand(w) < 0.45).astype(float)

    H = h + pad
    out = np.zeros((H, w, 4), np.uint8); out[:h] = a
    paper = np.zeros((H, w, 4), np.uint8)
    for x in range(w):
        if not valid[x]: continue
        y0 = int(np.clip(y_line[x], 0, H - 1))
        t = max(4, int(round(thick[x] + fib[x] * 5.0)))
        paper[y0:min(H, y0 + t), x] = [white[0], white[1], white[2], 255]
        out[y0:, x, 3] = 0
    return Image.fromarray(out, "RGBA"), Image.fromarray(paper, "RGBA")


def char_img(ch, fontpath, target_h, color, sx=1.0, sy=1.0):
    f = ImageFont.truetype(str(fontpath), int(target_h * 2))
    d0 = ImageDraw.Draw(Image.new("RGB", (4, 4)))
    bb = d0.textbbox((0, 0), ch, font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(im).text((-bb[0], -bb[1]), ch, font=f, fill=color + (255,))
    s = target_h / h
    return im.resize((max(1, int(w * s * sx)), max(1, int(h * s * sy))), Image.LANCZOS)


def glyph_size(ch, em, color):
    """按字号 em 渲染单字，返回裁到墨迹的 RGBA（同一字号 = 各字真正等大）"""
    f = ImageFont.truetype(str(NAME_FONT), max(8, int(round(em))))
    d0 = ImageDraw.Draw(Image.new("RGB", (4, 4)))
    bb = d0.textbbox((0, 0), ch, font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    if w <= 0 or h <= 0:
        return None
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(im).text((-bb[0], -bb[1]), ch, font=f, fill=color + (255,))
    return im


def name_fonts(name):
    """姓名逐字字体：第1字吕建德行楷(贴近模板)，第2字龙藏，其余志莽行书；第2字横向压缩"""
    lv  = FONT_DIR / "方正字迹-吕建德行楷.ttf"
    lc  = FONT_DIR / "LongCang-Regular.ttf"
    zm  = FONT_DIR / "志莽行书-Regular.ttf"
    out = []
    for i, _ in enumerate(name):
        if i == 0:   out.append((lv, 1.0, 1.0))
        elif i == 1: out.append((lc, 0.75, 1.0))
        else:        out.append((zm, 0.78, 1.0))
    return [f for f, _, _ in out], [(a, b) for _, a, b in out]


def qr_image(qr_path, size, crop):
    q = Image.open(qr_path).convert("RGBA")
    if crop is None:                       # 自动去除四周白边
        a = np.array(q)
        nz = (a[:, :, :3].astype(int).sum(2) < 700) & (a[:, :, 3] > 40)
        ys, xs = np.where(nz)
        crop = (int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1)
    q = q.crop(crop).resize((size, size), Image.BOX)
    wq = Image.new("RGBA", (size, size), (255, 255, 255, 255))
    wq.alpha_composite(q)
    return wq




# ============================================================
# 纸条撕纸流程（与程序生成撕纸互斥，优先使用纸条）
# ============================================================
STRIP_IMG = SKILL_DIR / "assets" / "poster" / "tear_strip.png"   # 母版原版白边


def fuse_strip(person, strip, strip_y_ratio=None, verbose=True):
    """人像 + 纸条 融合（严格双向约束）。

    约束：
      ① 纸条【内容】必须先裁掉两侧空白，再拉满人像宽度
         —— 保证「人像不超出纸条、纸条不超出人像」
      ② 逐列按纸条上边缘裁切人像底部（撕裂）
      ③ 撕裂线在人像内的高度由 strip_y_ratio 控制

    返回 (一体图, 撕裂线在人像内的 y)
    """
    PW, PH = person.size

    # ① 纸条内容 -> 拉满人像宽（双向约束的核心）
    sa = np.array(strip)
    sal = sa[:, :, 3]
    cols = np.where((sal > 8).any(0))[0]
    if len(cols) == 0:
        raise RuntimeError("纸条没有有效内容")
    content = strip.crop((int(cols.min()), 0, int(cols.max()) + 1, strip.height))
    k = PW / content.width
    strip_f = content.resize((PW, max(1, int(round(content.height * k)))), Image.LANCZOS)

    st = np.array(strip_f)
    al = st[:, :, 3]
    tops = np.full(PW, -1, int)
    for x in range(PW):
        idx = np.where(al[:, x] > 8)[0]
        if len(idx): tops[x] = idx.min()
    valid = tops >= 0
    if valid.sum() < PW * 0.5:
        raise RuntimeError("纸条拉满后覆盖不足：%d/%d 列" % (valid.sum(), PW))
    med = int(np.median(tops[valid]))

    # ② 纵向位置：让撕裂线均值落在目标高度
    if strip_y_ratio is None:
        tear_local = int(round(PH * 0.90))       # 默认：人像底部上方 10%
    else:
        tear_local = int(round(PH * strip_y_ratio))
    strip_y = tear_local - med

    # ③ 撕裂人像
    pa = np.array(person).copy()
    for x in range(PW):
        cy = strip_y + (tops[x] if tops[x] >= 0 else med)
        ci = int(round(cy))
        if 0 <= ci < PH:
            pa[ci:, x, 3] = 0
    # 纸条贴到与人像同尺寸的画布上（超出人像高度范围的部分自然被裁掉）
    paper = Image.new("RGBA", (PW, PH), (0, 0, 0, 0))
    paper.alpha_composite(strip_f, (0, strip_y))

    # 双边约束：纸条（那层纸）只能出现在人像轮廓之内
    #   母版几何：纸条 = 「人像下半身被撕开的那一层纸」，轮廓贴着人像走。
    #   半身照在撕裂线处人像宽≈肩宽，做成固定宽度矩形也看不出来；
    #   全身照就会露馅：撕裂线以上人像更宽、纸条却是矩形，
    #   于是人像左右悬在纸条外面（或者反过来纸条超出人像）。
    #   要让「人像不超纸条」和「纸条不超人像」同时成立，唯一做法是
    #   把纸条 alpha 与人像轮廓相乘 —— 纸条轮廓 = 人像轮廓。
    sil = np.array(person)[:, :, 3].astype(np.float32) / 255.0
    _p = np.array(paper).astype(np.float32)
    _p[:, :, 3] *= sil
    paper = Image.fromarray(_p.astype(np.uint8), "RGBA")
    if verbose:
        vis = np.where(np.array(person)[tear_local, :, 3] > 8)[0]
        print("  [纸条] 纸条轮廓已贴合人像轮廓（撕裂线处人像宽 %dpx）"
              % ((int(vis.max()) - int(vis.min()) + 1) if len(vis) else 0))

    merged = Image.new("RGBA", (PW, PH), (0, 0, 0, 0))
    merged.alpha_composite(paper, (0, 0))
    merged.alpha_composite(Image.fromarray(pa, "RGBA"), (0, 0))

    if verbose:
        print("  [纸条] 内容 %dpx -> 拉满 %dpx；撕裂线(人像内) y=%d / %d"
              % (content.width, PW, tear_local, PH))
    return merged, tear_local


def build(person_path, name, cohort, out_path, qr_path=None, phone=None,
          ratio=None, verbose=True):
    ratio = L["person_ratio"] if ratio is None else ratio
    canvas = Image.open(BASE_IMG).convert("RGBA")
    W, H = canvas.size

    person = cutout(person_path)
    tw = L["person_w"]; th = int(round(person.height * tw / person.width))
    person = person.resize((tw, th), Image.LANCZOS)
    anchor_y, info = detect_blue_anchor(BASE_IMG, ratio)
    if verbose:
        print("  [锚定] 蓝笔触 y[%d,%d] 比例%s -> 人物底边 y=%d" % (info[0], info[1], info[2], anchor_y))
    px = L["person_right"] - tw
    py = anchor_y - th
    # ★ 优先走"纸条"流程；无纸条时退回程序生成撕纸
    if STRIP_IMG.exists() and not L.get("use_procedural_tear"):
        merged, tear_local = fuse_strip(person, Image.open(STRIP_IMG).convert("RGBA"),
                                        L.get("strip_tear_ratio"))
        # ★ 定位：用【撕裂线】对齐锚点（人物实际边缘，不是图块边界）
        scale = L["person_w"] / merged.width
        top_y = anchor_y - int(round(tear_local * scale))
        px = L["person_right"] - int(round(merged.width * scale))
        py = top_y
        merged_s = merged.resize((int(round(merged.width * scale)), int(round(merged.height * scale))), Image.LANCZOS)
        person, paper = merged_s, None
    else:
        anchor_local = anchor_y - py
        person, paper = torn(person, anchor_local, L["tear_white"], L["tear_lift"])
        px = L["person_right"] - tw
        py = anchor_y - th

    # 姓名右边界 = 人像在名字高度带内的最左边缘 - 安全间距
    band = L["name_band"]
    _lefts = []
    for _lay in ([person] if paper is None else [person, paper]):
        _a = np.array(_lay)[:, :, 3]
        _y0 = max(0, band[0] - py); _y1 = min(_a.shape[0], band[1] - py)
        if _y1 > _y0:
            _xs = np.where(_a[_y0:_y1].any(0))[0]
            if len(_xs):
                _lefts.append(int(_xs.min()) + px)
    name_right = (min(_lefts) - L["name_right_pad"]) if _lefts else W

    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    # 姓名：所有字同一字号；总宽超出可用区则整体等比缩小（永不压到人像）
    if name:
        S, gap = float(L["name_size0"]), L["name_gap"]
        avail = name_right - L["name_x"] + 1
        need = len(name) * S + (len(name) - 1) * gap
        if need > avail:
            S = max(10.0, (avail - (len(name) - 1) * gap) / float(len(name)))
        x = L["name_x"]
        for ch in name:
            gi = glyph_size(ch, S, L["name_color"])
            if gi is None:
                continue
            layer.alpha_composite(gi, (int(round(x)), int(L["name_top"])))
            x += gi.width + gap
        if verbose:
            print("  [姓名] 字号 %.1f  右边界 %d  可用宽 %d  实际右缘 %d"
                  % (S, name_right, avail, x - gap - 1))

    # 期数
    qf = ImageFont.truetype(SONG_FONT, L["qi_size"])
    bb = d.textbbox((0, 0), cohort, font=qf)
    d.text((L["qi_x"] - bb[0], L["qi_y"] - bb[1]), cohort, font=qf, fill=L["qi_color"])

    # 电话
    if phone:
        pf = ImageFont.truetype(SONG_FONT, L["phone_size"])
        bbp = d.textbbox((0, 0), phone, font=pf)
        d.text((L["phone_x"] - bbp[0], L["phone_y"] - bbp[1]), phone, font=pf, fill=L["phone_color"])

    # 二维码 + 扫码咨询
    if qr_path:
        wq = qr_image(qr_path, L["qr_size"], L["qr_crop"])
        sm = "扫码咨询"; sf = ImageFont.truetype(SONG_FONT, L["saoma_size"])
        bbs = d.textbbox((0, 0), sm, font=sf)
        d.text((L["qr_x"] + L["qr_size"] / 2 - (bbs[2] - bbs[0]) / 2 - bbs[0],
                L["saoma_y"] - bbs[1]), sm, font=sf, fill=L["saoma_color"])

    # 竖排文案
    vf = ImageFont.truetype(SONG_FONT, L["vert_size"])
    for i, ch in enumerate("让世界听见你的声音"):
        bbv = d.textbbox((0, 0), ch, font=vf)
        d.text((L["vert_x"] - (bbv[2] - bbv[0]) / 2 - bbv[0],
                L["vert_y"] + i * L["vert_step"] - bbv[1]), ch, font=vf, fill=L["vert_color"])

    canvas.alpha_composite(layer)
    if qr_path:
        canvas.alpha_composite(wq, (L["qr_x"], L["qr_y"]))
    if paper is not None:
        canvas.alpha_composite(paper, (px, py))
    canvas.alpha_composite(person, (px, py))
    canvas.convert("RGB").save(out_path)
    if verbose:
        print("  已输出 %s  (%dx%d)  人物 %dx%d @(%d,%d)" % (out_path, W, H, tw, th, px, py))


def main():
    ap = argparse.ArgumentParser(description="学员海报生成（固定底板 + 自动抠图 + 自动锚定）")
    ap.add_argument("person", help="学员照片路径")
    ap.add_argument("--name", required=True, help="学员姓名")
    ap.add_argument("--cohort", required=True, help='期数，如 "第35期学员"')
    ap.add_argument("--qr", help="二维码图片（每个老师不同，可选）")
    ap.add_argument("--phone", help="联系电话（可选）")
    ap.add_argument("--ratio", type=float, help="人物底边占蓝笔触高度比例，默认 1.0（对齐母版）")
    ap.add_argument("--name-size", type=float, help="姓名基准字号，默认 170")
    ap.add_argument("--name-gap", type=float, help="姓名字间距，默认 13")
    ap.add_argument("--name-x", type=float, help="姓名字墨迹左缘 x，默认 116")
    ap.add_argument("--out", required=True, help="输出 PNG 路径")
    a = ap.parse_args()
    if not BASE_IMG.exists():
        sys.exit("缺少底板：%s" % BASE_IMG)
    if a.name_size: L["name_size0"] = a.name_size
    if a.name_gap:  L["name_gap"]   = a.name_gap
    if a.name_x is not None: L["name_x"] = a.name_x
    build(a.person, a.name, a.cohort, a.out, a.qr, a.phone, a.ratio)


if __name__ == "__main__":
    main()
