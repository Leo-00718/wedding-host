# -*- coding: utf-8 -*-
"""生成“欢迎新学员”婚礼主持培训海报。

标准用法：
    python scripts/build_poster.py 人像.jpg --name "倪媛群" --out "欢迎海报.png"

可选素材：
    --logo 品牌 logo.png
    --qr 咨询二维码.png

设计原则：
    - 暖白纸感背景，左侧留白排版，右侧人像
    - 中文文字本地绘制，保证姓名和文案不会变成错别字
    - 有 OpenCV 时自动做人像抠图；没有时退回到圆角矩形人像卡片
    - 不把客户照片上传到网络
"""
import argparse
import os
import random
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont, ImageOps

W, H = 1080, 1800
CARD_WHITE = (255, 255, 255, 245)
INK = (28, 27, 25, 255)
GRAY = (105, 103, 97, 255)
LIGHT_GRAY = (190, 188, 183, 255)
BLUE = (25, 83, 160, 255)
RED = (176, 45, 52, 255)


def font_paths():
    return [
        r"C:\Windows\Fonts\simkai.ttf",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\Dengb.ttf",
        r"C:\Windows\Fonts\Deng.ttf",
    ]


def find_font(size, bold=True):
    paths = font_paths() if bold else font_paths()[1:]
    for path in paths:
        if os.path.exists(path):
            return ImageFont.truetype(path, size=size)
    raise RuntimeError("找不到可用中文字体，请安装微软雅黑、黑体或楷体")


def make_paper_background():
    """生成接近参考模板的暖白纸感底色。"""
    base = Image.new("RGB", (W, H), (247, 245, 239))
    rnd = random.Random(20261009)
    px = base.load()
    for y in range(H):
        for x in range(0, W, 2):
            if rnd.random() < 0.055:
                d = rnd.choice((-2, -1, 0, 1, 2))
                r, g, b = px[x, y]
                px[x, y] = (max(0, min(255, r + d)),
                            max(0, min(255, g + d)),
                            max(0, min(255, b + d)))
    return base.convert("RGBA")


def brush_stroke(draw, points, color, width, alpha=90):
    """画一条有手绘感的短线，用于蓝/红笔触。"""
    c = color[:3] + (alpha,)
    draw.line(points, fill=c, width=width, joint="curve")
    draw.line([(x + 2, y + 2) for x, y in points], fill=color[:3] + (max(10, alpha // 3),), width=max(1, width // 2), joint="curve")


def remove_portrait_background(photo):
    """优先用 GrabCut 去背景；缺 OpenCV 时返回原图。"""
    try:
        import cv2
        import numpy as np
    except Exception:
        return photo.convert("RGBA")

    img = cv2.cvtColor(np.array(photo.convert("RGB")), cv2.COLOR_RGB2BGR)
    h, w = img.shape[:2]
    mask = np.zeros((h, w), np.uint8)
    rect = (int(w * 0.08), 0, int(w * 0.84), int(h * 0.98))
    bgd = np.zeros((1, 65), np.float64)
    fgd = np.zeros((1, 65), np.float64)
    try:
        cv2.grabCut(img, mask, rect, bgd, fgd, 5, cv2.GC_INIT_WITH_RECT)
    except Exception:
        return photo.convert("RGBA")
    alpha = np.where((mask == cv2.GC_FGD) | (mask == cv2.GC_PR_FGD), 255, 0).astype(np.uint8)
    alpha = cv2.GaussianBlur(alpha, (7, 7), 0)
    rgba = cv2.cvtColor(img, cv2.COLOR_BGR2RGBA)
    rgba[:, :, 3] = alpha
    return Image.fromarray(rgba)


def trim_alpha(image):
    bbox = image.getbbox()
    return image.crop(bbox) if bbox else image


def fit_alpha(image, max_size):
    image = trim_alpha(image)
    image.thumbnail(max_size, Image.Resampling.LANCZOS)
    return image


def paste_dark_logo(canvas, logo_path, xy, max_width=210):
    if not logo_path or not os.path.exists(logo_path):
        return False
    logo = Image.open(logo_path).convert("RGB")
    if logo.width > max_width:
        h = int(logo.height * max_width / logo.width)
        logo = logo.resize((max_width, h), Image.Resampling.LANCZOS)
    gray = ImageOps.grayscale(logo)
    alpha = gray.point(lambda p: max(0, min(255, int((245 - p) * 1.8))))
    logo_rgba = logo.convert("RGBA")
    logo_rgba.putalpha(alpha)
    canvas.alpha_composite(logo_rgba, xy)
    return True


def paste_qr(canvas, qr_path, xy, size=112):
    if not qr_path or not os.path.exists(qr_path):
        return False
    qr = Image.open(qr_path).convert("RGB").resize((size, size), Image.Resampling.LANCZOS)
    box = Image.new("RGBA", (size + 16, size + 16), CARD_WHITE)
    box.alpha_composite(qr.convert("RGBA"), (8, 8))
    canvas.alpha_composite(box, (xy[0] - 8, xy[1] - 8))
    return True


def draw_placeholder_logo(draw):
    """没有品牌素材时画一个极简占位标记。"""
    draw.line((74, 74, 122, 74), fill=INK, width=5)
    draw.line((74, 74, 74, 122), fill=INK, width=5)
    draw.line((98, 74, 98, 110), fill=INK, width=5)
    draw.line((98, 110, 137, 110), fill=INK, width=5)
    draw.text((74, 145), "PASSION · 品言", font=find_font(27, bold=False), fill=INK)


def build_poster(args):
    photo = Image.open(args.portrait).convert("RGB")
    canvas = make_paper_background()
    draw = ImageDraw.Draw(canvas, "RGBA")

    # 纸面细节：参考模板中的红/蓝手绘笔触，只放在人像后方和左下。
    brush_stroke(draw, [(425, 1285), (505, 1265), (588, 1280), (672, 1262), (750, 1278)],
                 BLUE, 18, alpha=38)
    brush_stroke(draw, [(445, 1320), (530, 1305), (615, 1320), (700, 1308)],
                 RED, 8, alpha=20)

    # 右下人像：自动抠图后只保留上半身，贴近参考模板的留白比例。
    if args.no_cutout:
        portrait = photo.convert("RGBA")
    else:
        portrait = remove_portrait_background(photo)
    portrait = trim_alpha(portrait)
    if portrait.height > 0:
        portrait = portrait.crop((0, 0, portrait.width, int(portrait.height * 0.78)))
    portrait = fit_alpha(portrait, (590, 820))
    px = 1080 - portrait.width - 28
    py = 1800 - portrait.height - 30
    # 人像下方轻微阴影，避免贴在纸面上显得悬浮。
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.ellipse((px + 40, py + portrait.height - 18, px + portrait.width - 10, py + portrait.height + 18),
               fill=(35, 32, 28, 38))
    shadow = shadow.filter(ImageFilter.GaussianBlur(16))
    canvas.alpha_composite(shadow)
    canvas.alpha_composite(portrait, (px, py))

    # 顶部品牌。
    if not paste_dark_logo(canvas, args.logo, (62, 58), 230):
        draw_placeholder_logo(draw)

    # 标题区。
    f_title = find_font(46, bold=False)
    f_subtitle = find_font(22, bold=False)
    f_slogan = find_font(22, bold=False)
    draw.text((104, 380), args.title, font=f_title, fill=INK)
    draw.text((105, 447), args.subtitle, font=f_subtitle, fill=GRAY)

    # 右侧竖排口号。
    for i, ch in enumerate(args.slogan):
        draw.text((944, 390 + i * 34), ch, font=f_slogan, fill=(88, 86, 80, 230))

    # 左侧欢迎与姓名。
    f_welcome = find_font(82, bold=False)
    f_name = find_font(112, bold=True)
    f_cohort = find_font(31, bold=False)
    f_tagline = find_font(32, bold=False)

    draw.text((155, 690), "欢迎", font=f_welcome, fill=(196, 194, 189, 210))
    draw.text((96, 780), args.name, font=f_name, fill=INK)
    draw.text((108, 955), args.cohort, font=f_cohort, fill=(53, 52, 49, 255))

    draw.line((140, 1265, 230, 1265), fill=BLUE, width=3)
    draw.text((140, 1300), args.tagline, font=f_tagline, fill=(49, 48, 45, 255))

    # 底部联系方式。
    f_contact = find_font(24, bold=False)
    f_address = find_font(19, bold=False)
    if not paste_qr(canvas, args.qr, (72, 1560), 112):
        draw.rounded_rectangle((72, 1560, 190, 1678), radius=8, outline=(70, 68, 64, 80), width=2)
        draw.text((96, 1605), "QR", font=find_font(22, bold=True), fill=GRAY)
    draw.text((214, 1570), "电话：" + args.phone, font=f_contact, fill=(55, 54, 51, 255))
    draw.text((214, 1628), "地址：" + args.address1, font=f_address, fill=(55, 54, 51, 255))
    draw.text((214, 1658), args.address2, font=f_address, fill=(55, 54, 51, 255))
    if args.qr:
        draw.text((77, 1682), "扫码咨询", font=find_font(17, bold=False), fill=GRAY)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    canvas.convert("RGB").save(out, quality=96, subsampling=0)
    return out


def main():
    ap = argparse.ArgumentParser(description="生成婚礼主持培训欢迎海报")
    ap.add_argument("portrait", help="人像照片路径")
    ap.add_argument("--name", required=True, help="要展示的姓名")
    ap.add_argument("--out", required=True, help="输出 PNG/JPG 路径")
    ap.add_argument("--title", default="品言 · 语言艺术", help="海报标题")
    ap.add_argument("--subtitle", default="婚礼主持 · 语言表达 · 口才训练", help="标题下副标题")
    ap.add_argument("--slogan", default="让世界听见你的声音", help="右侧竖排口号")
    ap.add_argument("--cohort", default="第33期学员", help="姓名下方期数")
    ap.add_argument("--tagline", default="在场，即主角", help="左下短句")
    ap.add_argument("--phone", default="156 9689 7507", help="页脚电话")
    ap.add_argument("--address1", default="重庆市两江新区五里店街道", help="页脚地址第一行")
    ap.add_argument("--address2", default="天澜大道11号星耀天地29栋17-1", help="页脚地址第二行")
    ap.add_argument("--logo", default="", help="可选品牌 logo 图片")
    ap.add_argument("--qr", default="", help="可选咨询二维码图片")
    ap.add_argument("--no-cutout", action="store_true", help="跳过自动抠图，使用原图")
    args = ap.parse_args()

    try:
        out = build_poster(args)
    except Exception as exc:
        print("生成失败：%s" % exc)
        return 1
    print("生成成功：%s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
