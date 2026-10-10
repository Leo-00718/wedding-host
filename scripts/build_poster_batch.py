# -*- coding: utf-8 -*-
"""按「老师清单」批量出图 —— 同一张学员海报 × N 个老师（各自二维码 + 电话）

用法：
    python scripts/build_poster_batch.py 学员照.jpg --name "张文" --cohort "第34期学员" --teachers "老师清单.csv" --outdir "出图"

老师清单 CSV（3 列：老师、二维码、电话；可带表头；Excel 另存为 CSV 即可）：
    老师,二维码,电话
    王老师,D:\二维码\王.png,138 0000 0001
    李老师,D:\二维码\李.png,138 0000 0002

输出：{学员姓名}-{老师}.png，每个老师一张，二维码与电话一一对应。
人物抠图只做一次（N 张共享），批量比逐张快得多。
"""
import argparse
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_poster as bp


def read_teachers(path):
    """读老师清单；自动跳过表头，兼容 UTF-8 / GBK（Excel 两种另存为都支持）"""
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            rows = []
            first = True
            with open(path, encoding=enc, newline="") as f:
                for r in csv.reader(f):
                    r = [(c or "").strip() for c in r]
                    if len(r) < 3 or not r[0]:
                        continue
                    if first:
                        first = False
                        # 只对第一行判表头（否则路径里带"二维码"的数据行会被误杀）
                        if any(k in r[1] for k in ("二维码", "qr", "QR", "路径", "图片")) or "电话" in r[2]:
                            continue
                    rows.append((r[0], r[1], r[2]))
            return rows, enc
        except UnicodeDecodeError:
            continue
    sys.exit("老师清单读不出来（试过 UTF-8 / GBK）：%s" % path)


def main():
    ap = argparse.ArgumentParser(description="按老师清单批量出海报")
    ap.add_argument("person", help="学员照片路径")
    ap.add_argument("--name", required=True, help="学员姓名")
    ap.add_argument("--cohort", required=True, help="期数，例如 第34期学员")
    ap.add_argument("--teachers", required=True, help="老师清单 CSV（老师,二维码,电话）")
    ap.add_argument("--outdir", required=True, help="输出目录")
    ap.add_argument("--ratio", type=float, help="人物底边占蓝笔触高度比例，默认 1.0")
    a = ap.parse_args()

    teachers, enc = read_teachers(a.teachers)
    if not teachers:
        sys.exit("老师清单为空或格式不对（需要 3 列：老师,二维码,电话）")
    outdir = Path(a.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    print("老师清单：%d 个老师（编码 %s）" % (len(teachers), enc))
    print("人物抠图只做一次，%d 张共享…" % len(teachers))
    cut = bp.cutout(a.person)

    ok = 0
    for i, (teacher, qr, phone) in enumerate(teachers, 1):
        if not Path(qr).exists():
            print("  [%d/%d] %-8s 跳过：二维码文件不存在 %s" % (i, len(teachers), teacher, qr))
            continue
        out = outdir / ("%s-%s.png" % (a.name, teacher))
        bp.build(None, a.name, a.cohort, str(out), qr, phone, a.ratio, verbose=False, person_cut=cut)
        ok += 1
        print("  [%d/%d] %-8s 电话 %-16s -> %s" % (i, len(teachers), teacher, phone, out.name))
    print("")
    print("完成：%d/%d 张，输出目录 %s" % (ok, len(teachers), outdir))


if __name__ == "__main__":
    main()
