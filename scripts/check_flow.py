# -*- coding: utf-8 -*-
"""流程表成品自检 —— 一条命令跑完所有检查。

用法：
    python check_flow.py <成品.docx> [更多成品.docx ...]
                         [--couple "新 郎 & 新 娘"]
                         [--stale "上一场新人名,旧酒店名"]
                         [--pages 2]

检查项：
  1. 旧场次残留（--stale 给特征词：上一场的新人 / 酒店 / 日期…）
  2. 新人姓名是否真的填进了成品（--couple）
  3. 图片数量（logo、二维码、水印不能被弄丢）
  4. 红字重点、蓝色 BGM 是否还在
  5. 落款祝福语是否用对（表格版「幸福恒久」/ 彩排版「生活愉快」）
  6. Word 真实分页：页数、落款整块是否同页、是否贴页底、正文有没有溢出

退出码：全部通过 0，有 FAIL 1。
依赖：python-docx；分页部分需要本机装 Word（pywin32），没有就跳过并提示。
"""
import argparse, os, sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import docx
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml.ns import qn

def _graphics(el):
    """数一个 XML 元素下的图片对象：drawing + pict（VML）。"""
    return (len(el.findall(".//" + qn("w:drawing")))
            + len(el.findall(".//" + qn("w:pict"))))


RED = ("FF0000", "EE0000")
BLUE = "00B0F0"
BOTTOM_TOLERANCE = 60.0   # 署名离正文底超过这么多 pt 就算"没贴底"
BGM_TAB_PT = 336.0        # 彩排版 BGM 固定左制表位（距左边距；页面 x≈407.2pt）
BGM_TAB_TOLERANCE = 0.5


def iter_units(path):
    """所有可见文本 -> [(位置, 文本)]"""
    d = docx.Document(path)
    units = []
    for i, p in enumerate(d.paragraphs):
        if p.text.strip():
            units.append(("段落%d" % i, p.text.strip()))
    for ti, t in enumerate(d.tables):
        for ri, row in enumerate(t.rows):
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                for p in c.paragraphs:
                    if p.text.strip():
                        units.append(("表%d 第%d行" % (ti, ri), p.text.strip()))
    for s in d.sections:
        for p in list(s.header.paragraphs) + list(s.footer.paragraphs):
            if p.text.strip():
                units.append(("页眉页脚", p.text.strip()))
    return units


def all_paragraphs(path):
    d = docx.Document(path)
    out = list(d.paragraphs)
    for t in d.tables:
        for row in t.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                out.extend(c.paragraphs)
    return out


def colour_stats(path):
    reds, blues = [], []
    for p in all_paragraphs(path):
        for r in p.runs:
            if not r.text.strip():
                continue
            col = r.font.color.rgb if (r.font.color and r.font.color.type) else None
            if col is None:
                continue
            v = str(col)
            if v in RED:
                reds.append(r.text.strip())
            elif v == BLUE:
                blues.append(r.text.strip())
    return reds, blues


def bgm_tab_stats(path):
    """检查正文层 BGM 段落是否落在固定左制表位 336pt。

    只检查正文段落（彩排版 BGM 在正文层）；基础版/最终版的 BGM 在表格单元格里，
    不受这个制表位规则影响。
    """
    d = docx.Document(path)
    out = []
    for i, p in enumerate(d.paragraphs):
        if "BGM" not in p.text:
            continue
        left_stop = False
        for tab in p.paragraph_format.tab_stops:
            if tab.position is None:
                continue
            if (abs(tab.position.pt - BGM_TAB_PT) <= BGM_TAB_TOLERANCE
                    and int(tab.alignment or 0) == int(WD_TAB_ALIGNMENT.LEFT)):
                left_stop = True
                break
        out.append(dict(index=i, text=p.text.strip(),
                        has_tab="\t" in p.text, ok=left_stop and "\t" in p.text))
    return out


def get_word():
    """开一个 Word 实例给整轮自检复用（每次开/关 Word 要十几秒，非常慢）。"""
    try:
        import win32com.client as wc
    except ImportError:
        return None
    try:
        app = wc.DispatchEx("Word.Application")
    except Exception:
        return None
    app.Visible = False
    app.DisplayAlerts = 0
    return app


def word_pages(path, app, collect_bgm=False):
    """用 Word 量真实分页；app 为 None（没装 Word）返回 None

    collect_bgm 只在彩排版正文确有 BGM 固定制表位时开启；
    基础版/最终版的 BGM 在表格里，不做正文制表位扫描。
    """
    if app is None:
        return None
    doc = app.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
    try:
        total = doc.Paragraphs.Count
        logo = phone = sign = None
        for i in range(total, 0, -1):
            r = doc.Paragraphs(i).Range
            t = r.Text.replace("\r", "").strip()
            if sign is None and "品山主持工作室" in t:
                sign = (r.Information(3), r.Information(6))
            if phone is None and "联系电话" in t:
                phone = (r.Information(3), r.Information(6))
            if logo is None and r.InlineShapes.Count and i > total * 0.4:
                logo = (r.Information(3), r.Information(6))
        # 正文末段（落款之前最后一个有字的段落）
        content = None
        if logo:
            for i in range(1, total + 1):
                r = doc.Paragraphs(i).Range
                t = r.Text.replace("\r", "").strip()
                if t and "联系电话" not in t and "品山主持工作室" not in t:
                    if r.Information(6) < logo[1] or r.Information(3) < logo[0]:
                        content = (r.Information(3), r.Information(6))
        bgm_positions = []
        if collect_bgm:
            # 逐段取 BGM 的段内位置，避免 Word Find 在不同文档状态下漏掉结果。
            for i in range(1, total + 1):
                r = doc.Paragraphs(i).Range
                t = r.Text.replace("\r", "")
                if "BGM" not in t:
                    continue
                idx = t.index("BGM")
                bgm_range = doc.Range(r.Start + idx, r.Start + idx + 3)
                bgm_positions.append(round(bgm_range.Information(5), 1))
        return dict(pages=doc.ComputeStatistics(2), logo=logo, phone=phone,
                    sign=sign, content=content, bgm=bgm_positions,
                    left_margin=doc.PageSetup.LeftMargin,
                    bottom=doc.PageSetup.PageHeight - doc.PageSetup.BottomMargin)
    finally:
        doc.Close(0)


def check(path, couple, stale, expect_pages, word_app=None):
    name = os.path.basename(path)
    fails, warns, notes = [], [], []

    units = iter_units(path)
    text_all = "\n".join(t for _, t in units)

    if stale:
        for loc, txt in units:
            for w in stale:
                if w and w in txt:
                    fails.append("旧场次残留 [%s] %s" % (w, txt[:60]))
    def _norm(x):
        return x.replace(" ", "").replace("\u3000", "")
    if couple and _norm(couple) not in _norm(text_all):
        fails.append("没找到新人姓名「%s」" % couple)

    d = docx.Document(path)
    nimg = _graphics(d.element.body)
    for s in d.sections:
        for part in (s.header, s.footer):
            nimg += _graphics(part._element)
    # 流程表一定带 logo / 落款；没有这些标记的是普通文档（如开场白），跳过模板结构检查
    is_flow = ("品山主持工作室" in text_all) or nimg >= 3
    bgm_tabs = []

    if is_flow:
        if nimg < 3:
            fails.append("图形只剩 %d 个（模板共 3 个：正文 logo、水印 + 页眉图，缺一不可）" % nimg)
        else:
            notes.append("图形 %d 个（正文 + 页眉）" % nimg)

        reds, blues = colour_stats(path)
        bgm_tabs = bgm_tab_stats(path)
        if bgm_tabs:
            bad_tabs = [x for x in bgm_tabs if not x["ok"]]
            if bad_tabs:
                fails.append("BGM 固定制表位异常：%d/%d 条不在 %.1fpt（%s）" % (
                    len(bad_tabs), len(bgm_tabs), BGM_TAB_PT,
                    " / ".join(x["text"][:24] for x in bad_tabs[:3])))
            else:
                notes.append("BGM 固定制表位 %.1fpt（%d 条）" % (BGM_TAB_PT, len(bgm_tabs)))
        if not reds:
            warns.append("没有红字重点了（模板里是有的，确认是不是被冲掉）")
        else:
            notes.append("红字重点 %d 处：%s" % (len(reds), " / ".join(reds[:4])))
        if not blues:
            warns.append("没有蓝色 BGM（模板里是有的，确认是不是被冲掉）")
        else:
            notes.append("蓝色 BGM %d 处" % len(blues))

        if "品山主持工作室" not in text_all:
            fails.append("找不到落款署名")
        elif "彩排" in name or "简易" in name:
            if "生活愉快" not in text_all:
                fails.append("彩排版祝福语应该是「生活愉快」")
        else:
            if "幸福恒久" not in text_all:
                fails.append("表格版祝福语应该是「幸福恒久」")
    else:
        notes.append("普通文档（无 logo / 落款），跳过模板结构检查")

    w = word_pages(path, word_app, collect_bgm=is_flow and bool(bgm_tabs))
    if w is None:
        warns.append("没装 Word 或 pywin32，跳过真实分页检查")
    else:
        notes.append("Word 页数 %d" % w["pages"])
        if is_flow and bgm_tabs:
            positions = w.get("bgm", [])
            left_margin = w.get("left_margin", 0)
            expected_x = left_margin + BGM_TAB_PT
            # Word 不同版本/状态下，Information(5) 可能返回制表位坐标（336）
            # 或页面坐标（407.2）；统一折算成页面坐标再判断。
            normalized = [
                x + left_margin if abs(x - BGM_TAB_PT) <= 1.0 else x
                for x in positions
            ]
            bad_pos = [x for x in normalized if abs(x - expected_x) > 1.0]
            if len(positions) != len(bgm_tabs):
                fails.append("BGM 实际定位数 = %d，结构检查 = %d" % (len(positions), len(bgm_tabs)))
            elif bad_pos:
                fails.append("BGM 实际位置异常：%s（应为 %.1fpt）" % (bad_pos, expected_x))
            else:
                notes.append("BGM 实际位置 %.1fpt（页面坐标，%d 条）" % (expected_x, len(positions)))
        if expect_pages and w["pages"] != expect_pages:
            fails.append("页数是 %d，期望 %d" % (w["pages"], expect_pages))
        pages = [x[0] for x in (w["logo"], w["phone"], w["sign"]) if x]
        if not is_flow:
            pass  # 普通文档不做落款分页判断
        elif len(pages) == 3:
            if len(set(pages)) != 1:
                fails.append("落款被拆页了：logo 第%s页 / 电话 第%s页 / 署名 第%s页" % tuple(pages))
            else:
                notes.append("落款整块在第 %d 页" % pages[0])
                if w["content"] and (w["content"][0] != pages[0] or w["content"][1] > w["logo"][1]):
                    fails.append("落款和正文不在同一页（正文末页第 %s 页）" % w["content"][0])
                gap = w["bottom"] - w["sign"][1]
                if gap < 0:
                    fails.append("落款超出正文区 %.1fpt" % -gap)
                elif gap > BOTTOM_TOLERANCE:
                    warns.append("落款离页底还有 %.1fpt，没贴到底" % gap)
                else:
                    notes.append("落款贴页底（距正文底 %.1fpt）" % gap)
        else:
            warns.append("没找到完整落款块，跳过分页判断")

    print("=" * 78)
    print("%s   %s" % ("FAIL" if fails else "PASS", name))
    for x in notes:
        print("   . %s" % x)
    for x in warns:
        print("   ! %s" % x)
    for x in fails:
        print("   X %s" % x)
    return not fails


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--couple", default=None, help="新人姓名，例如 \"新 郎 & 新 娘\"")
    ap.add_argument("--stale", default=None, help="旧场次特征词，逗号分隔")
    ap.add_argument("--pages", type=int, default=None, help="期望页数（对所有文件生效）")
    a = ap.parse_args()
    stale = [x for x in (a.stale or "").replace("，", ",").split(",") if x.strip()]
    ok = True
    word = get_word()          # 整轮只开一次 Word
    try:
        for f in a.files:
            if not os.path.exists(f):
                print("FAIL 找不到文件:", f)
                ok = False
                continue
            try:
                ok = check(f, a.couple, stale, a.pages, word) and ok
            except Exception as e:
                print("FAIL %s 检查出错: %s" % (os.path.basename(f), e))
                ok = False
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
    print("=" * 78)
    print("全部通过" if ok else "有 FAIL，需要处理")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())