# -*- coding: utf-8 -*-
"""一份数据文件 → 开场白 + 三个流程表版本（原地填模板，1:1 保留格式）。

用法：
    python build_flow.py 客户数据.json [--check]

设计目的：把「每对新人会变」的内容全部放进 JSON，脚本本身**不含任何客户信息**。
这样换一对新人 = 改一份数据文件 + 跑一条命令，不需要再写/改代码。

数据文件格式见 templates/客户数据模板.json。

固定规则：**交付一律 Word（.docx），不导 PDF**。生成 + 自检全绿即可交付。
依赖：python-docx；落款贴底需要 pywin32 + 本机装 Word。
"""
import argparse
import copy
import json
import os
import re
import subprocess
import sys

import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn

# 彩排版 BGM 的固定制表位（相对左边距，pt）。
# 模板里 5 条 BGM 都精确落在 x=407.2；左边距 71.2 → 制表位 = 336.0。
# 用固定制表位取代"手调的尾随空格 + N 个 Tab"，内容长短变化也能对齐。
BGM_TAB_PT = 336.0

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from fill_flow import _place_signature_at_bottom  # noqa: E402


# ---------------- 通用小工具 ----------------
def set_font(run, name="楷体", size=12, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = rPr.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), name)


def set_para_text(p, text):
    """把段落文字换成 text —— 复用第一个 run（保住它的字体格式）。"""
    runs = p.runs
    if not runs:
        return p.add_run(text)
    runs[0].text = text
    for r in runs[1:]:
        r._element.getparent().remove(r._element)
    return runs[0]


def set_cell_lines(cell, lines):
    if isinstance(lines, str):
        lines = [lines]
    while len(cell.paragraphs) < len(lines):
        cell.paragraphs[-1]._p.addnext(copy.deepcopy(cell.paragraphs[0]._p))
    while len(cell.paragraphs) > len(lines):
        p = cell.paragraphs[-1]._p
        p.getparent().remove(p)
    for p, line in zip(cell.paragraphs, lines):
        set_para_text(p, line)


def cells_of(row):
    out, seen = [], set()
    for c in row.cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append(c)
    return out


def save(doc, path):
    if os.path.exists(path):
        os.chmod(path, 0o666)
    doc.save(path)


def plain_name(s):
    """蒋 涛 & 姜 桦 -> 蒋涛 & 姜 桦（只去掉名字内部的空格，保留 & 两边）。"""
    return re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", s)


def set_subtitle(d, text):
    for p in d.paragraphs[:3]:
        if "拟定版" in p.text:
            p.runs[0].text = text
            for r in p.runs[1:]:
                r.text = ""
            return


def keep_sig(doc):
    ps = doc.paragraphs
    for i, para in enumerate(ps):
        t = para.text.strip()
        if "联系电话" in t:
            para.paragraph_format.keep_together = True
            para.paragraph_format.keep_with_next = True
            if i + 1 < len(ps):
                ps[i + 1].paragraph_format.keep_with_next = True
            if i - 1 >= 0 and ps[i - 1]._p.findall(".//" + qn("w:drawing")):
                ps[i - 1].paragraph_format.keep_with_next = True
        elif "品山主持工作室" in t:
            para.paragraph_format.keep_together = True


def strip_color(cell):
    """把单元格里所有 run 的字体颜色清掉（回到默认黑），
    避免模板首字的颜色外溢到整段新文字。红色重点随后由 colorize 精确加回。"""
    for p in cell.paragraphs:
        for r in p.runs:
            rPr = r._element.find(qn("w:rPr"))
            if rPr is None:
                continue
            c = rPr.find(qn("w:color"))
            if c is not None:
                rPr.remove(c)


def colorize(p, spans, color="FF0000"):
    """把 p 里指定区间的文字标成 color，其余保持黑色（重建 run）。"""
    text = p.text
    if not text:
        return
    base = p.runs[0]._element.find(qn("w:rPr"))
    rPr = copy.deepcopy(base) if base is not None else None
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    marks = [False] * len(text)
    for a, b in spans:
        for i in range(max(0, a), min(b, len(text))):
            marks[i] = True
    pos = 0
    while pos < len(text):
        cur = marks[pos]
        start = pos
        while pos < len(text) and marks[pos] == cur:
            pos += 1
        run = p.add_run(text[start:pos])
        el = run._element
        old = el.find(qn("w:rPr"))
        if old is not None:
            el.remove(old)
        if rPr is not None:
            rp = copy.deepcopy(rPr)
            col = rp.find(qn("w:color"))
            if col is not None:
                rp.remove(col)
            if cur:
                col = rp.makeelement(qn("w:color"), {})
                col.set(qn("w:val"), color)
                rf = rp.find(qn("w:rFonts"))
                if rf is not None:
                    rf.addnext(col)
                else:
                    rp.insert(0, col)
            el.insert(0, rp)


def fill_desc(cell, lines, bgm=None):
    """填"具体描述"：保留模板的红字重点 + 蓝色 BGM 格式。"""
    reds = []
    para_i = None
    rPr_bgm = None
    for pi, p in enumerate(cell.paragraphs):
        for r in p.runs:
            col = r.font.color.rgb if (r.font.color and r.font.color.type) else None
            t = r.text.strip()
            if t.startswith("BGM"):
                el = r._element.find(qn("w:rPr"))
                if el is not None:
                    para_i, rPr_bgm = pi, copy.deepcopy(el)
            elif t and col is not None and str(col) in ("FF0000", "EE0000"):
                reds.append(t)
    set_cell_lines(cell, lines)
    strip_color(cell)
    if reds:
        for p in cell.paragraphs:
            spans = []
            for ph in reds:
                i = p.text.find(ph)
                if i >= 0:
                    spans.append((i, i + len(ph)))
            spans.sort()
            merged = []
            for a, b in spans:
                if merged and a <= merged[-1][1]:
                    merged[-1] = (merged[-1][0], max(merged[-1][1], b))
                else:
                    merged.append((a, b))
            if merged:
                colorize(p, merged)
    if bgm and rPr_bgm is not None:
        p = cell.paragraphs[min(para_i, len(cell.paragraphs) - 1)]
        run = p.add_run("        " + bgm)
        el = run._element
        old = el.find(qn("w:rPr"))
        if old is not None:
            el.remove(old)
        el.insert(0, copy.deepcopy(rPr_bgm))


# ---------------- 三个版本 ----------------
def build_opening(data, outdir):
    d = docx.Document()
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.top_margin = s.bottom_margin = Cm(2.5)
    s.left_margin = s.right_margin = Cm(3.0)
    st = d.styles["Normal"]
    st.font.name = "楷体"
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "楷体")
    h = d.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(h.add_run("开场白 · " + plain_name(data["couple"])), size=16, bold=True)
    h.paragraph_format.space_after = Pt(18)
    for text in data["opening"]:
        p = d.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.5
        set_font(p.add_run(text), size=12)
    path = os.path.join(outdir, "开场白.docx")
    save(d, path)
    return path


def fill_head(t, data):
    rows = t.rows
    set_cell_lines(cells_of(rows[1])[1], data["couple"])
    set_cell_lines(cells_of(rows[2])[1], data["time"])
    set_cell_lines(cells_of(rows[3])[1], data["hotel"])
    set_cell_lines(cells_of(rows[4])[1], data["tables"])
    set_cell_lines(cells_of(rows[5])[1], data["people"])
    set_cell_lines(cells_of(rows[6])[2], data["new_prep"])
    set_cell_lines(cells_of(rows[7])[2], data["staff_prep"])
    set_cell_lines(cells_of(rows[8])[1], data["summary"])
    set_cell_lines(cells_of(rows[9])[1], data["music"])
    set_cell_lines(cells_of(rows[10])[1], data["mood"])
    set_cell_lines(cells_of(rows[11])[1], data["seats"])


def short_hotel(data):
    return data.get("short_hotel") or data["hotel"].split(" - ")[0].strip()


def build_base(data, outdir):
    d = docx.Document(data["tpl_base"])
    t = d.tables[0]
    fill_head(t, data)
    flow = data["flow"]
    need = len(flow) - 18
    if need > 0:
        src = t.rows[22]._tr
        for _ in range(need):
            src.addnext(copy.deepcopy(src))
    for k, item in enumerate(flow):
        row = t.rows[14 + k]
        cs = cells_of(row)
        set_cell_lines(cs[0], item["name"])
        bgm = data["bgm"] if item["name"] in data.get("bgm_rows", []) else None
        fill_desc(cs[1], item["desc"], bgm)
    keep_sig(d)
    set_subtitle(d, "（%s 于%s沟通后拟定版）" % (data["talk_date"], short_hotel(data)))
    path = os.path.join(outdir, "仪式流程表_基础版.docx")
    save(d, path)
    return path


def build_final(data, outdir):
    d = docx.Document(data["tpl_final"])
    t = d.tables[0]
    fill_head(t, data)
    for k, item in enumerate(data["flow"]):
        row = t.rows[14 + k]
        cs = cells_of(row)
        set_cell_lines(cs[0], item["name"])
        bgm = data["bgm"] if item["name"] in data.get("bgm_rows", []) else None
        fill_desc(cs[1], item["desc"], bgm)
    keep_sig(d)
    set_subtitle(d, "（%s 于%s彩排后拟定版）" % (data["wedding_date"], short_hotel(data)))
    path = os.path.join(outdir, "仪式流程表_最终版.docx")
    save(d, path)
    return path


def fill_reh_para(p, name, note, bgm):
    runs = p.runs
    name_run, bgm_run, note_runs = None, None, []
    for r in runs:
        sz = r.font.size.pt if r.font.size else None
        t = r.text.strip()
        if t.startswith("BGM"):
            bgm_run = r
        elif sz and sz >= 17 and name_run is None:
            name_run = r
        elif sz and sz <= 13 and t:
            note_runs.append(r)
    if name_run:
        name_run.text = name
    for r in runs:
        sz = r.font.size.pt if r.font.size else None
        if (r is not name_run and sz and sz >= 17 and r.text.strip()
                and not r.text.strip().startswith("BGM")):
            r.text = ""
    if note_runs:
        note_runs[0].text = (note or "").strip()
        for r in note_runs[1:]:
            r.text = ""

    # BGM：挂固定制表位 + 把填充用的空白 run 归一成「一个 Tab」
    if bgm:
        p.paragraph_format.tab_stops.add_tab_stop(Pt(BGM_TAB_PT), WD_TAB_ALIGNMENT.LEFT)
        ws = [r for r in p.runs if r.text and not r.text.strip()]
        if ws:
            ws[0].text = "\t"
            for r in ws[1:]:
                r.text = ""
        if bgm_run is not None:
            bgm_run.text = bgm if ws else ("\t" + bgm)
        else:
            r = p.add_run("\t" + bgm)
            set_font(r, size=12)
    elif bgm_run is not None:
        bgm_run.text = ""


def build_rehearsal(data, outdir):
    d = docx.Document(data["tpl_reh"])
    ps = d.paragraphs
    first = next(i for i, p in enumerate(ps) if p.text.strip().startswith("1、控场"))
    n = 0
    j = first
    while j < len(ps) and ps[j].text.strip() and not ps[j].text.strip().startswith("新人准备"):
        n += 1
        j += 1
    last = first + n - 1
    items = data["rehearsal"]
    for _ in range(max(0, len(items) - n)):
        ps[last]._p.addnext(copy.deepcopy(ps[last]._p))
    for _ in range(max(0, n - len(items))):
        p = d.paragraphs[first + len(items)]._p
        p.getparent().remove(p)
    ps = d.paragraphs
    for k, item in enumerate(items):
        fill_reh_para(ps[first + k], item["name"], item.get("note"), item.get("bgm"))
    # 表头三行
    for p in d.paragraphs:
        t = p.text
        if "Wedding婚礼流程" in t:
            p.runs[0].text = data["couple"] + " "
            for r in p.runs[1:]:
                r.text = ""
            p.runs[1].text = "Wedding婚礼流程"
        elif t.startswith("婚礼日期"):
            p.runs[0].text = "婚礼日期：" + data["wedding_date"]
            for r in p.runs[1:7]:
                r.text = ""
            p.runs[-1].text = data["time"]
        elif t.startswith("婚礼基调"):
            p.runs[0].text = "婚礼基调："
            p.runs[1].text = data.get("rehearsal_mood", data["mood"])
            for r in p.runs[2:4]:
                r.text = ""
            p.runs[-1].text = data["hotel"].replace(" - ", "-")
    # 准备清单
    for p in d.paragraphs:
        t = p.text.strip()
        if t.startswith("新人准备"):
            set_para_text(p, "新人准备：" + data.get("rehearsal_new_prep", data["new_prep"]))
        elif t.startswith("婚礼人准备"):
            set_para_text(p, "婚礼人准备：" + data.get("rehearsal_staff_prep", data["staff_prep"]))
    keep_sig(d)
    path = os.path.join(outdir, "婚礼流程_彩排简易版.docx")
    save(d, path)
    return path


# ---------------- 主流程 ----------------
def load(path):
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    need = ["out", "couple", "talk_date", "wedding_date", "time", "hotel",
            "opening", "flow", "rehearsal", "tpl_base", "tpl_final", "tpl_reh"]
    miss = [k for k in need if k not in data]
    if miss:
        raise SystemExit("数据文件缺字段：%s" % ", ".join(miss))
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data", help="客户数据 json")
    ap.add_argument("--check", action="store_true", help="生成后跑 check_flow.py 自检")
    ap.add_argument("--stale", default="", help="自检用的旧场次特征词（逗号分隔）")
    a = ap.parse_args()

    data = load(a.data)
    outdir = data["out"]
    os.makedirs(outdir, exist_ok=True)

    made = [build_opening(data, outdir), build_base(data, outdir),
            build_final(data, outdir), build_rehearsal(data, outdir)]
    # 落款贴底（需要 Word）
    for p in made[1:]:
        print("落款:", os.path.basename(p))
        _place_signature_at_bottom(p)
    for p in made:
        print("生成:", p)

    rc = 0
    if a.check:
        cmd = [sys.executable, os.path.join(HERE, "check_flow.py")] + made + ["--couple", data["couple"]]
        if a.stale:
            cmd += ["--stale", a.stale]
        print("\n--- 自检 ---")
        rc |= subprocess.call(cmd)
    if rc:
        print("\n注意：有环节失败（见上面输出）")
    return 1 if rc else 0


if __name__ == "__main__":
    sys.exit(main())
