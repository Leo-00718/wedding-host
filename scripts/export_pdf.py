# -*- coding: utf-8 -*-
"""docx → PDF —— 用 Word 渲染，分页与 Word 完全一致。

⚠️ 非交付物：**交付一律用 Word（.docx），不导 PDF**。
   本脚本只作内部预览 / 临时用途，不进交付流程。


为什么不用 Word 自带的「另存为 / 导出 PDF」：
    本机装的是第三方「office2016 精简版」，PDF 导出组件被删掉了，
    doc.SaveAs(PDF) / ExportAsFixedFormat 会直接报「此功能尚未安装」。
    这里改用 Word 的打印引擎 + 「Microsoft Print to PDF」虚拟打印机，
    渲染和分页仍然由 Word 完成 —— 实测页数与 Word 一致（2/2/1 对得上）。

用法：
    python export_pdf.py 成品.docx [更多.docx ...]
    python export_pdf.py 交付物文件夹              # 导出其中所有 .docx
    python export_pdf.py 交付物文件夹 --outdir 输出目录

依赖：pywin32；本机装 Word；「Microsoft Print to PDF」打印机（Win10/11 自带）。
注意：导出时会短暂弹出 Word 窗口（打印引擎要求文档窗口处于活动状态，隐藏会报错）。
"""
import argparse
import glob
import os
import sys
import time


def collect(paths):
    out = []
    for p in paths:
        if os.path.isdir(p):
            out += sorted(glob.glob(os.path.join(p, "*.docx")))
        elif os.path.isfile(p):
            out.append(p)
        else:
            print("[跳过] 找不到:", p)
    # 排除 Word 的临时锁定文件 ~$xxx.docx
    return [p for p in out if not os.path.basename(p).startswith("~$")]


def _remove_quiet(path):
    if os.path.exists(path):
        try:
            os.chmod(path, 0o666)
            os.remove(path)
            return True
        except Exception:
            return False
    return True


def _close_quiet(doc, tries=8):
    """Word 打印是异步的，忙的时候 Close 会报「调用被拒绝」，重试即可。"""
    for _ in range(tries):
        try:
            doc.Close(0)
            return
        except Exception:
            time.sleep(1.0)


def _wait_file(pdf, timeout=60):
    """等 PDF 真正落盘（打印是异步的）。"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        if os.path.exists(pdf) and os.path.getsize(pdf) > 0:
            time.sleep(1.0)          # 再等一会儿，确保写盘收尾
            return True
        time.sleep(0.5)
    return os.path.exists(pdf) and os.path.getsize(pdf) > 0


def export(files, outdir):
    import win32com.client as wc
    app = wc.DispatchEx("Word.Application")
    app.Visible = True          # 打印引擎要求文档窗口处于活动状态，隐藏会报错
    app.DisplayAlerts = 0
    try:
        app.ActivePrinter = "Microsoft Print to PDF"
    except Exception as e:
        print("[警告] 切不到 Microsoft Print to PDF：", e)
    ok = 0
    try:
        for docx in files:
            base = os.path.splitext(os.path.basename(docx))[0]
            pdf = os.path.join(outdir or os.path.dirname(os.path.abspath(docx)), base + ".pdf")
            # 关键：先删掉旧 PDF，否则会把上一版误判成这次导出的结果
            if not _remove_quiet(pdf):
                print("失败 %s：旧 PDF 被占用，无法覆盖" % base)
                continue
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True,
                                   AddToRecentFiles=False, Visible=True)
            try:
                # 官方版 Office 能原生导出就走原生；精简版会抛异常，退回打印
                try:
                    d.ExportAsFixedFormat(pdf, 17)      # 17 = wdExportFormatPDF
                except Exception:
                    d.PrintOut(Background=False, OutputFileName=pdf, Item=0, PrintToFile=True)
                if _wait_file(pdf):
                    print("OK   %s  (%d bytes)" % (base + ".pdf", os.path.getsize(pdf)))
                    ok += 1
                else:
                    print("失败 %s（没有生成 PDF）" % base)
            except Exception as e:
                print("失败 %s：%s" % (base, str(e)[:140]))
            finally:
                _close_quiet(d)
    finally:
        for _ in range(8):
            try:
                app.Quit()
                break
            except Exception:
                time.sleep(1.0)
    return ok


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+", help="docx 文件，或包含 docx 的文件夹")
    ap.add_argument("--outdir", default=None, help="PDF 输出目录（默认与源文件同目录）")
    a = ap.parse_args()
    files = collect(a.paths)
    if not files:
        print("没有找到 .docx")
        return 1
    if a.outdir:
        os.makedirs(a.outdir, exist_ok=True)
    n = export(files, a.outdir)
    print("完成：%d/%d 个导出成功" % (n, len(files)))
    return 0 if n == len(files) else 1


if __name__ == "__main__":
    sys.exit(main())
