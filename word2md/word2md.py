# -*- coding: utf-8 -*-
"""
word2md.py —— 将 Word 文档(.doc / .docx / .rtf)稳定转换为 Markdown。

原理（为什么用这套方案）：
    本机未安装 pandoc / LibreOffice，但安装了 Microsoft Word(Office 16)。
    Word 自带最忠实的 .doc/.docx 解析引擎，因此用 Word 的 COM 自动化接口
    (pywin32) 打开文档并另存为「筛选过的 HTML」，再用 markdownify 把 HTML
    转成 Markdown。这是当前机器上保真度最高、对 .doc 老格式支持最好的路径。

依赖：
    pip install pywin32 markdownify

用法：
    python word2md.py 文件1.docx 文件2.doc ...         # 转换指定文件
    python word2md.py --dir 文件夹                      # 转换文件夹内所有 Word 文件
    python word2md.py --dir 文件夹 --out 输出目录 -r    # 递归 + 指定输出目录
"""

import os
import re
import sys
import shutil
import tempfile

try:
    import win32com.client
except ImportError:
    win32com.client = None

try:
    from markdownify import markdownify as _md
except ImportError:
    _md = None

# Word 保存格式常量
WD_FORMAT_FILTERED_HTML = 10   # 筛选过的 HTML（格式干净，体积小）
WD_DO_NOT_SAVE_CHANGES = 0     # 关闭时不保存更改

WORD_EXTS = (".doc", ".docx", ".rtf")


def _cleanup_html(html):
    """预处理 Word 输出的 HTML，去掉 Word 专用空段落等冗余标签。"""
    html = re.sub(r"<o:p>\s*</o:p>", "", html)
    html = re.sub(r"</?o:p>", "", html)
    return html


def _cleanup_md(md):
    """压缩多余空行、统一换行，保持 Markdown 整洁。"""
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    md = re.sub(r"[ \t]+\n", "\n", md)
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip() + "\n"


def _read_html_text(path):
    """读取 Word 导出的 HTML，自动识别编码（Word 中文版常输出 GBK）。"""
    data = open(path, "rb").read()
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", errors="replace")
    if data.startswith(b"\xff\xfe"):
        return data.decode("utf-16-le", errors="replace")
    if data.startswith(b"\xfe\xff"):
        return data.decode("utf-16-be", errors="replace")
    # 优先按 <meta charset=...> 声明解码
    m = re.search(rb"charset\s*=\s*[\"']?([A-Za-z0-9_\-]+)", data[:4000])
    if m:
        enc = m.group(1).decode("ascii", "ignore")
        try:
            return data.decode(enc, errors="replace")
        except (LookupError, UnicodeDecodeError):
            pass
    # 回退：严格尝试 utf-8，再中文编码
    for enc in ("utf-8", "gb18030", "gbk", "big5"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _start_word():
    """启动一个隐藏的 Word 实例。"""
    if win32com.client is None:
        raise RuntimeError("缺少依赖 pywin32，请先安装：pip install pywin32")
    word = win32com.client.Dispatch("Word.Application")
    try:
        word.Visible = False
        word.DisplayAlerts = 0            # wdAlertsNone：不弹任何对话框
        word.Options.ConfirmConversions = False
        word.DefaultWebOptions.Encoding = 65001   # msoEncodingUTF8：尽量让 HTML 输出为 UTF-8
    except Exception:
        pass
    return word


def _html_to_md(html):
    if _md is None:
        raise RuntimeError("缺少依赖 markdownify，请先安装：pip install markdownify")
    return _md(html, heading_style="ATX", bullets="-")


def convert_one(word, src, out_dir):
    """转换单个文件。返回 (ok, message, md_path)。"""
    src = os.path.abspath(src)
    stem = os.path.splitext(os.path.basename(src))[0]
    md_path = os.path.join(out_dir, stem + ".md")

    if not os.path.isfile(src):
        return False, f"文件不存在：{src}", None

    tmp = tempfile.mkdtemp(prefix="word2md_")
    html_path = os.path.join(tmp, stem + ".htm")
    files_dir = os.path.join(tmp, stem + "_files")

    doc = None
    try:
        doc = word.Documents.Open(
            FileName=src,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        # 另存为筛选过的 HTML（FileFormat=10），图片自动导出到 stem_files/
        doc.SaveAs2(FileName=html_path, FileFormat=WD_FORMAT_FILTERED_HTML)
        doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        doc = None

        if not os.path.isfile(html_path):
            return False, f"转换失败：未能生成 HTML（{os.path.basename(src)}）", None

        html = _read_html_text(html_path)
        html = _cleanup_html(html)
        md = _html_to_md(html)
        md = _cleanup_md(md)

        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md)

        # 复制图片文件夹（若有），保证 Markdown 里相对引用的图片可显示
        if os.path.isdir(files_dir):
            dst = os.path.join(out_dir, stem + "_files")
            if os.path.isdir(dst):
                shutil.rmtree(dst)
            shutil.copytree(files_dir, dst)

        return True, md_path, md_path

    except Exception as e:
        return False, f"转换失败 {os.path.basename(src)}：{e}", None
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        shutil.rmtree(tmp, ignore_errors=True)


def convert_files(src_list, out_dir=None, progress=None):
    """批量转换。

    参数：
        src_list : 待转换的 Word 文件路径列表
        out_dir  : 输出目录；None 表示输出到各源文件同目录
        progress : 可选回调 progress(done, total, filename)

    返回：
        [(ok, message, md_path), ...]
    """
    if win32com.client is None or _md is None:
        raise RuntimeError(
            "缺少依赖，请先执行：pip install pywin32 markdownify"
        )

    word = None
    try:
        word = _start_word()
    except Exception as e:
        return [(False, f"无法启动 Word：{e}", None) for _ in src_list]

    results = []
    total = len(src_list)
    try:
        for i, src in enumerate(src_list, 1):
            if progress:
                progress(i, total, os.path.basename(src))
            out = (os.path.abspath(out_dir)
                   if out_dir else os.path.dirname(os.path.abspath(src)))
            os.makedirs(out, exist_ok=True)
            results.append(convert_one(word, src, out))
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    return results


def _collect_files(paths, recursive):
    """根据输入路径（文件/文件夹）收集待转换的 Word 文件。"""
    files = []
    for p in paths:
        if os.path.isdir(p):
            if recursive:
                for root, _, names in os.walk(p):
                    for n in names:
                        if n.lower().endswith(WORD_EXTS):
                            files.append(os.path.join(root, n))
            else:
                for n in os.listdir(p):
                    fp = os.path.join(p, n)
                    if os.path.isfile(fp) and n.lower().endswith(WORD_EXTS):
                        files.append(fp)
        elif os.path.isfile(p) and p.lower().endswith(WORD_EXTS):
            files.append(p)
    return files


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    import argparse
    ap = argparse.ArgumentParser(description="Word(.doc/.docx/.rtf) → Markdown 转换器")
    ap.add_argument("paths", nargs="*", help="要转换的 Word 文件或文件夹")
    ap.add_argument("--dir", help="转换文件夹内所有 Word 文件")
    ap.add_argument("--out", help="输出目录（默认：源文件同目录）")
    ap.add_argument("-r", "--recursive", action="store_true", help="递归处理子文件夹")
    args = ap.parse_args(argv)

    inputs = list(args.paths)
    if args.dir:
        inputs.append(args.dir)
    if not inputs:
        ap.print_help()
        return 1

    files = _collect_files(inputs, args.recursive)
    if not files:
        print("未找到任何 Word 文件(.doc/.docx/.rtf)。")
        return 1

    def progress(done, total, name):
        print(f"[{done}/{total}] 正在转换：{name}")

    results = convert_files(files, out_dir=args.out, progress=progress)

    ok = sum(1 for r in results if r[0])
    print("-" * 40)
    for ok_flag, msg, _ in results:
        print(("OK  " if ok_flag else "FAIL") + "  " + msg)
    print(f"完成：成功 {ok}/{len(results)}")
    return 0 if ok == len(results) else 2


if __name__ == "__main__":
    sys.exit(main())
