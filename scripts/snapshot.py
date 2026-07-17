# -*- coding: utf-8 -*-
"""snapshot.py — 把 pptx 每页导出为 PNG 截图, 供选版式和成品 QA。

用法:
    python snapshot.py <deck.pptx> <out_dir> [--width 1440]

实现优先级:
1. Windows + PowerPoint: COM 导出(保真度最高)
2. LibreOffice(soffice 在 PATH 上): 转 PDF 后用 PyMuPDF 切页(若装有 fitz)
3. 都不可用: 报错并提示——截图是可选增强, 没有截图时仍可仅凭 manifest 工作
"""

import argparse
import io
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

SCRIPT_DIR = Path(__file__).parent


def try_powerpoint_com(pptx, out_dir, width):
    if sys.platform != "win32":
        return False
    ps1 = SCRIPT_DIR / "snapshot.ps1"
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(ps1), "-Path", str(pptx), "-OutDir", str(out_dir),
         "-Width", str(width)],
        capture_output=True, text=True, timeout=300,
    )
    if proc.returncode == 0:
        print(proc.stdout.strip())
        return True
    print(proc.stderr.strip()[:800], file=sys.stderr)
    return False


def find_soffice():
    """在 PATH 与常见安装位置寻找 LibreOffice。"""
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    import glob
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice", "/usr/bin/libreoffice",
        "/usr/local/bin/soffice", "/snap/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ] + glob.glob("/opt/libreoffice*/program/soffice")
    for c in candidates:
        if Path(c).is_file():
            return c
    return None


def try_soffice(pptx, out_dir, width):
    soffice = find_soffice()
    if not soffice:
        return False
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", str(out_dir), str(pptx)],
        capture_output=True, text=True, timeout=300,
    )
    pdf = out_dir / (Path(pptx).stem + ".pdf")
    if proc.returncode != 0 or not pdf.is_file():
        return False
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print(f"已生成 PDF(未安装 pymupdf, 无法切成 PNG): {pdf}")
        return True
    doc = fitz.open(pdf)
    n_pages = len(doc)
    for i, page in enumerate(doc, 1):
        zoom = width / page.rect.width
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        pix.save(out_dir / f"slide-{i:02d}.png")
    doc.close()
    pdf.unlink()
    print(f"exported {n_pages} slides -> {out_dir}")
    return True


def probe():
    """报告本机可用的截图后端, 供 agent 在工作流早期决定 QA 方式。"""
    backends = []
    if sys.platform == "win32":
        backends.append("PowerPoint COM(将在导出时验证)")
    soffice = find_soffice()
    if soffice:
        try:
            import fitz  # noqa: F401
            backends.append(f"LibreOffice: {soffice} + PyMuPDF")
        except ImportError:
            backends.append(f"LibreOffice: {soffice} (缺 pymupdf, 只能出整份 PDF)")
    if backends:
        print("可用截图后端: " + "; ".join(backends))
        return 0
    print("无可用截图后端(无 PowerPoint COM, 无 LibreOffice)。\n"
          "QA 降级方案: 以 build_deck.py 校验输出为准逐条复核, 交付时说明未做视觉复核。")
    return 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", nargs="?")
    ap.add_argument("out_dir", nargs="?")
    ap.add_argument("--width", type=int, default=1440)
    ap.add_argument("--probe", action="store_true",
                    help="只探测截图后端是否可用, 不做导出")
    args = ap.parse_args()

    if args.probe:
        sys.exit(probe())
    if not args.pptx or not args.out_dir:
        ap.error("需要 <pptx> <out_dir> 两个参数(或使用 --probe)")

    pptx = Path(args.pptx).resolve()
    if not pptx.is_file():
        print(f"错误: 找不到文件 {pptx}", file=sys.stderr)
        sys.exit(2)
    Path(args.out_dir).mkdir(parents=True, exist_ok=True)

    if try_powerpoint_com(pptx, Path(args.out_dir).resolve(), args.width):
        return
    if try_soffice(pptx, args.out_dir, args.width):
        return
    print("错误: 本机既无 PowerPoint(COM)也无 LibreOffice, 无法截图。\n"
          "QA 降级方案: 以 build_deck.py 校验输出为准逐条复核(字数/残留), "
          "并在交付时说明未做视觉复核。", file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    main()
