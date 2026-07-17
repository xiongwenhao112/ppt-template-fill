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
import shutil
import subprocess
import sys
from pathlib import Path

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


def try_soffice(pptx, out_dir, width):
    soffice = shutil.which("soffice")
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
    for i, page in enumerate(doc, 1):
        zoom = width / page.rect.width
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        pix.save(out_dir / f"slide-{i:02d}.png")
    doc.close()
    pdf.unlink()
    print(f"exported {len(doc)} slides -> {out_dir}")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("out_dir")
    ap.add_argument("--width", type=int, default=1440)
    args = ap.parse_args()

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
          "截图是可选步骤: 可跳过, 仅凭 manifest 的槽位信息继续。", file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    main()
