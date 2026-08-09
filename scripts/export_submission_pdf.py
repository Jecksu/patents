"""把 docs/submission 下的 DOCX 产物导出为同名 PDF。

CNIPA 电子申请以 PDF 收件，DOCX 在不同 Word 版本上的分页与字体回退并不
一致，因此正式提交前需要一份由本机 Word 渲染并固化的 PDF。

用法::

    py -3.12 scripts/export_submission_pdf.py          # 永久命名两份
    py -3.12 scripts/export_submission_pdf.py --all    # 连同轻量化摘要两份

本机未安装 pywin32，也没有 soffice，所以通过 PowerShell 驱动 Word COM，
导出格式常量 wdFormatPDF 取 17。
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SUB_DIR = ROOT / "docs" / "submission"
DISCLOSURE_DIR = SUB_DIR / "技术交底书"

#: wdFormatPDF；Word 的 SaveAs 用它指定 PDF 输出。
WD_FORMAT_PDF = 17

PERMANENT_TARGETS = [
    SUB_DIR / "永久命名-CAD模型差异对比-发明专利申请文件.docx",
    DISCLOSURE_DIR / "永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx",
]

LIGHTWEIGHT_TARGETS = [
    SUB_DIR / "跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-发明专利申请文件.docx",
    DISCLOSURE_DIR / "跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-专利技术交底书.docx",
]


PS_TEMPLATE = """
$ErrorActionPreference = 'Stop'
$pairs = @(
{pairs}
)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {{
    foreach ($pair in $pairs) {{
        # ReadOnly=$true、AddToRecentFiles=$false：不改动源 DOCX，也不污染最近文件列表。
        $doc = $word.Documents.Open($pair[0], $false, $true)
        try {{
            $doc.SaveAs([ref]$pair[1], [ref]{fmt})
            Write-Output ("OK`t" + $pair[1])
        }} finally {{
            $doc.Close([ref]0)
        }}
    }}
}} finally {{
    $word.Quit()
}}
"""


def check_lock(docx: Path) -> str | None:
    """DOCX 被 Word 打开时会留下 ``~$`` 占位文件，此时导出的是过期内容。"""
    lock = docx.with_name("~$" + docx.name[2:]) if len(docx.name) > 2 else None
    if lock and lock.exists():
        return f"{docx.name} 正被 Word 打开（存在 {lock.name}），请先关闭该文档"
    return None


def export(targets: list[Path]) -> list[Path]:
    missing = [t for t in targets if not t.exists()]
    if missing:
        raise FileNotFoundError("缺少待导出的 DOCX：" + "、".join(str(m) for m in missing))
    locks = [msg for msg in (check_lock(t) for t in targets) if msg]
    if locks:
        raise RuntimeError("；".join(locks))

    outputs = [t.with_suffix(".pdf") for t in targets]
    pairs = ",\n".join(
        f"    @('{src}', '{dst}')" for src, dst in zip(targets, outputs)
    )
    script = PS_TEMPLATE.format(pairs=pairs, fmt=WD_FORMAT_PDF)

    with tempfile.TemporaryDirectory() as tmp:
        ps1 = Path(tmp) / "export.ps1"
        # Windows PowerShell 5.1 按 BOM 判定 .ps1 编码，中文路径必须带 BOM。
        ps1.write_text(script, encoding="utf-8-sig")
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    if proc.returncode != 0:
        raise RuntimeError(f"Word 导出失败：\n{proc.stdout}\n{proc.stderr}")

    produced = [out for out in outputs if out.exists()]
    if len(produced) != len(outputs):
        raise RuntimeError("Word 未生成全部 PDF：\n" + proc.stdout)
    return produced


def main() -> None:
    targets = list(PERMANENT_TARGETS)
    if "--all" in sys.argv[1:]:
        targets += LIGHTWEIGHT_TARGETS
    for pdf in export(targets):
        print(f"{pdf}  {pdf.stat().st_size} bytes")


if __name__ == "__main__":
    main()
