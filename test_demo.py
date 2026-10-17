"""无需真实名单或外部字体的最小验收。"""

import csv
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

from pypdf import PdfReader


with TemporaryDirectory() as temporary:
    output = Path(temporary) / "output"
    script = Path(__file__).with_name("生成号码布.py")
    result = subprocess.run(
        [sys.executable, str(script), "--output", str(output)],
        capture_output=True, text=True, check=False,
    )
    assert result.returncode == 0, result.stderr + result.stdout
    batches = [path for path in output.iterdir() if path.is_dir() and path.name != "_demo_templates"]
    assert len(batches) == 1
    batch = batches[0]
    pdfs = list(batch.rglob("*.pdf"))
    assert len(pdfs) == 3
    for pdf in pdfs:
        reader = PdfReader(pdf)
        assert len(reader.pages) == 2
        text = "".join(reader.pages[1].extract_text().split())
        assert "".join(pdf.stem.split()) in text
        assert "示例大学" in text
        assert "PLACEHOLDER" not in text
    with (batch / "生成清单.csv").open(encoding="utf-8-sig", newline="") as handle:
        records = list(csv.DictReader(handle))
    assert len(records) == 3 and all(record["状态"] == "已生成" for record in records)
    print("匿名示例通过：3 份双页 PDF，姓名和学校可检索，占位文字已清除。")
