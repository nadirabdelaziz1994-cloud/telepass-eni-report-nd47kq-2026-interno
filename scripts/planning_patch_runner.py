from pathlib import Path
import subprocess
import sys

p = Path("scripts/planning_patch.py")
text = p.read_text(encoding="utf-8")
text = text.replace("min(ws.max_row, 20) + 1", "min(ws.max_row or 20, 20) + 1")
text = text.replace("range(1, ws.max_column + 1)", "range(1, (ws.max_column or 80) + 1)")
p.write_text(text, encoding="utf-8")

raise SystemExit(subprocess.call([sys.executable, "scripts/planning_patch.py"]))
