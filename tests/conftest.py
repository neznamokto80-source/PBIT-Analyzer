import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WORK_DIR = ROOT / "_work"

if str(WORK_DIR) not in sys.path:
    sys.path.insert(0, str(WORK_DIR))
