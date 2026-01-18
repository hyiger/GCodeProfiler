import sys
from pathlib import Path

# Ensure the repo root (containing the  gcode_profiler package) is importable when
# running pytest without installation.
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
