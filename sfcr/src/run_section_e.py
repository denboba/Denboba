"""Run from date/:  python3 ../src/run_section_e.py"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from sfcr.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
