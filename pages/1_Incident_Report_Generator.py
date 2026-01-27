
"""
Wrapper page that runs your existing IR_gen.py without modifying it.
Place this file under /pages in the same project as IR_gen.py.
"""
import runpy
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
runpy.run_path(str(ROOT / "IR_gen.py"), run_name="__main__")
