"""
PyInstaller entry-point.
When packaged as a Windows .exe this module is executed instead of app.py.
It resolves the data directory relative to the executable location.
"""
import os
import sys

# When frozen by PyInstaller, sys._MEIPASS is the temp extraction dir.
# The .exe itself lives at sys.executable.
_FROZEN = getattr(sys, "frozen", False)

if _FROZEN:
    _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))

# Default data directory: same folder as the .exe (or script)
_default_data = os.path.join(_base_dir, "data")

# Inject default --data if user didn't specify one
if "--data" not in sys.argv:
    sys.argv += ["--data", _default_data]

from app import main  # noqa: E402

main()
