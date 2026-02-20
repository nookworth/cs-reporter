# CS Reporter
# Launcher V2
# Launcher for CS Reporter V2 (Refactored System)
# **Date:** 2026-02-20
# **Status:** ✅ Completed
# **Component:** root

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.main_v2 import main

if __name__ == "__main__":
    if '--config' not in sys.argv:
        sys.argv.extend(['--config', 'config/demo_mapping_v2.yaml'])
    sys.exit(main())
