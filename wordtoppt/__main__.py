"""
Main entry point for the wordtoppt package.
"""

import sys
from pathlib import Path
from wordtoppt.wordtoppt import app

# Add parent directory to path if running directly
parent_dir = Path(__file__).resolve().parent.parent
if str(parent_dir) not in sys.path:
    sys.path.insert(0, str(parent_dir))

if __name__ == "__main__":
    app()
