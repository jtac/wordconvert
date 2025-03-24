#!/usr/bin/env python3
"""
WordToPPT CLI Tool - Official Entry Point

This script launches the WordToPPT command-line interface.
"""

import sys
from pathlib import Path

try:
    from wordtoppt.wordtoppt import app
except ImportError:
    # Add the package directory to the Python path if import fails
    package_dir = Path(__file__).resolve().parent
    if str(package_dir) not in sys.path:
        sys.path.insert(0, str(package_dir))
    from wordtoppt.wordtoppt import app

if __name__ == "__main__":
    # Run the application
    app()
