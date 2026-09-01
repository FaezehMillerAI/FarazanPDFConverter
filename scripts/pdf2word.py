#!/usr/bin/env python3
"""
Convenience CLI script wrapper for OmniPDF.
Can be run directly with: python scripts/pdf2word.py <command>
"""

import sys
import os

# Add src/ to python search path
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src_dir = os.path.join(base_dir, "src")
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

from omnipdf.cli.main import app

if __name__ == "__main__":
    app()
