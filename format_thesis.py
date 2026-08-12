#!/usr/bin/env python3
"""
format_thesis.py — Top-level entry point.

Usage:
    python format_thesis.py input.docx
    python format_thesis.py input.docx output.docx
    python format_thesis.py input.docx output.docx --report --analysis
"""
import sys
import os

# Allow running from the project directory without installing the package
sys.path.insert(0, os.path.dirname(__file__))

from thesis_formatter.cli import main

if __name__ == "__main__":
    sys.exit(main())
