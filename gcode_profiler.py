#!/usr/bin/env python3
"""Backwards-compatible entry point.

This wrapper keeps the original `python gcode_profiler.py <file.gcode>` UX,
while the implementation lives in the `gcode_profiler` package.
"""

from gcode_profiler.cli import main


if __name__ == '__main__':
    main()
