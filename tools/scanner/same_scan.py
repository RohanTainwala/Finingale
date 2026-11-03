#!/usr/bin/env python3
"""Exit 0 when a freshly written scan matches the committed one (same trading day and
the same matches), so the daily jobs and their backup runs only commit real changes.

    python3 same_scan.py data/daily-scan.json
"""
import json
import subprocess
import sys


def main(path):
    try:
        old = json.loads(subprocess.run(["git", "show", "HEAD:" + path], capture_output=True,
                                        check=True).stdout)
    except (subprocess.CalledProcessError, ValueError):
        return 1  # nothing committed yet
    with open(path) as f:
        new = json.load(f)
    for d in (old, new):
        d.pop("generated_at", None)
    return 0 if old == new else 1


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    sys.exit(main(sys.argv[1]))
