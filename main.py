"""CLI entry point: python main.py INPUT [INPUT ...] [--output DIR] [--date-order {mdy,dmy}]"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from datacleaner.pipeline import run  # noqa: E402


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Clean, validate and de-duplicate customer lists (CSV / Excel).")
    parser.add_argument("inputs", nargs="+", metavar="INPUT",
                        help="CSV/XLSX file or a folder containing them")
    parser.add_argument("--output", default="output", metavar="DIR",
                        help="output folder (default: output)")
    parser.add_argument("--date-order", choices=["mdy", "dmy"],
                        help="how to read ambiguous dates such as 03/04/2024 "
                             "(default: send them to needs_review)")
    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="replace")  # never crash on console encoding
    return run(args.inputs, args.output, args.date_order)


if __name__ == "__main__":
    sys.exit(main())
