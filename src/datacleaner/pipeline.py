"""Overall flow: read -> normalize/validate -> dedupe -> classify -> write (SPEC sections 11-13)."""
import hashlib
import logging
import os
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

from . import io as dc_io
from .dedupe import mark_duplicates
from .report import OUTPUT_FILES, build_summary, summary_text, write_reports
from .validate import assign_status, validate_row

log = logging.getLogger("datacleaner")
CATEGORIES = ("clean", "duplicate", "invalid", "needs_review")


def _setup_logging(out_dir=None):
    log.setLevel(logging.INFO)
    log.propagate = False
    for h in list(log.handlers):
        h.close()
        log.removeHandler(h)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%H:%M:%S")
    handlers = [logging.StreamHandler(sys.stdout)]
    if out_dir is not None:
        handlers.append(logging.FileHandler(Path(out_dir) / "process.log", "w", encoding="utf-8"))
    for h in handlers:
        h.setFormatter(fmt)
        log.addHandler(h)


def _sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _norm_path(p):
    return os.path.normcase(str(Path(p).resolve()))


def run(inputs, output_dir="output", date_order=None):
    """Process the inputs; returns the exit code (0 ok, 1 no/unreadable input, 2 required columns missing)."""
    started = time.perf_counter()
    processed_at = datetime.now()
    out_dir = Path(output_dir)
    _setup_logging()

    files = dc_io.find_input_files(inputs)
    if not files:
        log.error("No input files found.")
        return 1
    outputs = {_norm_path(out_dir / name) for name in OUTPUT_FILES}
    clash = [f for f in files if _norm_path(f) in outputs]
    if clash:
        log.error("Refusing to run: an input file would be overwritten by an output file: %s",
                  ", ".join(str(f) for f in clash))
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    _setup_logging(out_dir)
    try:
        return _run(files, out_dir, date_order, started, processed_at)
    except (OSError, RuntimeError) as e:
        log.error("%s", e)
        return 1
    finally:
        _setup_logging()


def _run(files, out_dir, date_order, started, processed_at):
    log.info("Input files: %s", ", ".join(f.name for f in files))
    hashes = {f: _sha256(f) for f in files}
    alias_map = dc_io.load_alias_map()

    rows, processed, skipped = [], 0, []
    for f in files:
        try:
            rows.extend(dc_io.read_file(f, alias_map))
            processed += 1
        except dc_io.SkipFile as e:
            log.error("%s", e)
            skipped.append((f.name, e))
    if not processed:
        return 2 if all(e.missing_columns for _, e in skipped) else 1

    for row in rows:
        validate_row(row, date_order)
    mark_duplicates(rows)
    assign_status(rows)

    counts = Counter(r["status"] for r in rows)
    log.info("Processed %d rows: clean %d / duplicate %d / invalid %d / needs_review %d",
             len(rows), counts["clean"], counts["duplicate"], counts["invalid"],
             counts["needs_review"])
    unknown = {r["status"] for r in rows} - set(CATEGORIES)
    if unknown:
        raise RuntimeError(f"Row check failed: unexpected status {sorted(unknown)}.")

    summary = build_summary(rows, processed, [name for name, _ in skipped], processed_at,
                            time.perf_counter() - started)
    written = write_reports(rows, summary, out_dir)
    if sum(written.values()) != len(rows) or written != {c: counts[c] for c in CATEGORIES}:
        raise RuntimeError(f"Row count check failed: {written} written vs {len(rows)} input rows.")

    unchanged = [f for f in files if _sha256(f) == hashes[f]]
    if len(unchanged) != len(files):
        raise RuntimeError("An input file changed during processing (SHA-256 mismatch).")
    log.info("Input files unchanged (SHA-256 verified): %d file(s)", len(files))

    print()
    print(summary_text(summary))
    print(f"Output folder: {out_dir.resolve()}")
    for name in OUTPUT_FILES:
        print(f"  {name}")
    return 0
