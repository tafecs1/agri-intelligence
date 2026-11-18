"""Run agri_colab_auto.py outside Colab (GitHub Actions or any PC), exactly like "Runtime -> Run all" in Colab.

The notebook logic is NOT re-implemented here: this file reads agri_colab_auto.py, splits it at the
"# ===== CELL n :" markers and executes the cells in order in ONE shared namespace (as Colab does).

  python run.py                      # all cells, all India
  python run.py --cells 1-14,17      # main workbook only (no trend module)
  python run.py --cells 1-3,7,9,11,15-17   (any subset; order is always the notebook order)
  AGRI_STATES="TAMIL NADU" python run.py   # quick test on one state

Working files / checkpoints go to ./work/<scope>/ so a test run never mixes with the all-India run.
"""
import argparse
import hashlib
import os
import re
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
CODE = os.path.join(HERE, "agri_colab_auto.py")
ORDER = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "12B", "13", "14", "15", "16", "17"]


def split_cells(src):
    parts = re.split(r"(?m)^(?=# ===== CELL [0-9A-Z]+ :)", src)
    cells = {}
    for p in parts[1:]:
        cells[re.match(r"# ===== CELL ([0-9A-Z]+)", p).group(1)] = p
    return cells


def parse_cells(spec):
    if not spec:
        return list(ORDER)
    want = set()
    for tok in spec.split(","):
        tok = tok.strip().upper()
        if "-" in tok:
            a, b = tok.split("-")
            i, j = ORDER.index(a), ORDER.index(b)
            want.update(ORDER[i:j + 1])
        elif tok:
            want.add(tok)
    bad = want - set(ORDER)
    if bad:
        sys.exit(f"Unknown cell(s): {sorted(bad)}. Valid: {ORDER}")
    return [c for c in ORDER if c in want]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cells", default=os.environ.get("AGRI_CELLS", ""), help="e.g. 1-14,17 (default: all)")
    args = ap.parse_args()
    cells = split_cells(open(CODE, encoding="utf-8").read())
    run = parse_cells(args.cells)
    states = os.environ.get("AGRI_STATES", "").strip()
    scope = "all-india" if not states else "test-" + hashlib.md5(states.upper().encode()).hexdigest()[:8]
    work = os.path.join(HERE, "work", scope)
    os.makedirs(work, exist_ok=True)
    os.chdir(work)
    print(f"Cells: {', '.join(run)} | scope: {scope} | working folder: {work}", flush=True)
    G = {"__name__": "__main__", "__file__": CODE}
    failed = None
    for c in run:
        t0 = time.time()
        print(f"\n{'=' * 20} CELL {c} {'=' * 20}", flush=True)
        try:
            exec(compile(cells[c], f"agri_colab_auto.py [CELL {c}]", "exec"), G)
        except SystemExit as e:                         # the notebook stops a cell this way with a message
            failed = (c, f"stopped: {e}")
        except Exception:
            traceback.print_exc()
            failed = (c, "error (traceback above)")
        print(f"[CELL {c}] {(time.time() - t0) / 60:.1f} min", flush=True)
        if failed:
            print(f"\nRUN STOPPED at CELL {failed[0]} - {failed[1]}  (Colab 'Run all' stops here too)", flush=True)
            if c != "17" and "17" in run and "RECOVERY_LOG" in G:
                try:
                    exec(compile(cells["17"], "agri_colab_auto.py [CELL 17]", "exec"), G)
                except Exception:
                    pass
            break
    links = G.get("SHEET_URLS") or {}
    with open(os.path.join(HERE, "sheet_links.md"), "w", encoding="utf-8") as f:
        f.write("| Sheet | Link |\n|---|---|\n" + "".join(f"| {k} | {v or 'not updated'} |\n" for k, v in links.items()))
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
