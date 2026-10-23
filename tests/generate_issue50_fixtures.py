"""Create the compact WATER resume matrix used by issue #50.

The source fixture deliberately has different profiles W (A) and W2 (B).
The default case removes the optional AC:AL time-detail rows while retaining
the stage-plan row.  This is the path that previously restored HOLD/A instead
of the effective START/B.  The other cases make the intended mode explicit.
"""
from __future__ import annotations
import argparse, copy, json
from pathlib import Path


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()
    base = json.loads(args.source.read_text(encoding="utf-8"))
    # Stage 4 is the WATER transition W -> W2; stage 5 is the following WAIT.
    # Keep the stage-plan rows in every case so this matrix does not exercise
    # the unrelated "no WATER plan" rejection path.
    modes = {
        "default_start_no_detail": None,
        "explicit_hold": "HOLD",
        "explicit_start": "START",
        "explicit_linear": "LINEAR",
    }
    args.out.mkdir(parents=True, exist_ok=True)
    for suffix, mode in modes.items():
        d = copy.deepcopy(base)
        d["name"] = "issue50_" + suffix
        if mode is None:
            d["water"].pop("timeOptions", None)
        else:
            d["water"]["timeOptions"] = [[4, "D", mode, "INSTANT", 0.0001, 0.1, 0.2, "", "", ""]]
        (args.out / (d["name"] + ".json")).write_text(
            json.dumps(d, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )


if __name__ == "__main__":
    main()
