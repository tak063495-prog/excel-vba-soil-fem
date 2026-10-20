"""Check the actual SVGs embedded by Excel, which removes custom data-* tags."""
import argparse
import json
import math
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


def extract(directory, output):
    records = {}
    crosses = {}
    for mode in ("A", "B"):
        with zipfile.ZipFile(directory / f"{mode}.xlsm") as archive:
            candidates = [archive.read(n) for n in archive.namelist()
                          if n.startswith("xl/media/") and n.endswith(".svg")]
        candidates = [b for b in candidates if b"increment-overlay" in b]
        assert len(candidates) == 1
        raw = candidates[0]
        root = ET.fromstring(raw)
        contour = root.find(".//*[@id='shear-increment-contours']")
        cross = root.find(".//*[@id='plastic-principal-crosses']")
        assert contour is not None and cross is not None
        n = sum(x.tag.endswith("path") for x in contour.iter())
        lines = [x.attrib for x in cross.iter() if x.tag.endswith("line")]
        assert lines
        max_half_axis = max(math.hypot(float(x["x2"])-float(x["x1"]),
                                       float(x["y2"])-float(x["y1"])) for x in lines)
        assert max_half_axis > 6, "real cross must remain visible despite a very short deformed edge"
        assert (n > 0) if mode == "A" else (n == 0)
        crosses[mode] = lines
        text = " ".join(x.text or "" for x in root.iter() if x.tag.endswith("text"))
        assert "Fs=2.2000" in text
        assert ("Y方向応力" if mode == "A" else "最大せん断ひずみ増分") in text
        (output / f"actual_{mode}.svg").write_bytes(raw)
        records[mode] = dict(contour_segments=n, principal_half_axes=len(lines),
                             max_half_axis_pixels=max_half_axis,
                             embedded_svg_bytes=len(raw))
    assert crosses["A"] == crosses["B"], "A/B must preserve the same independent plastic cross"
    (output / "embedded_svg_checks.json").write_text(
        json.dumps(dict(status="PASS", modes=records), indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--directory", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)
    extract(args.directory, args.out)
    print("PASS actual embedded A contours / B color and identical plastic principal crosses")
