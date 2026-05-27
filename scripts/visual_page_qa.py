#!/usr/bin/env python3
"""Screen rendered PDF page images for isolated figure/table page layout."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def load_image(path: Path):
    try:
        from PIL import Image
    except Exception as exc:  # pragma: no cover
        raise SystemExit("Pillow is required: python -m pip install pillow") from exc
    return Image.open(path).convert("RGB")


def content_bbox(
    path: Path,
    *,
    threshold: int,
    crop_top_ratio: float,
    crop_bottom_ratio: float,
    crop_left_ratio: float,
    crop_right_ratio: float,
) -> dict[str, object]:
    image = load_image(path)
    width, height = image.size
    left_crop = int(width * crop_left_ratio)
    right_crop = width - int(width * crop_right_ratio)
    top_crop = int(height * crop_top_ratio)
    bottom_crop = height - int(height * crop_bottom_ratio)
    pixels = image.load()

    xs: list[int] = []
    ys: list[int] = []
    for y in range(top_crop, bottom_crop):
        for x in range(left_crop, right_crop):
            r, g, b = pixels[x, y]
            if min(r, g, b) < threshold:
                xs.append(x)
                ys.append(y)

    if not xs:
        return {
            "path": str(path),
            "width": width,
            "height": height,
            "content_found": False,
            "ok": False,
            "reason": "no non-white content detected in analysis area",
        }

    bbox = {
        "left": min(xs),
        "top": min(ys),
        "right": max(xs),
        "bottom": max(ys),
    }
    center_x = (bbox["left"] + bbox["right"]) / 2
    center_y = (bbox["top"] + bbox["bottom"]) / 2
    page_center_x = width / 2
    page_center_y = (top_crop + bottom_crop) / 2

    return {
        "path": str(path),
        "width": width,
        "height": height,
        "content_found": True,
        "bbox": bbox,
        "center_offset_ratio": {
            "x": (center_x - page_center_x) / width,
            "y": (center_y - page_center_y) / height,
        },
        "edge_margin_ratio": {
            "left": (bbox["left"] - left_crop) / width,
            "right": (right_crop - bbox["right"]) / width,
            "top": (bbox["top"] - top_crop) / height,
            "bottom": (bottom_crop - bbox["bottom"]) / height,
        },
    }


def evaluate(result: dict[str, object], center_tolerance: float, edge_tolerance: float) -> dict[str, object]:
    if not result.get("content_found"):
        return result
    offsets = result["center_offset_ratio"]  # type: ignore[index]
    margins = result["edge_margin_ratio"]  # type: ignore[index]
    failures: list[str] = []
    if abs(offsets["x"]) > center_tolerance:  # type: ignore[index]
        failures.append("horizontal center offset too large")
    if abs(offsets["y"]) > center_tolerance:  # type: ignore[index]
        failures.append("vertical center offset too large")
    for side in ("left", "right", "top", "bottom"):
        if margins[side] < edge_tolerance:  # type: ignore[index]
            failures.append(f"content too close to {side} edge")
    result["ok"] = not failures
    if failures:
        result["failures"] = failures
    return result


def collect_paths(inputs: list[str]) -> list[Path]:
    paths: list[Path] = []
    for item in inputs:
        path = Path(item).expanduser()
        if path.is_dir():
            paths.extend(sorted(path.glob("*.png")))
            paths.extend(sorted(path.glob("*.jpg")))
            paths.extend(sorted(path.glob("*.jpeg")))
        else:
            paths.append(path)
    return paths


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("images", nargs="+", help="Rendered PDF page images or directories.")
    parser.add_argument("--threshold", type=int, default=245, help="Pixel channel threshold for non-white content.")
    parser.add_argument("--center-tolerance", type=float, default=0.08, help="Allowed center offset ratio.")
    parser.add_argument("--edge-tolerance", type=float, default=0.02, help="Minimum content margin ratio.")
    parser.add_argument("--crop-top-ratio", type=float, default=0.03)
    parser.add_argument("--crop-bottom-ratio", type=float, default=0.08, help="Ignore page number/footer area.")
    parser.add_argument("--crop-left-ratio", type=float, default=0.02)
    parser.add_argument("--crop-right-ratio", type=float, default=0.02)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    paths = collect_paths(args.images)
    missing = [str(path) for path in paths if not path.exists()]
    if missing:
        raise SystemExit("Missing image(s): " + ", ".join(missing))

    results = [
        evaluate(
            content_bbox(
                path,
                threshold=args.threshold,
                crop_top_ratio=args.crop_top_ratio,
                crop_bottom_ratio=args.crop_bottom_ratio,
                crop_left_ratio=args.crop_left_ratio,
                crop_right_ratio=args.crop_right_ratio,
            ),
            args.center_tolerance,
            args.edge_tolerance,
        )
        for path in paths
    ]
    ok = all(bool(result.get("ok")) for result in results)
    if args.json:
        print(json.dumps({"ok": ok, "pages": results}, ensure_ascii=False, indent=2))
    else:
        for result in results:
            status = "OK" if result.get("ok") else "CHECK"
            print(f"{status}: {result['path']}")
            for failure in result.get("failures", []):  # type: ignore[union-attr]
                print(f"  - {failure}")
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())

