from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

ALLOWED_SOURCES = ("browser", "terminal", "ide", "diagram", "other")
_IDX = re.compile(r"^(\d+)-")


class ManifestError(Exception):
    pass


def load_manifest(out_dir: Path) -> list[dict]:
    path = Path(out_dir) / "manifest.json"
    if not path.is_file():
        return []
    return json.loads(path.read_text(encoding="utf-8"))


def _save_manifest(out_dir: Path, data: list[dict]) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "manifest.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _next_index(out_dir: Path) -> int:
    nums = []
    if out_dir.is_dir():
        for p in out_dir.iterdir():
            m = _IDX.match(p.name)
            if m:
                nums.append(int(m.group(1)))
    return (max(nums) + 1) if nums else 1


def register_image(
    src: Path,
    out_dir: Path,
    caption: str,
    source: str,
    slug: str | None = None,
) -> Path:
    if source not in ALLOWED_SOURCES:
        raise ValueError("Неизвестный source")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    n = _next_index(out_dir)
    slug = slug or "shot"
    dest = out_dir / f"{n:02d}-{slug}.png"
    shutil.copy2(src, dest)
    data = load_manifest(out_dir)
    data.append({"file": dest.name, "caption": caption, "source": source, "ok": False})
    _save_manifest(out_dir, data)
    return dest


def capture_screen(out_dir: Path, caption: str, source: str, slug: str | None = None) -> Path:
    from PIL import ImageGrab

    tmp = Path(out_dir) / "_capture.png"
    tmp.parent.mkdir(parents=True, exist_ok=True)
    ImageGrab.grab().save(tmp)
    try:
        return register_image(tmp, out_dir, caption, source, slug=slug)
    finally:
        tmp.unlink(missing_ok=True)


def mark_ok(out_dir: Path, file_name: str) -> None:
    data = load_manifest(out_dir)
    for item in data:
        if item["file"] == file_name:
            item["ok"] = True
            _save_manifest(Path(out_dir), data)
            return
    raise ManifestError(f"Нет записи в манифесте: {file_name}")


def main() -> int:
    p = argparse.ArgumentParser()
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--from-file")
    g.add_argument("--capture", action="store_true")
    g.add_argument("--mark-ok")
    p.add_argument("--out-dir", required=True)
    p.add_argument("--caption", default="")
    p.add_argument("--source", default="other")
    p.add_argument("--slug")
    args = p.parse_args()
    out = Path(args.out_dir)
    try:
        if args.mark_ok:
            mark_ok(out, args.mark_ok)
        elif args.capture:
            if not args.caption:
                print("Нужен --caption", file=sys.stderr)
                return 1
            path = capture_screen(out, args.caption, args.source, slug=args.slug)
            print(path)
        else:
            if not args.caption:
                print("Нужен --caption", file=sys.stderr)
                return 1
            path = register_image(Path(args.from_file), out, args.caption, args.source, slug=args.slug)
            print(path)
    except (ManifestError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
