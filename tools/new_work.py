from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path

from tools.paths import assert_not_sample, repo_root


class NewWorkError(Exception):
    pass


_LAB = re.compile(r"^lab-(\d+)-")


def _next_number(subject_dir: Path) -> int:
    nums = []
    if subject_dir.is_dir():
        for p in subject_dir.iterdir():
            m = _LAB.match(p.name)
            if m:
                nums.append(int(m.group(1)))
    return (max(nums) + 1) if nums else 1


def create_work(
    subject: str,
    slug: str,
    pdf: Path | None = None,
    *,
    number: int | None = None,
    reset_docx: bool = False,
    move_pdf: bool = False,
    repo: Path | None = None,
) -> Path:
    root = Path(repo) if repo is not None else repo_root()
    subject_dir = root / "subjects" / subject
    sample = subject_dir / "sample.docx"
    if not sample.is_file():
        raise NewWorkError(f"Нет эталона subjects/{subject}/sample.docx")
    n = number if number is not None else _next_number(subject_dir)
    work = subject_dir / f"lab-{n:02d}-{slug}"
    work.mkdir(parents=True, exist_ok=True)
    dest = work / "report.docx"
    assert_not_sample(dest)
    if reset_docx or not dest.is_file():
        shutil.copy2(sample, dest)
    if pdf is not None:
        target_pdf = work / "assignment.pdf"
        if not target_pdf.is_file():
            if move_pdf:
                shutil.move(str(pdf), str(target_pdf))
            else:
                shutil.copy2(pdf, target_pdf)
    return work


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--subject", required=True)
    p.add_argument("--slug", required=True)
    p.add_argument("--pdf")
    p.add_argument("--number", type=int)
    p.add_argument("--reset-docx", action="store_true")
    p.add_argument("--repo")
    args = p.parse_args()
    try:
        work = create_work(
            args.subject,
            args.slug,
            Path(args.pdf) if args.pdf else None,
            number=args.number,
            reset_docx=args.reset_docx,
            repo=Path(args.repo) if args.repo else None,
        )
    except NewWorkError as e:
        print(str(e), file=sys.stderr)
        return 1
    print(work)
    return 0


if __name__ == "__main__":
    sys.exit(main())
