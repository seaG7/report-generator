from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

from tools.new_work import NewWorkError, create_work
from tools.paths import repo_root


class PickupError(Exception):
    pass


@dataclass
class Pending:
    subject: str
    pdf: Path


@dataclass
class Preview:
    subject: str
    pdf: Path
    md_path: Path


def slug_from_filename(name: str) -> str:
    stem = Path(name).stem
    slug = re.sub(r"[^\w]+", "-", stem, flags=re.UNICODE).strip("-").lower()
    slug = re.sub(r"-{2,}", "-", slug)
    return (slug or "assignment")[:40]


def list_pending(repo: Path | None = None) -> list[Pending]:
    root = Path(repo) if repo is not None else repo_root()
    subjects = root / "subjects"
    if not subjects.is_dir():
        return []
    pending: list[Pending] = []
    for subj in sorted(subjects.iterdir()):
        if not subj.is_dir() or subj.name.startswith("."):
            continue
        inbox = subj / "inbox"
        if inbox.is_dir():
            for pdf in sorted(inbox.glob("*.pdf")):
                pending.append(Pending(subject=subj.name, pdf=pdf))
        for pdf in sorted(subj.glob("*.pdf")):
            if pdf.parent == subj:
                pending.append(Pending(subject=subj.name, pdf=pdf))
    return pending


def preview_pending(repo: Path | None = None, convert=None) -> list[Preview]:
    if convert is None:
        from tools.pdf_to_md import convert_pdf

        convert = convert_pdf
    root = Path(repo) if repo is not None else repo_root()
    previews: list[Preview] = []
    used_dirs: set[Path] = set()
    for item in list_pending(root):
        subject_dir = root / "subjects" / item.subject
        base = slug_from_filename(item.pdf.name)
        preview_dir = subject_dir / ".preview" / base
        n = 2
        while preview_dir in used_dirs:
            preview_dir = subject_dir / ".preview" / f"{base}-{n}"
            n += 1
        used_dirs.add(preview_dir)
        result = convert(item.pdf, preview_dir)
        previews.append(Preview(subject=item.subject, pdf=item.pdf, md_path=result.md_path))
    return previews


def claim_pending(item: Pending, repo: Path | None = None) -> Path:
    return claim_bundle(item.subject, [item.pdf], repo=repo)


def _move_to(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()
    shutil.move(str(src), str(dest))


def claim_bundle(
    subject: str,
    assignment_pdfs: list[Path],
    lecture_pdfs: list[Path] | None = None,
    other_pdfs: list[Path] | None = None,
    *,
    slug: str | None = None,
    repo: Path | None = None,
) -> Path:
    assignment_pdfs = [Path(p) for p in assignment_pdfs]
    lecture_pdfs = [Path(p) for p in (lecture_pdfs or [])]
    other_pdfs = [Path(p) for p in (other_pdfs or [])]
    if not assignment_pdfs:
        raise PickupError("Нет PDF задания")
    root = Path(repo) if repo is not None else repo_root()
    first = assignment_pdfs[0]
    try:
        work = create_work(
            subject,
            slug or slug_from_filename(first.name),
            first,
            move_pdf=True,
            repo=root,
        )
    except NewWorkError as exc:
        raise PickupError(str(exc)) from exc
    for i, pdf in enumerate(assignment_pdfs[1:], start=2):
        dest = work / f"assignment-{i:02d}-{slug_from_filename(pdf.name)}.pdf"
        _move_to(pdf, dest)
    subject_dir = root / "subjects" / subject
    for pdf in lecture_pdfs:
        dest_subj = subject_dir / "lectures" / pdf.name
        _move_to(pdf, dest_subj)
        lab_lec = work / "lectures" / pdf.name
        lab_lec.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(dest_subj, lab_lec)
    for pdf in other_pdfs:
        _move_to(pdf, subject_dir / "materials" / pdf.name)
    return work


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--repo")
    p.add_argument("--json", action="store_true")
    p.add_argument("--preview", action="store_true")
    p.add_argument("--claim", action="store_true")
    p.add_argument("--subject")
    p.add_argument("--assignment", action="append", default=[])
    p.add_argument("--lecture", action="append", default=[])
    p.add_argument("--other", action="append", default=[])
    args = p.parse_args()
    repo = Path(args.repo) if args.repo else None
    try:
        if args.claim:
            if not args.subject or not args.assignment:
                print("Для --claim нужны --subject и хотя бы один --assignment", file=sys.stderr)
                return 1
            work = claim_bundle(
                args.subject,
                [Path(x) for x in args.assignment],
                [Path(x) for x in args.lecture],
                [Path(x) for x in args.other],
                repo=repo,
            )
            print(work)
            return 0
        previews = preview_pending(repo)
    except PickupError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    if not previews:
        print(
            "Нет PDF. Положите задание и при необходимости лекции в subjects/<предмет>/inbox/",
            file=sys.stderr,
        )
        return 2
    payload = [
        {"subject": item.subject, "pdf": str(item.pdf), "md": str(item.md_path)}
        for item in previews
    ]
    if args.json or args.preview:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
