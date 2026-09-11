from pathlib import Path
import yaml

ROOT = Path(__file__).resolve().parents[1]


def test_student_example_has_required_keys():
    data = yaml.safe_load((ROOT / "student.yaml.example").read_text(encoding="utf-8"))
    for key in ("full_name", "name_short", "group", "faculty", "university", "city", "year"):
        assert key in data
