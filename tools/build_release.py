"""Build an explicit, data-free Azure Functions deployment artifact."""

import argparse
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def build(output: Path) -> None:
    files = [
        ROOT / name
        for name in (
            "function_app.py",
            "host.json",
            "requirements.txt",
            "pyproject.toml",
            "constraints.txt",
        )
    ]
    files.extend(sorted((ROOT / "src" / "course_planner").rglob("*.py")))
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, path.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(output) as archive:
        if archive.testzip() is not None:
            raise RuntimeError("Deployment archive verification failed")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=ROOT / "dist" / "release.zip")
    args = parser.parse_args()
    build(args.output)
    print(args.output.resolve())
