"""Record the tested environment without installing unrelated packages at runtime."""

import importlib.metadata
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main():
    resolved = {}
    for distribution in importlib.metadata.distributions():
        raw_name = distribution.metadata.get("Name")
        if not raw_name:
            warnings.warn(
                f"Ignoring metadata without a package name under {distribution.locate_file('')}",
                stacklevel=1,
            )
            continue
        name = raw_name.lower().replace("_", "-")
        if name in {"smart-course-planner", "pip", "setuptools"}:
            continue
        resolved[name] = distribution.version
    output = "# Tested environment constraints. Regenerate with tools/lock_dependencies.py.\n"
    output += (
        "# Constraints do not install packages; pyproject.toml selects runtime/dev dependencies.\n"
    )
    output += "\n".join(f"{name}=={version}" for name, version in sorted(resolved.items())) + "\n"
    (ROOT / "constraints.txt").write_text(output, encoding="utf-8")


if __name__ == "__main__":
    main()
