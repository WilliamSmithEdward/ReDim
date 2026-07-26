"""Static analysis gate: pyvbaanalysis over every VBA source in the repo.

Any finding fails the gate. ROneCOne.cls joins the analysis set so
cross-module references from ReDim sources resolve, but findings inside
ROneCOne itself (there are none today) would be reported under its own name
and are not ReDim's to fix.
"""

from __future__ import annotations

import sys
from pathlib import Path

from pyvbaanalysis import analyze_loose_files

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vba_sources import DEMO_VBA, ROOT, SRC, TESTS_VBA, ronecone_class_path  # noqa: E402


def collect_sources() -> list[str]:
    sources: list[Path] = [ronecone_class_path()]
    for folder in (SRC, TESTS_VBA, DEMO_VBA):
        if folder.is_dir():
            sources.extend(sorted(folder.glob("*.bas")))
            sources.extend(sorted(folder.glob("*.cls")))
    return [str(path) for path in sources]


def main() -> int:
    sources = collect_sources()
    results = analyze_loose_files(sources)
    finding_count = 0
    for module_name, problems in results.items():
        for problem in problems:
            finding_count += 1
            print(
                f"{module_name}: {problem.severity.value} {problem.code} "
                f"{problem.message}"
            )
    print(f"analyzed {len(sources)} sources, findings: {finding_count}")
    return 1 if finding_count else 0


if __name__ == "__main__":
    raise SystemExit(main())
