"""Builds ReDim test and demo workbooks by injecting VBA sources with pyOpenVBA.

Every build verifies a byte-for-byte round trip of each injected module,
matching the ROneCOne pipeline discipline.
"""

from __future__ import annotations

import sys
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vba_sources import (  # noqa: E402
    DEMO_VBA,
    OUTPUT,
    ROOT,
    SRC,
    TESTS_VBA,
    prepare_class_source,
    read_vba,
    ronecone_class_path,
)


def framework_modules() -> dict[str, tuple[str, VBAModuleKind]]:
    return {
        "ROneCOne": (
            prepare_class_source(ronecone_class_path()),
            VBAModuleKind.other,
        ),
        "ReDimUI": (
            prepare_class_source(SRC / "ReDimUI.cls"),
            VBAModuleKind.other,
        ),
        "ReDimHost": (read_vba(SRC / "ReDimHost.bas"), VBAModuleKind.standard),
    }


def build_workbook(target: Path, extra_modules: dict[str, tuple[str, VBAModuleKind]]) -> Path:
    modules = framework_modules() | extra_modules
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    with ExcelFile.create_new(target) as workbook:
        project = workbook.vba_project()
        for name, (source, kind) in modules.items():
            project.add_module(name, source, kind=kind)
        workbook.save()
    with ExcelFile(target) as verification:
        actual = set(verification.module_names())
        missing = set(modules) - actual
        if missing:
            raise RuntimeError(f"{target.name} is missing modules: {sorted(missing)}")
        for name, (source, _) in modules.items():
            if verification.get_module(name) != source:
                raise RuntimeError(
                    f"{target.name} module {name} did not round-trip byte for byte"
                )
    return target


def build_test_workbook() -> Path:
    return build_workbook(
        OUTPUT / "redim_tests.xlsm",
        {
            "TestReDimCore": (
                read_vba(TESTS_VBA / "TestReDimCore.bas"),
                VBAModuleKind.standard,
            ),
            "TestReDimAsync": (
                read_vba(TESTS_VBA / "TestReDimAsync.bas"),
                VBAModuleKind.standard,
            ),
            "TestReDimWidgets": (
                read_vba(TESTS_VBA / "TestReDimWidgets.bas"),
                VBAModuleKind.standard,
            ),
        },
    )


DEMOS = {
    "ReDim_Mission_Control.xlsm": "MissionControl",
    "ReDim_Widget_Gallery.xlsm": "WidgetGallery",
    "ReDim_Snake.xlsm": "SnakeGame",
    "ReDim_Navigator.xlsm": "Navigator",
}


def build_demo_workbooks() -> list[Path]:
    built = []
    for workbook_name, module_name in DEMOS.items():
        built.append(
            build_workbook(
                ROOT / "demo" / workbook_name,
                {
                    module_name: (
                        read_vba(DEMO_VBA / f"{module_name}.bas"),
                        VBAModuleKind.standard,
                    ),
                },
            )
        )
    return built


if __name__ == "__main__":
    print(build_test_workbook())
    for path in build_demo_workbooks():
        print(path)
