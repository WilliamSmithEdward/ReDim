"""Performance bench for ReDim: framework scenarios plus demo build times.

Runs tests/vba/BenchReDim.bas in a workbook built from the current sources,
then opens each demo workbook and times its build procedure three ways: the
session's first build (cold, carrying one-time costs), a rebuild that adopts
the existing shapes, and a warm build that draws everything again from
empty sheets. Prints milliseconds per metric. With --save NAME the results
land in tests/output/bench_NAME.json; with --compare NAME the table adds
that run's numbers and the change.

Each scenario runs --repeat times (default 3) in one Excel session and the
minimum per metric is kept, since interference only adds time. Every run
starts from a framework with no apps, so the first run carries the
session's one-time costs and the minimum reports the warm cost. Compare
runs taken with the same --repeat.

Timings come from a live, usually hidden Excel instance, so they carry the
object-model cost without screen painting; compare runs from the same
machine, not absolute numbers across machines.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pyopenvba import VBAModuleKind
from pyvbaharness import ExcelSession

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_workbooks import DEMOS, build_demo_workbooks, build_workbook  # noqa: E402
from vba_sources import OUTPUT, ROOT, TESTS_VBA, read_vba  # noqa: E402

SCENARIOS = [
    "BenchRender",
    "BenchTyping",
    "BenchLists",
    "BenchItems",
    "BenchPump",
    "BenchNavigate",
    "BenchComponents",
]

BUILD_PROCS = {
    "ReDim_Mission_Control.xlsm": "BuildMissionControl",
    "ReDim_Widget_Gallery.xlsm": "BuildWidgetGallery",
    "ReDim_Snake.xlsm": "BuildSnake",
    "ReDim_Navigator.xlsm": "BuildNavigator",
    "ReDim_ReDex.xlsm": "BuildPokeDex",
}

TIME_BUILD = """
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
    ByRef counter As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
    ByRef ticksPerSecond As LongLong) As Long

Private Function NowMs() As Double
    Static ticksPerSecond As LongLong
    Dim counter As LongLong

    If ticksPerSecond = 0 Then QueryPerformanceFrequency ticksPerSecond
    QueryPerformanceCounter counter
    NowMs = counter * 1000# / ticksPerSecond
End Function

Public Function TimeBuild(ByVal buildProc As String) As String
    Dim started As Double
    Dim coldMs As Double
    Dim rebuildMs As Double
    Dim sheetValue As Worksheet
    Dim idx As Long

    ReDimUI.AutoPump False
    started = NowMs()
    Application.Run buildProc
    coldMs = NowMs() - started
    started = NowMs()
    Application.Run buildProc
    rebuildMs = NowMs() - started
    ' The warm build draws everything again from sheets without shapes, in a
    ' session that has already paid the one-time costs the cold build carries.
    ReDimUI.Shutdown
    For Each sheetValue In ThisWorkbook.Worksheets
        sheetValue.Unprotect
        For idx = sheetValue.Shapes.Count To 1 Step -1
            sheetValue.Shapes(idx).Delete
        Next idx
    Next sheetValue
    ReDimUI.AutoPump False
    started = NowMs()
    Application.Run buildProc
    TimeBuild = "build=" & Format$(coldMs, "0.000") & _
        "|rebuild=" & Format$(rebuildMs, "0.000") & _
        "|warmBuild=" & Format$(NowMs() - started, "0.000")
    ReDimUI.Shutdown
End Function
"""


def parse(raw: str) -> dict[str, float]:
    return {key: float(value) for key, value in
            (token.split("=", 1) for token in raw.split("|"))}


def run_scenarios(repeat: int) -> dict[str, dict[str, float]]:
    """Each scenario runs `repeat` times; the minimum per metric is kept,
    since interference from other work on the machine only adds time."""
    workbook = build_workbook(
        OUTPUT / "bench.xlsm",
        {"BenchReDim": (read_vba(TESTS_VBA / "BenchReDim.bas"),
                        VBAModuleKind.standard)},
    )
    results: dict[str, dict[str, float]] = {}
    with ExcelSession() as excel:
        excel.open_workbook(str(workbook))
        for scenario in SCENARIOS:
            for _ in range(repeat):
                outcome = excel.run_macro(f"BenchReDim.{scenario}", timeout=600)
                if outcome.outcome != "passed":
                    raise RuntimeError(
                        f"{scenario}: {outcome.outcome} {outcome.error}")
                best = results.setdefault(scenario, {})
                for metric, value in parse(str(outcome.value)).items():
                    best[metric] = min(value, best.get(metric, value))
    return results


def run_demo_builds() -> dict[str, dict[str, float]]:
    results: dict[str, dict[str, float]] = {}
    for path in build_demo_workbooks():
        with ExcelSession() as excel:
            excel.open_workbook(str(path))
            outcome = excel.run_vba(TIME_BUILD, proc="TimeBuild",
                                    args=(BUILD_PROCS[path.name],), timeout=600)
            if outcome.outcome != "passed":
                raise RuntimeError(f"{path.name}: {outcome.outcome} {outcome.error}")
            results[DEMOS[path.name]] = parse(str(outcome.value))
    return results


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    parser.add_argument("--save", help="store results as tests/output/bench_NAME.json")
    parser.add_argument("--compare", help="compare with tests/output/bench_NAME.json")
    parser.add_argument("--skip-demos", action="store_true")
    parser.add_argument("--repeat", type=int, default=3,
                        help="runs per scenario; the minimum is reported")
    options = parser.parse_args()

    results = run_scenarios(options.repeat)
    if not options.skip_demos:
        results.update({f"demo:{name}": data
                        for name, data in run_demo_builds().items()})
    baseline: dict[str, dict[str, float]] = {}
    if options.compare:
        baseline = json.loads(
            (OUTPUT / f"bench_{options.compare}.json").read_text(encoding="utf-8"))
    header = f"{'scenario':<16}{'metric':<30}{'ms':>12}"
    if baseline:
        header += f"{'was':>12}{'change':>9}"
    print(header)
    for scenario, metrics in results.items():
        for metric, value in metrics.items():
            line = f"{scenario:<16}{metric:<30}{value:>12.3f}"
            before = baseline.get(scenario, {}).get(metric)
            if before is not None:
                change = (value - before) / before * 100 if before else 0.0
                line += f"{before:>12.3f}{change:>8.0f}%"
            print(line)
    if options.save:
        target = OUTPUT / f"bench_{options.save}.json"
        target.write_text(json.dumps(results, indent=1), encoding="utf-8")
        print(f"saved {target.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
