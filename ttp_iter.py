import json
from dataclasses import dataclass
from pathlib import Path
from typing import Generator


@dataclass(frozen=True)
class TTPRecord:
    ttp_id: str
    ttp_name: str
    tactic: str | None


def normalize_id(value: str) -> str:
    return value.strip().upper()


def load_reports(dataset_path: Path) -> list[dict[str, object]]:
    with dataset_path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, list):
        raise ValueError("Expected the dataset root to be a JSON array of reports")
    return data


def _parse_record(item: dict[str, str]) -> TTPRecord | None:
    ttp_id = item["ttp_id"]
    ttp_name = item["ttp_name"]
    tactic = item["tactic"]
    return TTPRecord(
        normalize_id(ttp_id),
        ttp_name,
        tactic,
    )


def iter_report_ttps(
    report: dict[str, object],
    *,
    include_goal: bool = False,
) -> Generator[TTPRecord, None, None]:
    if include_goal:
        goal = report.get("goal")
        if isinstance(goal, dict):
            record = _parse_record(goal)
            if record is not None:
                yield record

    ttps = report.get("ttps", [])
    if isinstance(ttps, list):
        for item in ttps:
            record = _parse_record(item)
            if record is not None:
                yield record


def iter_all_ttps(
    reports: list[dict[str, object]],
    *,
    include_goal: bool = False,
) -> Generator[tuple[int, TTPRecord], None, None]:
    for idx, report in enumerate(reports):
        for record in iter_report_ttps(report, include_goal=include_goal):
            yield idx, record


def iter_consecutive_pairs(
    reports: list[dict[str, object]],
    *,
    include_goal: bool = False,
) -> Generator[tuple[TTPRecord, TTPRecord | None], None, None]:
    for report in reports:
        sequence = list(iter_report_ttps(report, include_goal=include_goal))
        for i, current in enumerate(sequence):
            nxt = sequence[i + 1] if i + 1 < len(sequence) else None
            yield current, nxt
