import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator


@dataclass(frozen=True)
class TTPRecord:
    id: str
    name: str
    tactic: str


@dataclass
class Report:
    _raw: dict[str, object]

    @property
    def ttps(self) -> list[TTPRecord]:
        return list(iter_report_ttps(self._raw))

    # Convenience passthrough for common report fields
    def get(self, key: str, default=None):
        return self._raw.get(key, default)


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
) -> Iterator[TTPRecord]:
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


def iter_ttps(path: Path) -> Iterator[Report]:
    """Yield one Report per entry in the dataset file."""
    for raw in load_reports(path):
        yield Report(_raw=raw)


def iter_consecutive_pairs(
    reports: list[dict[str, object]],
    *,
    include_goal: bool = False,
) -> Iterator[tuple[TTPRecord, TTPRecord | None]]:
    for report in reports:
        sequence = list(iter_report_ttps(report, include_goal=include_goal))
        for i, current in enumerate(sequence):
            nxt = sequence[i + 1] if i + 1 < len(sequence) else None
            yield current, nxt
