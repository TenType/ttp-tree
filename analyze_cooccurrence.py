import argparse
import json
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Font


@dataclass(frozen=True)
class TTPRecord:
    ttp_id: str
    ttp_name: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a TTP co-occurrence spreadsheet from a JSON dataset.",
    )
    parser.add_argument(
        "dataset",
        nargs="?",
        default="data/gti-235-linux.json",
        help="Path to the input dataset JSON file.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Path to the output .xlsx file. Defaults to <dataset stem>-cooccurrence.xlsx.",
    )
    return parser.parse_args()


def normalize_id(value: str) -> str:
    return value.strip().upper()


def load_reports(dataset_path: Path) -> list[dict[str, object]]:
    with dataset_path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, list):
        raise ValueError("Expected the dataset root to be a JSON array of reports.")
    return data


def extract_ttp_records(report: dict[str, object]) -> list[TTPRecord]:
    records: dict[str, TTPRecord] = {}

    goal = report.get("goal")
    if isinstance(goal, dict):
        goal_id = goal.get("ttp_id")
        goal_name = goal.get("ttp_name")
        if isinstance(goal_id, str) and isinstance(goal_name, str):
            normalized_goal_id = normalize_id(goal_id)
            records[normalized_goal_id] = TTPRecord(normalized_goal_id, goal_name)

    ttps = report.get("ttps", [])
    if isinstance(ttps, list):
        for item in ttps:
            if not isinstance(item, dict):
                continue
            ttp_id = item.get("ttp_id")
            ttp_name = item.get("ttp_name")
            if isinstance(ttp_id, str) and isinstance(ttp_name, str):
                normalized_ttp_id = normalize_id(ttp_id)
                records[normalized_ttp_id] = TTPRecord(normalized_ttp_id, ttp_name)

    return list(records.values())


def build_cooccurrence_tables(
    reports: Iterable[dict[str, object]],
) -> tuple[dict[str, dict[str, int]], dict[str, int], dict[str, str]]:
    pair_counts: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    denominator_counts: dict[str, int] = defaultdict(int)
    ttp_names: dict[str, str] = {}

    for report in reports:
        records = extract_ttp_records(report)
        ttp_ids = sorted({record.ttp_id for record in records})

        for record in records:
            ttp_names.setdefault(record.ttp_id, record.ttp_name)

        for ttp_b in ttp_ids:
            denominator_counts[ttp_b] += 1
            for ttp_a in ttp_ids:
                pair_counts[ttp_a][ttp_b] += 1

    return pair_counts, denominator_counts, ttp_names


def ordered_ttps(ttp_names: dict[str, str]) -> list[str]:
    return sorted(ttp_names, key=lambda ttp_id: (ttp_id, ttp_names[ttp_id].lower()))


def write_workbook(
    output_path: Path,
    pair_counts: dict[str, dict[str, int]],
    denominator_counts: dict[str, int],
    ttp_names: dict[str, str],
) -> None:
    workbook = Workbook()
    probability_sheet = workbook.active
    assert probability_sheet is not None
    probability_sheet.title = "probabilities"
    counts_sheet = workbook.create_sheet("counts")

    ttp_ids = ordered_ttps(ttp_names)
    header_font = Font(bold=True)

    def write_matrix(sheet, value_getter, number_format: str) -> None:
        sheet.freeze_panes = "B2"
        sheet.append(["TTP A \\ TTP B", *[f"{ttp_id}: {ttp_names[ttp_id]}" for ttp_id in ttp_ids]])
        for cell in sheet[1]:
            cell.font = header_font

        for ttp_a in ttp_ids:
            row = [f"{ttp_a}: {ttp_names[ttp_a]}"]
            for ttp_b in ttp_ids:
                row.append(value_getter(ttp_a, ttp_b))
            sheet.append(row)

        for row in sheet.iter_rows(min_row=2, min_col=2):
            for cell in row:
                cell.number_format = number_format

        sheet.auto_filter.ref = sheet.dimensions
        sheet.column_dimensions["A"].width = 28
        for column_cells in sheet.iter_cols(min_col=2, max_col=len(ttp_ids) + 1):
            column_letter = column_cells[0].column_letter
            sheet.column_dimensions[column_letter].width = 22

    def probability_value(ttp_a: str, ttp_b: str) -> float:
        denominator = denominator_counts.get(ttp_b, 0)
        if denominator == 0:
            return 0.0
        return pair_counts.get(ttp_a, {}).get(ttp_b, 0) / denominator

    def count_value(ttp_a: str, ttp_b: str) -> int:
        return pair_counts.get(ttp_a, {}).get(ttp_b, 0)

    write_matrix(probability_sheet, probability_value, "0.0000")
    write_matrix(counts_sheet, count_value, "0")

    workbook.save(output_path)


def main() -> None:
    args = parse_args()
    dataset_path = Path(args.dataset).expanduser().resolve()
    if not dataset_path.exists():
        raise FileNotFoundError(f"Dataset not found: {dataset_path}")

    output_path = (
        Path(args.output).expanduser().resolve()
        if args.output
        else dataset_path.with_name(f"{dataset_path.stem}-cooccurrence.xlsx")
    )

    reports = load_reports(dataset_path)
    pair_counts, denominator_counts, ttp_names = build_cooccurrence_tables(reports)
    write_workbook(output_path, pair_counts, denominator_counts, ttp_names)

    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
