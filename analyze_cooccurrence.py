import argparse
import json
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font


@dataclass(frozen=True)
class TTPRecord:
    ttp_id: str
    ttp_name: str
    tactic: str | None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a TTP co-occurrence spreadsheet from a JSON dataset",
    )
    parser.add_argument(
        "dataset",
        help="Path to the input dataset JSON file",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Path to the output .xlsx file. Defaults to <dataset stem>-cooccurrence.xlsx",
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
        ttp_id = goal.get("ttp_id")
        ttp_name = goal.get("ttp_name")
        tactic = goal.get("tactic")
        if isinstance(ttp_id, str) and isinstance(ttp_name, str):
            normalized_id = normalize_id(ttp_id)
            records[normalized_id] = TTPRecord(
                normalized_id,
                ttp_name,
                tactic if isinstance(tactic, str) else None,
            )

    ttps = report.get("ttps", [])
    if isinstance(ttps, list):
        for item in ttps:
            if not isinstance(item, dict):
                continue
            ttp_id = item.get("ttp_id")
            ttp_name = item.get("ttp_name")
            tactic = item.get("tactic")
            if isinstance(ttp_id, str) and isinstance(ttp_name, str):
                normalized_id = normalize_id(ttp_id)
                records[normalized_id] = TTPRecord(
                    normalized_id,
                    ttp_name,
                    tactic if isinstance(tactic, str) else None,
                )

    return list(records.values())


@dataclass
class CooccurrenceData:
    # pair_counts[a][b] = number of reports containing both TTP a and TTP b
    pair_counts: dict[str, dict[str, int]]
    # ttp_counts[a] = number of reports containing TTP a
    ttp_counts: dict[str, int]
    # tactic_ttp_counts[tactic][ttp_b] = number of reports that contain at least
    #   one TTP from `tactic` AND also contain TTP ttp_b
    tactic_ttp_counts: dict[str, dict[str, int]]
    # tactic_counts[tactic] = number of reports containing at least one TTP from tactic
    tactic_counts: dict[str, int]
    ttp_names: dict[str, str]
    # ttp_tactic[ttp_id] = tactic name (or None)
    ttp_tactic: dict[str, str | None]


def build_cooccurrence_data(reports: list[dict[str, object]]) -> CooccurrenceData:
    pair_counts: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    ttp_counts: dict[str, int] = defaultdict(int)
    tactic_ttp_counts: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    tactic_counts: dict[str, int] = defaultdict(int)
    ttp_names: dict[str, str] = {}
    ttp_tactic: dict[str, str | None] = {}

    for report in reports:
        records = extract_ttp_records(report)
        ttp_ids: set[str] = set()
        tactics_in_report: set[str] = set()

        for record in records:
            ttp_names.setdefault(record.ttp_id, record.ttp_name)
            ttp_tactic.setdefault(record.ttp_id, record.tactic)
            ttp_ids.add(record.ttp_id)
            if record.tactic:
                tactics_in_report.add(record.tactic)

        for ttp_b in ttp_ids:
            ttp_counts[ttp_b] += 1
            for ttp_a in ttp_ids:
                pair_counts[ttp_a][ttp_b] += 1

        for tactic in tactics_in_report:
            tactic_counts[tactic] += 1
            for ttp_b in ttp_ids:
                tactic_ttp_counts[tactic][ttp_b] += 1

    return CooccurrenceData(
        pair_counts, ttp_counts, tactic_ttp_counts, tactic_counts, ttp_names, ttp_tactic
    )


def ordered_ttps(ttp_names: dict[str, str]) -> list[str]:
    return sorted(ttp_names, key=lambda ttp_id: (ttp_id, ttp_names[ttp_id].lower()))


def compute_individual_probability(
    data: CooccurrenceData, ttp_a: str, ttp_b: str
) -> float:
    """P(ttp_a appears | ttp_b appears)"""
    denominator = data.ttp_counts.get(ttp_b, 0)
    if denominator == 0:
        return 0.0
    return data.pair_counts.get(ttp_a, {}).get(ttp_b, 0) / denominator


def compute_tactic_given_tactic_probability(
    data: CooccurrenceData,
    row_tactic: str,
    col_ttp: str,
) -> float:
    """P(at least one TTP from row_tactic appears | col_ttp appears)
    """
    denominator = data.ttp_counts.get(col_ttp, 0)
    if denominator == 0:
        return 0.0
    numerator = data.tactic_ttp_counts.get(row_tactic, {}).get(col_ttp, 0)
    return numerator / denominator


def sheet_name_for_tactic(tactic: str) -> str:
    invalid = r"\/*?:[]{}"
    cleaned = "".join(c if c not in invalid else "_" for c in tactic)
    return cleaned[:31]


def write_matrix(
    sheet,
    row_labels: list[str],
    col_labels: list[str],
    value_getter,  # callable(row_label, col_label) -> float
    number_format: str,
    row_header: str = "TTP A \\ TTP B",
) -> None:
    header_font = Font(bold=True)
    sheet.freeze_panes = "B2"
    sheet.append([row_header, *col_labels])
    for cell in sheet[1]:
        cell.font = header_font

    for row_label in row_labels:
        row = [row_label]
        for col_label in col_labels:
            row.append(value_getter(row_label, col_label))
        sheet.append(row)

    for row in sheet.iter_rows(min_row=2, min_col=2):
        for cell in row:
            cell.number_format = number_format

    sheet.auto_filter.ref = sheet.dimensions
    sheet.column_dimensions["A"].width = 28
    for column_cells in sheet.iter_cols(min_col=2, max_col=len(col_labels) + 1):
        sheet.column_dimensions[column_cells[0].column_letter].width = 22


def write_workbook(output_path: Path, data: CooccurrenceData) -> None:
    workbook = Workbook()
    all_ttp_ids = ordered_ttps(data.ttp_names)

    individual_sheet = workbook.active
    assert individual_sheet is not None
    individual_sheet.title = "individual"

    col_labels = [f"{t}: {data.ttp_names[t]}" for t in all_ttp_ids]

    write_matrix(
        sheet=individual_sheet,
        row_labels=[f"{t}: {data.ttp_names[t]}" for t in all_ttp_ids],
        col_labels=col_labels,
        value_getter=lambda a_label, b_label: compute_individual_probability(
            data,
            a_label.split(":")[0].strip(),
            b_label.split(":")[0].strip(),
        ),
        number_format="0.0000",
        row_header="TTP A \\ TTP B",
    )

    ttps_by_tactic: dict[str, list[str]] = defaultdict(list)
    for ttp_id in all_ttp_ids:
        tactic = data.ttp_tactic.get(ttp_id)
        if tactic:
            ttps_by_tactic[tactic].append(ttp_id)

    sorted_tactics = sorted(ttps_by_tactic)

    for tactic in sorted_tactics:
        tactic_ttps = ttps_by_tactic[tactic]
        sheet = workbook.create_sheet(title=sheet_name_for_tactic(tactic))

        tactic_col_labels = [f"{t}: {data.ttp_names[t]}" for t in tactic_ttps]
        tactic_row_labels = sorted_tactics

        write_matrix(
            sheet=sheet,
            row_labels=tactic_row_labels,
            col_labels=tactic_col_labels,
            value_getter=lambda row_tactic, col_label: (
                compute_tactic_given_tactic_probability(
                    data,
                    row_tactic,
                    col_label.split(":")[0].strip(),
                )
            ),
            number_format="0.0000",
            row_header=f"Tactic \\ {tactic} TTPs",
        )

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
    data = build_cooccurrence_data(reports)
    write_workbook(output_path, data)

    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
