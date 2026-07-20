#!/usr/bin/env python3
"""Generate the weekly EDA Jira ticket categorization workbook."""

from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


STATUS_ORDER = ("Open", "In Progress", "Resolved", "Closed")
STATUS_COLORS = {
    "Open": "D9EAF7",
    "In Progress": "FFF2CC",
    "Resolved": "E2F0D9",
    "Closed": "D9D9D9",
}
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SUBHEADER_FILL = PatternFill("solid", fgColor="5B9BD5")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
WHITE_BOLD_FONT = Font(color="FFFFFF", bold=True)
THIN_BLUE_BORDER = Border(
    left=Side(style="thin", color="B4C6E7"),
    right=Side(style="thin", color="B4C6E7"),
    top=Side(style="thin", color="B4C6E7"),
    bottom=Side(style="thin", color="B4C6E7"),
)


def value_name(value: Any) -> str:
    """Return a display value from a Jira field or normalized scalar."""
    if value is None:
        return ""
    if isinstance(value, dict):
        return str(value.get("name") or value.get("value") or "")
    return str(value)


def joined_names(value: Any) -> str:
    """Return comma-separated names from Jira list fields."""
    if not value:
        return ""
    if isinstance(value, str):
        return value
    return ", ".join(value_name(item) for item in value if value_name(item))


def normalize_status(jira_status: str, status_category_key: str) -> str:
    """Map Jira statuses into the four report status groups."""
    if jira_status.casefold() == "closed":
        return "Closed"
    if status_category_key == "done":
        return "Resolved"
    if status_category_key == "indeterminate":
        return "In Progress"
    return "Open"


def normalize_issue(issue: dict[str, Any]) -> dict[str, str]:
    """Normalize either a Jira API issue or a pre-normalized issue object."""
    fields = issue.get("fields", issue)
    status = fields.get("status") or {}
    issue_type = value_name(fields.get("issuetype") or fields.get("issue_type"))
    jira_status = value_name(status) or str(fields.get("jira_status") or "")
    category_key = (
        status.get("statusCategory", {}).get("key", "")
        if isinstance(status, dict)
        else ""
    ) or str(fields.get("status_category_key") or "")
    category = value_name(fields.get("customfield_19840") or fields.get("category"))

    return {
        "key": str(issue.get("key") or fields.get("key") or ""),
        "summary": str(fields.get("summary") or ""),
        "status_group": str(
            fields.get("normalized_status")
            or normalize_status(jira_status, category_key)
        ),
        "jira_status": jira_status,
        # EDA's Jira Category custom field is currently unpopulated. Issue Type
        # provides a stable, meaningful category without losing any tickets.
        "category": category or issue_type or "Uncategorized",
        "issue_type": issue_type,
        "priority": value_name(fields.get("priority")),
        "assignee": value_name(fields.get("assignee")),
        "reporter": value_name(fields.get("reporter")),
        "created": str(fields.get("created") or ""),
        "updated": str(fields.get("updated") or ""),
        "resolution": value_name(fields.get("resolution")),
        "components": joined_names(fields.get("components")),
        "labels": joined_names(fields.get("labels")),
    }


def load_issues(path: Path) -> list[dict[str, str]]:
    """Load and normalize issues from JSON."""
    with path.open(encoding="utf-8") as source:
        payload = json.load(source)
    issues = payload.get("issues", []) if isinstance(payload, dict) else payload
    normalized = [normalize_issue(issue) for issue in issues]
    if not normalized:
        raise ValueError("The input contains no Jira issues")
    if any(not issue["key"] for issue in normalized):
        raise ValueError("Every Jira issue must have a key")
    return sorted(normalized, key=lambda issue: issue["key"])


def style_header(ws: Any, row: int, start: int, end: int) -> None:
    for cell in ws.iter_cols(
        min_col=start, max_col=end, min_row=row, max_row=row
    ):
        header = cell[0]
        header.fill = HEADER_FILL
        header.font = WHITE_BOLD_FONT
        header.alignment = Alignment(horizontal="center", vertical="center")
        header.border = THIN_BLUE_BORDER


def build_raw_sheet(wb: Workbook, issues: list[dict[str, str]]) -> None:
    ws = wb.active
    ws.title = "Raw Tickets"
    headers = [
        "Ticket",
        "Summary",
        "Status Group",
        "Jira Status",
        "Category",
        "Issue Type",
        "Priority",
        "Assignee",
        "Reporter",
        "Created",
        "Updated",
        "Resolution",
        "Components",
        "Labels",
    ]
    ws.append(headers)
    for issue in issues:
        ws.append(
            [
                issue["key"],
                issue["summary"],
                issue["status_group"],
                issue["jira_status"],
                issue["category"],
                issue["issue_type"],
                issue["priority"],
                issue["assignee"],
                issue["reporter"],
                issue["created"],
                issue["updated"],
                issue["resolution"],
                issue["components"],
                issue["labels"],
            ]
        )

    style_header(ws, 1, 1, len(headers))
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    widths = [15, 70, 16, 16, 18, 15, 16, 31, 31, 25, 25, 16, 28, 24]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + index)].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = THIN_BLUE_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        key_cell = row[0]
        key_cell.hyperlink = f"https://wbdprod.atlassian.net/browse/{key_cell.value}"
        key_cell.style = "Hyperlink"

    table = Table(displayName="RawTickets", ref=ws.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)

    for status, color in STATUS_COLORS.items():
        ws.conditional_formatting.add(
            f"C2:C{ws.max_row}",
            FormulaRule(
                formula=[f'$C2="{status}"'],
                fill=PatternFill("solid", fgColor=color),
            ),
        )


def build_summary_sheet(
    wb: Workbook, issues: list[dict[str, str]], report_date: date
) -> None:
    ws = wb.create_sheet("Summary", 0)
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:F1")
    ws["A1"] = "EDA Ticket Categorization Summary"
    ws["A1"].fill = HEADER_FILL
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=16)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    assignees = sorted({issue["assignee"] for issue in issues if issue["assignee"]})
    metadata = [
        ("Report week", report_date.isoformat()),
        ("Assignee", ", ".join(assignees)),
        ("Total tickets", len(issues)),
        ("Category basis", "Jira Category; Issue Type when Category is empty"),
    ]
    for row_number, (label, value) in enumerate(metadata, start=3):
        ws.cell(row_number, 1, label).font = Font(bold=True, color="1F4E78")
        ws.cell(row_number, 2, value)

    matrix_start = 9
    matrix_headers = ["Category", *STATUS_ORDER, "Total"]
    for column, header in enumerate(matrix_headers, start=1):
        ws.cell(matrix_start, column, header)
    style_header(ws, matrix_start, 1, len(matrix_headers))

    counts: dict[str, Counter[str]] = defaultdict(Counter)
    for issue in issues:
        counts[issue["category"]][issue["status_group"]] += 1

    row_number = matrix_start + 1
    for category in sorted(counts):
        ws.cell(row_number, 1, category)
        for column, status in enumerate(STATUS_ORDER, start=2):
            ws.cell(row_number, column, counts[category][status])
        ws.cell(
            row_number,
            len(matrix_headers),
            sum(counts[category].values()),
        )
        row_number += 1

    total_row = row_number
    ws.cell(total_row, 1, "Total")
    overall = Counter(issue["status_group"] for issue in issues)
    for column, status in enumerate(STATUS_ORDER, start=2):
        ws.cell(total_row, column, overall[status])
    ws.cell(total_row, len(matrix_headers), len(issues))

    for row in ws.iter_rows(
        min_row=matrix_start + 1,
        max_row=total_row,
        min_col=1,
        max_col=len(matrix_headers),
    ):
        for cell in row:
            cell.border = THIN_BLUE_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row[0].alignment = Alignment(horizontal="left", vertical="center")
    for cell in ws[total_row]:
        if cell.column <= len(matrix_headers):
            cell.fill = TOTAL_FILL
            cell.font = Font(bold=True)

    status_start = total_row + 3
    ws.cell(status_start, 1, "Status")
    ws.cell(status_start, 2, "Count")
    style_header(ws, status_start, 1, 2)
    for offset, status in enumerate(STATUS_ORDER, start=1):
        row = status_start + offset
        ws.cell(row, 1, status)
        ws.cell(row, 2, overall[status])
        for cell in ws[row][:2]:
            cell.border = THIN_BLUE_BORDER
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=STATUS_COLORS[status])

    ws.freeze_panes = "A9"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 19
    ws.column_dimensions["C"].width = 19
    ws.column_dimensions["D"].width = 19
    ws.column_dimensions["E"].width = 19
    ws.column_dimensions["F"].width = 14


def generate_report(input_path: Path, output_path: Path, report_date: date) -> None:
    issues = load_issues(input_path)
    wb = Workbook()
    wb.properties.title = "EDA Ticket Categorization"
    wb.properties.subject = f"Weekly Jira report for {report_date.isoformat()}"
    wb.properties.creator = "EDA Jira report automation"
    build_raw_sheet(wb, issues)
    build_summary_sheet(wb, issues, report_date)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Jira issue JSON input")
    parser.add_argument("output", type=Path, help="Output .xlsx path")
    parser.add_argument(
        "--report-date",
        type=date.fromisoformat,
        default=date.today(),
        help="Report week date in YYYY-MM-DD format (default: today)",
    )
    return parser.parse_args()


if __name__ == "__main__":
    arguments = parse_args()
    generate_report(arguments.input, arguments.output, arguments.report_date)
