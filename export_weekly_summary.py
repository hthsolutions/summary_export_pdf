"""Export the Revel Summary page for the current week to Excel and PDF.

The week matches the Summary page: the trailing 7 days ending on
yesterday in Central Time, kept inside the dates Supabase actually
has. Day columns use that ending business date.

Credentials are the same ones the Streamlit app uses. Set
SUPABASE_URL and SUPABASE_KEY in the environment, or keep them in
revel-demo/.streamlit/secrets.toml.

Usage:
    py -3 export_weekly_summary.py
    py -3 export_weekly_summary.py --date 2026-10-26
"""

from __future__ import annotations

import argparse
import ast
import logging
import os
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

REPO_ROOT = Path(__file__).resolve().parent
REVEL_DEMO = Path(
    os.environ.get("REVEL_DEMO", REPO_ROOT.parent / "revel-demo")
)
SUMMARY_SOURCE = REVEL_DEMO / "views" / "summary.py"

SKIP_FUNCTIONS = {
    "sales_pie",
    "leaderboard_column_config",
    "summary_column_config",
    "table_height",
}

# Same cutoff as the Summary page pie: smaller slices keep
# their percent in the legend only.
SLICE_LABEL_MIN_SHARE = 0.035

# Altair's category20 palette, so the slices match the page.
CATEGORY_20 = [
    "1F77B4",
    "AEC7E8",
    "FF7F0E",
    "FFBB78",
    "2CA02C",
    "98DF8A",
    "D62728",
    "FF9896",
    "9467BD",
    "C5B0D5",
    "8C564B",
    "C49C9C",
    "E377C2",
    "F7B6D2",
    "7F7F7F",
    "C7C7C7",
    "BCBD22",
    "DBDB8D",
    "17BECF",
    "9EDAE5",
]

PIE_CAPTION = (
    "Each slice is one store's net sales for these 7 days. "
    "The percentage is drawn on the slice, and again "
    "beside the store name. A slice under "
    f"{SLICE_LABEL_MIN_SHARE * 100:.1f}% keeps the percentage "
    "in the legend so the labels stay apart."
)

NAVY = "1F2933"
INK = "1F2933"
MUTED = "667085"
ZEBRA = "F7F8FA"
TOTAL = "E8EEF4"
AMBER = "FFF7E8"
AMBER_INK = "7A4E00"
LINE = "D0D5DD"
WHITE = "FFFFFF"

# Same cutoffs as the Summary page. At or under is green.
HOURLY_LABOR_CUTOFF = 18.0
LANE_GOAL_SECONDS = 4 * 60
GOOD_FILL = "D7EAE0"
GOOD_INK = "157347"
BAD_FILL = "F3DCDA"
BAD_INK = "B42318"

EXCEL_FORMATS = {
    "money0": '"$"#,##0',
    "money2": '"$"#,##0.00',
    "points": '0.0"%"',
    "share": "0.0%",
    "num1": "0.0",
    "score": "0.00",
    "delta2": "+0.00;-0.00;0.00",
    "delta0": "+0;-0;0",
    "int": "0",
}


@dataclass(frozen=True)
class Column:
    """One export column: frame key, header, and number format."""

    key: str
    header: str
    kind: str
    width: float
    comment: str = ""


SALES_COLUMNS = [
    Column("location", "Store", "text", 24),
    Column(
        "day_net_sales",
        "Net sales (day)",
        "money0",
        15,
        "Net sales on the business date that ends this week.",
    ),
    Column(
        "week_net_sales",
        "Net sales (week)",
        "money0",
        16,
        "Net sales for the 7 days ending on the business date.",
    ),
    Column(
        "week_net_sales_growth",
        "Weekly Growth %",
        "points",
        16,
        "Percent change in net sales versus the previous 7 days. "
        "Green when zero or positive, red when negative.",
    ),
    Column(
        "week_share",
        "Share of week",
        "share",
        13,
        "This store's net sales divided by net sales for every store over those 7 days.",
    ),
    Column(
        "annualized",
        "Annualized",
        "money0",
        15,
        "Net sales for those 7 days, multiplied by 52.",
    ),
    Column(
        "day_labor_percent",
        "Hr Labor % (day)",
        "points",
        15,
        "Hourly shift wages divided by that day's net sales. "
        "Green at 18% or under, red above 18%.",
    ),
    Column(
        "week_labor_percent",
        "Hr Labor % (week)",
        "points",
        16,
        "Hourly shift wages divided by net sales for those 7 days. "
        "Green at 18% or under, red above 18%.",
    ),
    Column(
        "day_splh",
        "SpLH (day)",
        "money2",
        13,
        "That day's net sales divided by hourly labor hours. Salaried roles, including GM, are left out of the hours.",
    ),
    Column(
        "week_splh",
        "SpLH (week)",
        "money2",
        14,
        "Net sales for those 7 days divided by hourly labor hours. Salaried roles, including GM, are left out of the hours.",
    ),
    Column(
        "day_lane",
        "Drive-thru (day)",
        "text",
        16,
        "Car-weighted average lane time that day. "
        "Green at 4:00 or under, red above 4:00.",
    ),
    Column(
        "week_lane",
        "Drive-thru (week)",
        "text",
        17,
        "Car-weighted average lane time over those 7 days. "
        "Green at 4:00 or under, red above 4:00.",
    ),
    Column(
        "day_ot_hours",
        "OT hours (day)",
        "num1",
        14,
        "Overtime hours that day. The GM role is left out.",
    ),
    Column(
        "week_ot_hours",
        "OT hours (week)",
        "num1",
        15,
        "Overtime hours over those 7 days. The GM role is left out.",
    ),
]

REVIEW_COLUMNS = [
    Column("location", "Store", "text", 24),
    Column(
        "week_review_score",
        "Week avg",
        "score",
        12,
        "Average star rating of the reviews posted during those 7 days.",
    ),
    Column(
        "week_review_score_change",
        "Week avg (Change +/-)",
        "delta2",
        18,
        "This week's average minus the average for the previous 7 days. Blank when either week has no scored reviews.",
    ),
    Column(
        "week_review_count",
        "Week",
        "int",
        12,
        "Google reviews posted during those 7 days.",
    ),
    Column(
        "week_review_count_change",
        "Week (change +/-)",
        "delta0",
        16,
        "This week's review count minus the count for the previous 7 days.",
    ),
    Column(
        "day_review_score",
        "Day avg",
        "score",
        12,
        "Average star rating of the reviews posted on the business date.",
    ),
    Column(
        "day_review_count",
        "Day",
        "int",
        10,
        "Google reviews posted on the business date.",
    ),
    Column(
        "current_review_score",
        "Current",
        "score",
        12,
        "Average star rating of every Google review stored for this store.",
    ),
]


@dataclass
class Report:
    """The Summary page for one business date and its trailing week."""

    business_date: pd.Timestamp
    week_start: pd.Timestamp
    week_end: pd.Timestamp
    week_label: str
    recorded_days: int
    generated_at: datetime
    warnings: list[str]
    sales: pd.DataFrame
    reviews: pd.DataFrame
    sales_notes: list[str]
    review_notes: list[str]

    @property
    def heading(self) -> str:
        day = self.business_date
        return (
            f"{day:%A, %b %d, %Y}  ·  "
            f"Trailing 7 days, {self.week_label}"
        )

    @property
    def file_stem(self) -> str:
        start = self.week_start.strftime("%Y-%m-%d")
        end = self.week_end.strftime("%Y-%m-%d")
        return f"Weekly_Summary_{start}_to_{end}"


def unwrap(fn):
    """Call a Streamlit-cached loader without a Streamlit runtime."""

    seen: set[int] = set()
    while hasattr(fn, "__wrapped__") and id(fn) not in seen:
        seen.add(id(fn))
        fn = fn.__wrapped__
    return fn


def load_summary_api() -> dict:
    """Load the Summary calculations without running the Streamlit page."""

    if not SUMMARY_SOURCE.is_file():
        raise SystemExit(
            f"Could not find the Summary page at {SUMMARY_SOURCE}. "
            "Set REVEL_DEMO to the revel-demo repo if it is not the sibling folder."
        )

    tree = ast.parse(SUMMARY_SOURCE.read_text(encoding="utf-8"))
    body = []
    seen_function = False
    for node in tree.body:
        if isinstance(node, ast.Import):
            aliases = [
                alias
                for alias in node.names
                if alias.name not in {"altair", "streamlit"}
            ]
            if not aliases:
                continue
            node.names = aliases
            body.append(node)
        elif isinstance(node, ast.ImportFrom):
            body.append(node)
        elif isinstance(node, ast.FunctionDef):
            seen_function = True
            if node.name not in SKIP_FUNCTIONS:
                body.append(node)
        elif isinstance(node, ast.Assign) and not seen_function:
            body.append(node)

    module = ast.Module(body=body, type_ignores=[])
    ast.fix_missing_locations(module)
    namespace = {
        "__name__": "summary_calculations",
        "__file__": str(SUMMARY_SOURCE),
    }
    exec(compile(module, str(SUMMARY_SOURCE), "exec"), namespace)
    return namespace


def load_source(label: str, fn, fatal: bool):
    """Load one Supabase feed. A missing secret stops the export."""

    from revel_data import MissingSecretError

    try:
        return unwrap(fn)(), None
    except MissingSecretError as error:
        if fatal:
            raise SystemExit(str(error)) from error
        return pd.DataFrame(), f"{label} are unavailable: {error}"
    except Exception as error:
        return pd.DataFrame(), f"{label} are unavailable: {error}"


def build_report(api: dict, requested_date: str | None) -> Report:
    """Assemble the same store comparison the Summary page shows."""

    from revel_data import (
        STORE_TIMEZONE,
        load_google_reviews,
        load_hme_departures,
        load_salary_data,
        load_sales_data,
        load_shift_data,
    )

    warnings: list[str] = []
    sales_df, sales_error = load_source(
        "Sales",
        load_sales_data,
        fatal=True,
    )
    shift_df, shift_error = load_source(
        "Shift wages and hours",
        load_shift_data,
        fatal=True,
    )
    salary_df, salary_error = load_source(
        "Salary roles",
        load_salary_data,
        fatal=False,
    )
    hme_df, hme_error = load_source(
        "Drive-thru times",
        load_hme_departures,
        fatal=True,
    )
    review_df, review_error = load_source(
        "Google reviews",
        load_google_reviews,
        fatal=True,
    )
    for message in (
        sales_error,
        shift_error,
        salary_error,
        hme_error,
        review_error,
    ):
        if message:
            warnings.append(message)

    return assemble_report(
        api,
        sales_df,
        shift_df,
        salary_df,
        hme_df,
        review_df,
        requested_date,
        warnings,
        shifts_loaded=shift_error is None,
        reviews_loaded=review_error is None,
    )


def assemble_report(
    api: dict,
    sales_df: pd.DataFrame,
    shift_df: pd.DataFrame,
    salary_df: pd.DataFrame,
    hme_df: pd.DataFrame,
    review_df: pd.DataFrame,
    requested_date: str | None,
    warnings: list[str],
    shifts_loaded: bool,
    reviews_loaded: bool,
) -> Report:
    """Turn loaded feeds into the Summary page for one week."""

    from revel_data import STORE_TIMEZONE

    sales_daily = api["sales_by_day"](sales_df)
    wage_daily = api["wages_by_day"](shift_df)
    hours_daily = api["hours_by_day"](
        shift_df,
        api["salaried_role_keys"](salary_df),
    )
    ot_daily = api["ot_by_day"](shift_df)
    stores = api["store_catalog"](
        sales_daily,
        api["dated"](shift_df),
        api["dated"](hme_df),
        review_df,
    )
    if stores.empty:
        if not warnings:
            warnings.append(
                "Supabase returned no rows. Use the same "
                "SUPABASE_KEY the Summary app uses; a key that "
                "can sign in but cannot read these tables comes "
                "back empty."
            )
        detail = "\n".join(warnings)
        raise SystemExit(
            "No stores were returned from Supabase. Check the "
            "sales, shift, and drive-thru tables.\n" + detail
        )

    dates = api["available_dates"](
        sales_daily,
        wage_daily,
        api["dated"](hme_df),
        review_df,
    )
    if dates.empty:
        raise SystemExit("No business dates were returned from Supabase.")

    min_date = dates.min().date()
    max_date = dates.max().date()
    if requested_date:
        business_date = datetime.strptime(
            requested_date,
            "%Y-%m-%d",
        ).date()
        if business_date < min_date or business_date > max_date:
            raise SystemExit(
                f"{business_date:%Y-%m-%d} is outside the loaded "
                f"dates ({min_date:%Y-%m-%d} through {max_date:%Y-%m-%d})."
            )
    else:
        business_date = api["default_summary_date"](min_date, max_date)

    day = pd.Timestamp(business_date).normalize()
    week_end = day
    week_start = day - pd.Timedelta(days=6)

    day_metrics = api["store_window"](
        stores,
        sales_daily,
        wage_daily,
        hours_daily,
        ot_daily,
        hme_df,
        day,
        day,
        annualize=False,
        shifts_loaded=shifts_loaded,
    )
    week_metrics = api["store_window"](
        stores,
        sales_daily,
        wage_daily,
        hours_daily,
        ot_daily,
        hme_df,
        week_start,
        week_end,
        annualize=True,
        shifts_loaded=shifts_loaded,
    )
    comparison = api["combine_periods"](day_metrics, week_metrics)
    comparison = comparison.merge(
        api["week_net_sales_growth"](
            week_metrics,
            sales_daily,
            week_start,
            week_end,
        ),
        on="row_key",
        how="left",
    )
    comparison = comparison.merge(
        api["review_metrics"](
            review_df,
            stores,
            day,
            week_start,
            week_end,
            reviews_loaded=reviews_loaded,
        ),
        on="row_key",
        how="left",
    )

    total_sales = comparison.loc[
        comparison["row_key"].eq("ALL"),
        "week_net_sales",
    ]
    company_sales = (
        float(total_sales.iloc[0])
        if not total_sales.empty and pd.notna(total_sales.iloc[0])
        else float("nan")
    )
    if pd.notna(company_sales) and company_sales > 0:
        comparison["week_share"] = (
            pd.to_numeric(comparison["week_net_sales"], errors="coerce")
            / company_sales
        )
    else:
        comparison["week_share"] = float("nan")

    week_sales = api["in_window"](sales_daily, week_start, week_end)
    recorded_days = (
        int(week_sales["business_date"].nunique())
        if not week_sales.empty
        else 0
    )

    sales_notes = [
        "Stores follow opening order: Marble Falls, Lampasas, "
        "Leander, then Temple. "
        "Weekly growth % compares these 7 days with the 7 days "
        "before them. Zero or positive growth is green, and "
        "negative growth is red. "
        "Hr labor % is hourly shift wages divided by net sales. "
        "18% or under is green, and above 18% is red. "
        "SpLH is net sales divided by hourly labor hours, with "
        "salaried roles such as GM left out of the hours. "
        "Drive-thru time is the car-weighted average lane time. "
        "4:00 or under is green, and above 4:00 is red. "
        "A blank cell stays uncolored. "
        "OT hours leave out the GM role. All stores adds the "
        "stores together, then divides.",
        "Annualized net sales is the net sales from these 7 days, "
        "multiplied by 52. Share of week is each store's net sales "
        "divided by the All stores total.",
    ]
    if recorded_days < 7:
        sales_notes.insert(
            0,
            f"Sales are recorded on {recorded_days} of 7 days in "
            "this window. Annualized net sales still multiplies "
            "those 7 days by 52.",
        )
    week_lanes = (
        comparison["week_lane"].dropna().astype(str).str.strip()
    )
    if week_lanes.empty or week_lanes.isin({"", "—"}).all():
        sales_notes.append(
            "No drive-thru departures were recorded in this window."
        )

    review_notes = [
        "Stores follow opening order: Marble Falls, Lampasas, "
        "Leander, then Temple. All stores pools every review "
        "across stores. Week change columns are this week "
        "minus the previous 7 days. Day columns use the "
        "business date in the title. Current is the average "
        "of every Google review stored for the store.",
    ]

    return Report(
        business_date=day,
        week_start=week_start,
        week_end=week_end,
        week_label=api["format_week_range"](week_start, week_end),
        recorded_days=recorded_days,
        generated_at=datetime.now(STORE_TIMEZONE),
        warnings=warnings,
        sales=comparison,
        reviews=api["review_leaderboard"](comparison),
        sales_notes=sales_notes,
        review_notes=review_notes,
    )


def is_blank(value) -> bool:
    """True when a cell should be left empty."""

    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except TypeError:
        return False


def excel_value(kind: str, value):
    """A number Excel can format, or text. Blanks stay blank."""

    if kind == "text":
        if is_blank(value):
            return None
        text = str(value).strip()
        return text or None
    if is_blank(value):
        return None
    number = float(value)
    if kind == "int":
        return int(round(number))
    return number


def display_value(kind: str, value) -> str:
    """The same figure, formatted for the PDF."""

    if is_blank(value):
        return ""
    if kind == "text":
        return str(value)
    number = float(value)
    if kind == "money0":
        return f"${number:,.0f}"
    if kind == "money2":
        return f"${number:,.2f}"
    if kind == "points":
        return f"{number:.1f}%"
    if kind == "share":
        return f"{number * 100:.1f}%"
    if kind == "num1":
        return f"{number:.1f}"
    if kind == "score":
        return f"{number:.2f}"
    if kind == "delta2":
        return f"{number:+.2f}"
    if kind == "delta0":
        return f"{number:+.0f}"
    if kind == "int":
        return f"{number:.0f}"
    return str(value)


def _fills():
    return {
        "navy": PatternFill("solid", fgColor=NAVY),
        "zebra": PatternFill("solid", fgColor=ZEBRA),
        "total": PatternFill("solid", fgColor=TOTAL),
        "amber": PatternFill("solid", fgColor=AMBER),
        "white": PatternFill("solid", fgColor=WHITE),
        "good": PatternFill("solid", fgColor=GOOD_FILL),
        "bad": PatternFill("solid", fgColor=BAD_FILL),
    }


def performance_tone(key: str, row) -> str | None:
    """Good or bad, using the same rules as the Summary page.

    Hourly labor is green at 18% or under. Weekly growth is
    green at zero or above. Drive-thru time is green at 4:00
    or under, judged on the rounded clock the table shows.
    A blank cell stays uncolored.
    """

    if key in {"day_labor_percent", "week_labor_percent"}:
        value = row[key]
        if is_blank(value):
            return None
        return "good" if float(value) <= HOURLY_LABOR_CUTOFF else "bad"
    if key == "week_net_sales_growth":
        value = row[key]
        if is_blank(value):
            return None
        return "good" if float(value) >= 0 else "bad"

    seconds_key = {
        "day_lane": "day_lane_seconds",
        "week_lane": "week_lane_seconds",
    }.get(key)
    if seconds_key is None or seconds_key not in row.index:
        return None
    raw = row[seconds_key]
    if is_blank(raw):
        return None
    shown = int(round(float(raw)))
    return "good" if shown <= LANE_GOAL_SECONDS else "bad"


def _border(color: str = LINE, style: str = "thin") -> Border:
    side = Side(style=style, color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def configure_print(ws, fit_height: int) -> None:
    """Landscape Letter, scaled to the page width, ready to save as PDF."""

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(
        left=0.4,
        right=0.4,
        top=0.55,
        bottom=0.45,
        header=0.2,
        footer=0.22,
    )
    ws.sheet_view.showGridLines = False
    ws.sheet_view.view = "pageLayout"
    ws.sheet_view.zoom = 100
    ws.print_options.horizontalCentered = True
    ws.oddFooter.left.text = "&8&K667085HTH Solutions  ·  Summary (Weekly)"
    ws.oddFooter.right.text = "&8&K667085Page &P of &N"
    ws.sheet_properties.tabColor = NAVY


def write_block(
    ws,
    start_row: int,
    columns: list[Column],
    frame: pd.DataFrame,
    emphasize_last: bool,
) -> int:
    """Write a header and body. Returns the row after the table."""

    fills = _fills()
    header_font = Font(name="Calibri", bold=True, color=WHITE, size=9)
    header_alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )
    grid = _border()
    last_index = len(frame) - 1

    ws.row_dimensions[start_row].height = 32
    for offset, column in enumerate(columns, start=1):
        cell = ws.cell(start_row, offset, column.header)
        cell.font = header_font
        cell.fill = fills["navy"]
        cell.alignment = header_alignment
        cell.border = grid
        if column.comment:
            cell.comment = Comment(column.comment, "Summary")
        ws.column_dimensions[get_column_letter(offset)].width = column.width

    for row_offset in range(len(frame)):
        row_number = start_row + 1 + row_offset
        is_total = emphasize_last and row_offset == last_index
        record = frame.iloc[row_offset]
        ws.row_dimensions[row_number].height = 18
        for offset, column in enumerate(columns, start=1):
            raw = record[column.key]
            cell = ws.cell(
                row_number,
                offset,
                excel_value(column.kind, raw),
            )
            tone = performance_tone(column.key, record)
            if tone == "good":
                ink = GOOD_INK
                cell.fill = fills["good"]
            elif tone == "bad":
                ink = BAD_INK
                cell.fill = fills["bad"]
            else:
                ink = INK
                if is_total:
                    cell.fill = fills["total"]
                elif row_offset % 2 == 1:
                    cell.fill = fills["zebra"]
            cell.font = Font(
                name="Calibri",
                bold=is_total,
                size=9,
                color=ink,
            )
            cell.border = grid
            if column.kind == "text":
                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                )
            else:
                cell.alignment = Alignment(
                    horizontal="right",
                    vertical="center",
                )
                if column.kind in EXCEL_FORMATS:
                    cell.number_format = EXCEL_FORMATS[column.kind]

    return start_row + len(frame) + 1


def write_lines(
    ws,
    row: int,
    lines: list[str],
    width: int,
    fill_key: str | None,
    ink: str,
) -> int:
    """Write wrapped note lines under a table."""

    fills = _fills()
    for line in lines:
        ws.merge_cells(
            start_row=row,
            start_column=1,
            end_row=row,
            end_column=width,
        )
        cell = ws.cell(row, 1, line)
        cell.font = Font(name="Calibri", size=8, italic=True, color=ink)
        cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True,
        )
        if fill_key:
            cell.fill = fills[fill_key]
        ws.row_dimensions[row].height = 16 + 12 * max(
            0,
            len(line) // 140,
        )
        row += 1
    return row


def slice_color(index: int) -> str:
    """Category color for one slice, repeating after twenty stores."""

    return CATEGORY_20[index % len(CATEGORY_20)]


def week_sales_slices(sales: pd.DataFrame) -> pd.DataFrame:
    """One row per store with week net sales, largest slice first."""

    columns = [
        "location",
        "net_sales",
        "share",
        "percent_label",
        "legend_label",
        "slice_label",
    ]
    empty = pd.DataFrame(columns=columns)
    if sales.empty or "week_net_sales" not in sales.columns:
        return empty

    plotted = sales
    if "row_key" in plotted.columns:
        plotted = plotted.loc[plotted["row_key"] != "ALL"]
    plotted = plotted.loc[
        pd.to_numeric(plotted["week_net_sales"], errors="coerce")
        .fillna(0)
        .gt(0)
    ].copy()
    if plotted.empty:
        return empty

    plotted["net_sales"] = pd.to_numeric(
        plotted["week_net_sales"],
        errors="coerce",
    )
    total = float(plotted["net_sales"].sum())
    if total <= 0:
        return empty

    plotted = plotted.sort_values("net_sales", ascending=False)
    plotted["share"] = plotted["net_sales"] / total
    plotted["percent_label"] = plotted["share"].map(
        lambda share: f"{share * 100:.1f}%"
    )
    plotted["legend_label"] = (
        plotted["location"].astype(str) + " · " + plotted["percent_label"]
    )
    plotted["slice_label"] = plotted["percent_label"].where(
        plotted["share"] >= SLICE_LABEL_MIN_SHARE,
        "",
    )
    return plotted[columns].reset_index(drop=True)


def _pie_data_sheet(workbook, slices: pd.DataFrame):
    """Hidden rows the Excel doughnut chart reads."""

    name = "Pie data"
    if name in workbook.sheetnames:
        data_ws = workbook[name]
        for cells in data_ws.iter_rows():
            for cell in cells:
                cell.value = None
    else:
        data_ws = workbook.create_sheet(name)
    data_ws.sheet_state = "hidden"
    data_ws["A1"] = "Store"
    data_ws["B1"] = "Net sales"
    for index, record in slices.iterrows():
        excel_row = int(index) + 2
        data_ws.cell(excel_row, 1, record["legend_label"])
        value = data_ws.cell(excel_row, 2, float(record["net_sales"]))
        value.number_format = '"$"#,##0'
    return data_ws


def write_pie_block(ws, report: Report, start_row: int, width: int) -> int:
    """Donut of week net sales, then the slice caption. Returns the next row."""

    ws.merge_cells(
        start_row=start_row,
        start_column=1,
        end_row=start_row,
        end_column=width,
    )
    heading = ws.cell(start_row, 1, "Share of week net sales")
    heading.font = Font(name="Calibri", bold=True, size=13, color=NAVY)
    heading.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[start_row].height = 20
    row = start_row + 1

    slices = week_sales_slices(report.sales)
    if slices.empty:
        return write_lines(
            ws,
            row,
            ["No store recorded net sales in this week."],
            width,
            None,
            MUTED,
        )

    data_ws = _pie_data_sheet(ws.parent, slices)
    chart = DoughnutChart()
    chart.holeSize = 46
    chart.title = None
    last_row = 1 + len(slices)
    chart.add_data(
        Reference(data_ws, min_col=2, min_row=1, max_row=last_row),
        titles_from_data=True,
    )
    chart.set_categories(
        Reference(data_ws, min_col=1, min_row=2, max_row=last_row)
    )
    chart.legend.position = "r"
    chart.width = 22
    chart.height = 12

    points = []
    point_labels = []
    for index, share in enumerate(slices["share"]):
        point = DataPoint(idx=index)
        point.graphicalProperties.solidFill = slice_color(index)
        points.append(point)
        label = DataLabel(idx=index)
        label.showVal = False
        label.showCatName = False
        label.showSerName = False
        label.showPercent = bool(share >= SLICE_LABEL_MIN_SHARE)
        point_labels.append(label)
    series = chart.series[0]
    series.data_points = points
    label_list = DataLabelList()
    label_list.showPercent = False
    label_list.showVal = False
    label_list.showCatName = False
    label_list.showSerName = False
    label_list.dLbl = point_labels
    series.dLbls = label_list

    ws.add_chart(chart, f"A{row}")
    reserved = 26
    for offset in range(reserved):
        ws.row_dimensions[row + offset].height = 15
    row += reserved
    return write_lines(ws, row, [PIE_CAPTION], width, None, MUTED)


def write_workbook(report: Report, path: Path) -> None:
    """Build the print-ready workbook."""

    wb = Workbook()
    pie_ws = wb.active
    pie_ws.title = "Share of week"
    sales_ws = wb.create_sheet("Sales and Labor")
    reviews_ws = wb.create_sheet("Review leaderboard")

    write_pie_sheet(pie_ws, report)
    write_sheet(
        sales_ws,
        "Sales and Labor Performance",
        report,
        SALES_COLUMNS,
        report.sales,
        report.sales_notes,
        emphasize_last=True,
    )
    write_sheet(
        reviews_ws,
        "Review leaderboard",
        report,
        REVIEW_COLUMNS,
        report.reviews,
        report.review_notes,
        emphasize_last=False,
    )

    wb.properties.title = f"Summary (Weekly) {report.week_label}"
    wb.properties.subject = report.heading
    wb.properties.creator = "Revel weekly summary export"
    wb.save(path)


def write_header(ws, report: Report, width: int) -> None:
    """Title block shared by the chart page and each table page."""

    ws.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=width,
    )
    title = ws.cell(1, 1, "Summary (Weekly)")
    title.font = Font(name="Calibri", bold=True, size=18, color=NAVY)
    title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells(
        start_row=2,
        start_column=1,
        end_row=2,
        end_column=width,
    )
    subtitle = ws.cell(2, 1, report.heading)
    subtitle.font = Font(name="Calibri", size=11, color=INK)
    subtitle.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18

    generated = (
        "Generated "
        f"{report.generated_at:%b %d, %Y %I:%M %p} Central"
    )
    ws.merge_cells(
        start_row=3,
        start_column=1,
        end_row=3,
        end_column=width,
    )
    stamp = ws.cell(3, 1, generated)
    stamp.font = Font(name="Calibri", size=8, color=MUTED)
    ws.row_dimensions[3].height = 16


def write_pie_sheet(ws, report: Report) -> None:
    """Page 1 when the workbook is saved as a PDF: the week-sales pie."""

    width = 12
    for column in range(1, width + 1):
        ws.column_dimensions[get_column_letter(column)].width = 14
    ws.sheet_view.showGridLines = False
    write_header(ws, report, width)
    row = 5
    if report.warnings:
        row = write_lines(
            ws,
            row,
            report.warnings,
            width,
            "amber",
            AMBER_INK,
        )
        row += 1
    row = write_pie_block(ws, report, row, width)
    configure_print(ws, 1)
    ws.print_area = f"A1:{get_column_letter(width)}{max(row - 1, 1)}"
    ws.oddHeader.left.text = "&B&8&K1F2933Summary (Weekly)"
    ws.oddHeader.right.text = f"&8&K1F2933{report.week_label}"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def write_sheet(
    ws,
    section: str,
    report: Report,
    columns: list[Column],
    frame: pd.DataFrame,
    notes: list[str],
    emphasize_last: bool,
) -> None:
    """One worksheet: title, one table, and the definitions under it."""

    width = len(columns)
    ws.sheet_view.showGridLines = False
    write_header(ws, report, width)

    row = 5
    if report.warnings:
        row = write_lines(
            ws,
            row,
            report.warnings,
            width,
            "amber",
            AMBER_INK,
        )
        row += 1

    ws.merge_cells(
        start_row=row,
        start_column=1,
        end_row=row,
        end_column=width,
    )
    heading = ws.cell(row, 1, section)
    heading.font = Font(name="Calibri", bold=True, size=13, color=NAVY)
    heading.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20
    row += 1

    header_row = row
    table_end = write_block(
        ws,
        header_row,
        columns,
        frame,
        emphasize_last,
    )
    last_table_row = max(header_row, table_end - 1)
    ws.auto_filter.ref = (
        f"A{header_row}:{get_column_letter(width)}{last_table_row}"
    )
    ws.print_title_rows = f"{header_row}:{header_row}"
    ws.freeze_panes = f"A{header_row + 1}"

    row = table_end + 1
    write_lines(ws, row, notes, width, None, MUTED)

    configure_print(ws, 1)
    note_end = row + max(len(notes), 1) - 1
    ws.print_area = f"A1:{get_column_letter(width)}{note_end}"
    ws.oddHeader.left.text = "&B&8&K1F2933Summary (Weekly)"
    ws.oddHeader.right.text = f"&8&K1F2933{report.week_label}"


def register_fonts() -> tuple[str, str]:
    """Use Calibri when Windows has it, so the PDF matches Excel."""

    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    regular = Path(r"C:\Windows\Fonts\calibri.ttf")
    bold = Path(r"C:\Windows\Fonts\calibrib.ttf")
    if regular.is_file() and bold.is_file():
        pdfmetrics.registerFont(TTFont("Calibri", str(regular)))
        pdfmetrics.registerFont(TTFont("Calibri-Bold", str(bold)))
        return "Calibri", "Calibri-Bold"
    return "Helvetica", "Helvetica-Bold"


def pie_drawing(slices: pd.DataFrame, regular: str, bold: str):
    """Donut chart matching the Summary page, legend on the right."""

    from reportlab.graphics.charts.legends import Legend
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.shapes import Drawing
    from reportlab.lib import colors

    count = len(slices)
    height = max(340, 48 + count * 14)
    drawing = Drawing(700, height)
    pie = Pie()
    pie.x = 40
    pie.y = max(24, (height - 280) / 2)
    pie.width = 260
    pie.height = 260
    pie.innerRadiusFraction = 68 / 148
    pie.startAngle = 90
    pie.direction = "clockwise"
    pie.data = [float(value) for value in slices["net_sales"]]
    pie.labels = [
        "" if label is None else str(label)
        for label in slices["slice_label"]
    ]
    pie.simpleLabels = 1
    pie.checkLabelOverlap = 1
    pie.slices.strokeColor = colors.white
    pie.slices.strokeWidth = 1.25
    pie.slices.fontName = bold
    pie.slices.fontSize = 8
    pie.slices.fontColor = colors.HexColor("#1F2933")
    pie.slices.labelRadius = 1.2
    swatches = []
    for index in range(count):
        color = colors.HexColor(f"#{slice_color(index)}")
        pie.slices[index].fillColor = color
        swatches.append((color, str(slices.iloc[index]["legend_label"])))

    legend = Legend()
    legend.x = 360
    legend.y = height - 16
    legend.boxAnchor = "nw"
    legend.fontName = regular
    legend.fontSize = 8
    legend.dx = 8
    legend.dy = 8
    legend.dxTextSpace = 6
    legend.deltay = 14
    legend.alignment = "right"
    legend.columnMaximum = count
    legend.colorNamePairs = swatches
    legend.strokeWidth = 0
    drawing.add(pie)
    drawing.add(legend)
    return drawing


def write_pdf(report: Report, path: Path) -> None:
    """Pie on page 1, then one page for each table."""

    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    regular, bold = register_fonts()
    page_size = landscape(letter)
    doc = SimpleDocTemplate(
        str(path),
        pagesize=page_size,
        leftMargin=0.4 * inch,
        rightMargin=0.4 * inch,
        topMargin=0.48 * inch,
        bottomMargin=0.42 * inch,
        title=f"Summary (Weekly) {report.week_label}",
        author="Revel weekly summary export",
    )
    ink = colors.HexColor("#1F2933")
    muted = colors.HexColor("#667085")
    navy = colors.HexColor("#1F2933")
    amber_ink = colors.HexColor("#7A4E00")

    title_style = ParagraphStyle(
        "title",
        fontName=bold,
        fontSize=16,
        leading=19,
        textColor=navy,
        alignment=TA_LEFT,
    )
    subtitle_style = ParagraphStyle(
        "subtitle",
        fontName=regular,
        fontSize=9,
        leading=12,
        textColor=ink,
    )
    meta_style = ParagraphStyle(
        "meta",
        fontName=regular,
        fontSize=8,
        leading=10,
        textColor=muted,
    )
    section_style = ParagraphStyle(
        "section",
        fontName=bold,
        fontSize=11,
        leading=14,
        textColor=navy,
        spaceBefore=2,
    )
    note_style = ParagraphStyle(
        "note",
        fontName=regular,
        fontSize=7.5,
        leading=9.5,
        textColor=muted,
    )
    warning_style = ParagraphStyle(
        "warning",
        fontName=regular,
        fontSize=8,
        leading=10,
        textColor=amber_ink,
    )
    header_style = ParagraphStyle(
        "th",
        fontName=bold,
        fontSize=7,
        leading=8.5,
        textColor=colors.white,
        alignment=TA_CENTER,
    )
    text_style = ParagraphStyle(
        "td",
        fontName=regular,
        fontSize=7.5,
        leading=9,
        textColor=ink,
        alignment=TA_LEFT,
    )
    text_bold = ParagraphStyle(
        "td-bold",
        parent=text_style,
        fontName=bold,
    )
    number_style = ParagraphStyle(
        "num",
        fontName=regular,
        fontSize=7.5,
        leading=9,
        textColor=ink,
        alignment=TA_RIGHT,
    )
    number_bold = ParagraphStyle(
        "num-bold",
        parent=number_style,
        fontName=bold,
    )

    usable = page_size[0] - doc.leftMargin - doc.rightMargin

    def section_table(columns: list[Column], frame: pd.DataFrame, emphasize_last: bool):
        store_width = 1.35 * inch
        other = (usable - store_width) / max(len(columns) - 1, 1)
        widths = []
        for column in columns:
            if column.key == "location":
                widths.append(store_width)
            else:
                widths.append(other)

        header = [
            Paragraph(escape(column.header), header_style)
            for column in columns
        ]
        rows = [header]
        last_index = len(frame) - 1
        tone_cells = []
        for row_offset in range(len(frame)):
            emphasize = emphasize_last and row_offset == last_index
            label_style = text_bold if emphasize else text_style
            value_style = number_bold if emphasize else number_style
            record = frame.iloc[row_offset]
            row = []
            for col_offset, column in enumerate(columns):
                shown = display_value(column.kind, record[column.key])
                tone = performance_tone(column.key, record)
                if tone == "good":
                    shown = f'<font color="#{GOOD_INK}">{escape(shown)}</font>'
                    tone_cells.append((row_offset + 1, col_offset, "good"))
                elif tone == "bad":
                    shown = f'<font color="#{BAD_INK}">{escape(shown)}</font>'
                    tone_cells.append((row_offset + 1, col_offset, "bad"))
                else:
                    shown = escape(shown)
                if column.kind == "text":
                    style = label_style
                else:
                    style = value_style
                row.append(Paragraph(shown, style))
            rows.append(row)

        table = Table(rows, colWidths=widths, repeatRows=1)
        style_commands = [
            ("BACKGROUND", (0, 0), (-1, 0), navy),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), bold),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D0D5DD")),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ]
        for row_offset in range(len(frame)):
            excel_row = row_offset + 1
            if emphasize_last and row_offset == last_index:
                style_commands.append(
                    ("BACKGROUND", (0, excel_row), (-1, excel_row), colors.HexColor("#E8EEF4"))
                )
            elif row_offset % 2 == 1:
                style_commands.append(
                    ("BACKGROUND", (0, excel_row), (-1, excel_row), colors.HexColor("#F7F8FA"))
                )
        for excel_row, col_offset, tone in tone_cells:
            fill = GOOD_FILL if tone == "good" else BAD_FILL
            style_commands.append(
                (
                    "BACKGROUND",
                    (col_offset, excel_row),
                    (col_offset, excel_row),
                    colors.HexColor(f"#{fill}"),
                )
            )
        table.setStyle(TableStyle(style_commands))
        return table

    def add_section(
        story,
        section,
        columns,
        frame,
        notes,
        emphasize_last,
    ):
        add_banner(story)
        story.append(Paragraph(escape(section), section_style))
        story.append(Spacer(1, 4))
        if frame.empty:
            story.append(Paragraph("No rows for this week.", note_style))
        else:
            story.append(
                section_table(columns, frame, emphasize_last)
            )
        story.append(Spacer(1, 8))
        for note in notes:
            story.append(Paragraph(escape(note), note_style))
            story.append(Spacer(1, 3))

    def add_banner(story):
        story.append(Paragraph("Summary (Weekly)", title_style))
        story.append(Spacer(1, 2))
        story.append(Paragraph(escape(report.heading), subtitle_style))
        story.append(
            Paragraph(
                escape(
                    "Generated "
                    f"{report.generated_at:%b %d, %Y %I:%M %p} Central"
                ),
                meta_style,
            )
        )
        story.append(Spacer(1, 8))
        for message in report.warnings:
            story.append(Paragraph(escape(message), warning_style))
            story.append(Spacer(1, 3))

    slices = week_sales_slices(report.sales)
    story = []
    add_banner(story)
    story.append(Paragraph("Share of week net sales", section_style))
    story.append(Spacer(1, 6))
    if slices.empty:
        story.append(
            Paragraph("No store recorded net sales in this week.", note_style)
        )
    else:
        story.append(pie_drawing(slices, regular, bold))
        story.append(Spacer(1, 8))
        story.append(Paragraph(escape(PIE_CAPTION), note_style))
    story.append(PageBreak())
    add_section(
        story,
        "Sales and Labor Performance",
        SALES_COLUMNS,
        report.sales,
        report.sales_notes,
        True,
    )
    story.append(PageBreak())
    add_section(
        story,
        "Review leaderboard",
        REVIEW_COLUMNS,
        report.reviews,
        report.review_notes,
        False,
    )

    def draw_footer(canvas, document):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#1F2933"))
        canvas.rect(
            0,
            page_size[1] - 8,
            page_size[0],
            8,
            stroke=0,
            fill=1,
        )
        canvas.setFillColor(muted)
        canvas.setFont(regular, 8)
        canvas.drawString(
            document.leftMargin,
            0.22 * inch,
            "HTH Solutions  ·  Summary (Weekly)",
        )
        canvas.drawRightString(
            page_size[0] - document.rightMargin,
            0.22 * inch,
            f"Page {canvas.getPageNumber()}",
        )
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)


def parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Export the Revel Summary page for the current week "
            "to an Excel file and a PDF."
        )
    )
    parser.add_argument(
        "--date",
        help=(
            "Business date (YYYY-MM-DD). The week is the 7 days "
            "ending on this date. Default: yesterday in Central "
            "Time, same as the Summary page."
        ),
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=REPO_ROOT,
        help="Folder for the Excel and PDF files. Defaults to this repo.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    if not REVEL_DEMO.is_dir():
        raise SystemExit(
            f"Could not find the revel-demo repo at {REVEL_DEMO}."
        )

    logging.disable(logging.WARNING)
    previous = Path.cwd()
    os.chdir(REVEL_DEMO)
    demo = str(REVEL_DEMO)
    if demo not in sys.path:
        sys.path.insert(0, demo)
    try:
        print("Loading Summary calculations...", flush=True)
        api = load_summary_api()
        print("Loading the current week from Supabase...", flush=True)
        report = build_report(api, args.date)
    finally:
        os.chdir(previous)

    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = output_dir / f"{report.file_stem}.xlsx"
    pdf_path = output_dir / f"{report.file_stem}.pdf"

    print(f"Writing {workbook_path.name}...", flush=True)
    write_workbook(report, workbook_path)
    print(f"Writing {pdf_path.name}...", flush=True)
    write_pdf(report, pdf_path)

    print(report.heading)
    print(f"Excel: {workbook_path}")
    print(f"PDF:   {pdf_path}")
    if report.warnings:
        print("Notes:")
        for message in report.warnings:
            print(f"- {message}")


if __name__ == "__main__":
    main()
