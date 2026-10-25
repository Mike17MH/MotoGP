#!/usr/bin/env python3
r"""
Build a season-level MotoGP finishing-position bump chart from race JSON files.

Expected input: JSON files produced by motogp_parser_generic.py.
The script scans one folder, sorts races by event_date, and creates:
  - Race Calendar
  - Season Results
  - Position Matrix
  - Season Progress with interactive rider checkboxes and bump chart

  The bump chart uses:

    -Race order on the X-axis
    -Finishing position on the Y-axis
    -Reversed Y-axis so first place appears at the top
    -One colored line per rider
    -Markers for every race result
    -Gaps when a rider did not participate
    -Checkboxes to show or hide each rider

Example:
    python motogp_season_bump_chart.py "C:\MotoGP\JSON" -o "C:\MotoGP\motogp_season.xlsx"
"""

import argparse
import colorsys
import json
import re
from datetime import datetime
from pathlib import Path

import xlsxwriter


DATE_FORMATS = (
    "%B %d, %Y",       # October 4, 2026
    "%b %d, %Y",      # Oct 4, 2026
    "%Y-%m-%d",       # 2026-10-04
    "%d %B %Y",       # 4 October 2026
    "%d %b %Y",       # 4 Oct 2026
)


def parse_event_date(value):
    """Convert the parser's event_date text to a Python datetime."""
    if not value:
        raise ValueError("Missing event_date")

    cleaned = str(value).strip()
    # Supports values such as: "Motegi, Sunday, October 4, 2026".
    weekday_match = re.search(
        r"(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),\s*(.+)$",
        cleaned,
        flags=re.IGNORECASE,
    )
    if weekday_match:
        cleaned = weekday_match.group(1).strip()

    for date_format in DATE_FORMATS:
        try:
            return datetime.strptime(cleaned, date_format)
        except ValueError:
            pass

    raise ValueError(f"Unsupported event_date format: {value!r}")


def load_races(input_folder):
    """Load, validate, and chronologically sort all race JSON files."""
    folder = Path(input_folder)
    if not folder.exists():
        raise FileNotFoundError(f"Input folder does not exist: {folder}")

    json_files = sorted(folder.glob("*.json"))
    if not json_files:
        raise FileNotFoundError(f"No JSON files found in: {folder}")

    races = []
    errors = []

    for json_file in json_files:
        try:
            data = json.loads(json_file.read_text(encoding="utf-8"))
            if not isinstance(data.get("riders"), list) or not data["riders"]:
                raise ValueError("Missing or empty riders list")

            event_date = parse_event_date(data.get("event_date"))
            races.append({
                "file": json_file.name,
                "date": event_date,
                "race": data.get("race") or json_file.stem,
                "circuit": data.get("circuit", ""),
                "riders": data["riders"],
            })
        except Exception as exc:
            errors.append(f"{json_file.name}: {exc}")

    if errors:
        raise ValueError("Unable to load all JSON files:\n- " + "\n- ".join(errors))

    races.sort(key=lambda race: (race["date"], race["race"], race["file"]))

    # Prevent accidentally treating duplicate JSON exports as separate rounds.
    duplicate_dates = {}
    for race in races:
        duplicate_dates.setdefault(race["date"].date(), []).append(race["file"])
    duplicates = {date: files for date, files in duplicate_dates.items() if len(files) > 1}
    if duplicates:
        details = "; ".join(f"{date}: {', '.join(files)}" for date, files in duplicates.items())
        raise ValueError(f"Multiple race JSON files use the same event date: {details}")

    for round_number, race in enumerate(races, 1):
        race["round"] = round_number

    return races


def rider_key(rider):
    """Stable rider identity using number and normalized name."""
    number = str(rider.get("number", "")).strip()
    name = " ".join(str(rider.get("name", "")).split()).upper()
    return f"{number}|{name}"


def season_riders(races):
    """Create a stable season rider directory preserving first appearance."""
    directory = {}
    for race in races:
        for rider in race["riders"]:
            key = rider_key(rider)
            if key not in directory:
                directory[key] = {
                    "key": key,
                    "number": rider.get("number"),
                    "name": rider.get("name", "Unknown Rider"),
                    "nationality": rider.get("nationality", ""),
                }
    return list(directory.values())


def rider_palette(count):
    """Generate a deterministic distinct-color palette."""
    colors = []
    golden_ratio = 0.618033988749895
    hue = 0.11
    for index in range(count):
        hue = (hue + golden_ratio) % 1.0
        saturation = 0.70 if index % 2 == 0 else 0.86
        value = 0.72 if index % 3 == 0 else 0.90
        red, green, blue = colorsys.hsv_to_rgb(hue, saturation, value)
        colors.append(f"#{int(red * 255):02X}{int(green * 255):02X}{int(blue * 255):02X}")
    return colors


def excel_column_name(column_index):
    """Convert a zero-based column index to an Excel column name."""
    result = ""
    value = column_index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def race_short_name(race_name, round_number):
    """Create a readable chart category label."""
    name = re.sub(r"\b(MOTUL|PERTAMINA|GRAND PRIX|OF|THE)\b", "", race_name, flags=re.IGNORECASE)
    name = " ".join(name.split()).strip(" -")
    return f"R{round_number} {name or race_name}"


def build_workbook(races, output_path):
    riders = season_riders(races)
    colors = rider_palette(len(riders))
    max_grid = max(len(race["riders"]) for race in races)

    # Result lookup: rider key -> round -> finishing position.
    results = {rider["key"]: {} for rider in riders}
    detailed_rows = []

    for race in races:
        for rider in race["riders"]:
            key = rider_key(rider)
            position = int(rider["position"])
            results[key][race["round"]] = position
            completed_laps = int(rider.get("total_laps", len(rider.get("laps", []))))
            detailed_rows.append([
                race["round"],
                race["date"],
                race["race"],
                race["circuit"],
                rider.get("name", ""),
                rider.get("number"),
                rider.get("team", ""),
                rider.get("manufacturer", ""),
                rider.get("nationality", ""),
                position,
                completed_laps,
                rider.get("valid_laps"),
                race["file"],
            ])

    workbook = xlsxwriter.Workbook(output_path)
    workbook.set_properties({
        "title": "MotoGP Season Finishing Position Progress",
        "comments": "Generated from MotoGP race JSON files",
    })

    title_format = workbook.add_format({
        "bold": True, "font_size": 16, "font_color": "#FFFFFF",
        "bg_color": "#1F4E78", "align": "center", "valign": "vcenter",
    })
    header_format = workbook.add_format({
        "bold": True, "font_color": "#FFFFFF", "bg_color": "#4472C4",
        "border": 1, "align": "center", "valign": "vcenter",
    })
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd", "border": 1})
    integer_format = workbook.add_format({"num_format": "0", "border": 1, "align": "center"})
    body_format = workbook.add_format({"border": 1})
    center_format = workbook.add_format({"border": 1, "align": "center"})
    missing_format = workbook.add_format({"border": 1, "font_color": "#A6A6A6", "align": "center"})

    # ------------------------------------------------------------------
    # Race Calendar
    # ------------------------------------------------------------------
    calendar = workbook.add_worksheet("Race Calendar")
    calendar.merge_range("A1:G1", "MotoGP Race Calendar", title_format)
    calendar_headers = ["Round", "Date", "Race", "Circuit", "Riders", "Maximum Laps", "Source JSON"]
    calendar.write_row("A3", calendar_headers, header_format)

    for row_index, race in enumerate(races, 3):
        maximum_laps = max(
            int(rider.get("total_laps", len(rider.get("laps", []))))
            for rider in race["riders"]
        )
        calendar.write_number(row_index, 0, race["round"], integer_format)
        calendar.write_datetime(row_index, 1, race["date"], date_format)
        calendar.write(row_index, 2, race["race"], body_format)
        calendar.write(row_index, 3, race["circuit"], body_format)
        calendar.write_number(row_index, 4, len(race["riders"]), integer_format)
        calendar.write_number(row_index, 5, maximum_laps, integer_format)
        calendar.write(row_index, 6, race["file"], body_format)

    calendar.add_table(
        2, 0, 2 + len(races), len(calendar_headers) - 1,
        {
            "name": "RaceCalendar",
            "columns": [{"header": header} for header in calendar_headers],
            "style": "Table Style Medium 2",
        },
    )
    calendar.freeze_panes(3, 0)
    calendar.set_column("A:A", 9)
    calendar.set_column("B:B", 13)
    calendar.set_column("C:D", 34)
    calendar.set_column("E:F", 14)
    calendar.set_column("G:G", 28)

    # ------------------------------------------------------------------
    # Normalized Season Results
    # ------------------------------------------------------------------
    season_data = workbook.add_worksheet("Season Results")
    season_headers = [
        "Round", "Date", "Race", "Circuit", "Rider", "Number", "Team",
        "Manufacturer", "Nationality", "Finish Position", "Completed Laps",
        "Valid Laps", "Source JSON",
    ]
    season_data.write_row(0, 0, season_headers, header_format)

    for row_index, row in enumerate(detailed_rows, 1):
        season_data.write_number(row_index, 0, row[0], integer_format)
        season_data.write_datetime(row_index, 1, row[1], date_format)
        for column_index, value in enumerate(row[2:], 2):
            if column_index in (5, 9, 10, 11) and value is not None:
                season_data.write_number(row_index, column_index, value, integer_format)
            else:
                season_data.write(row_index, column_index, value, body_format)

    season_data.add_table(
        0, 0, len(detailed_rows), len(season_headers) - 1,
        {
            "name": "SeasonResults",
            "columns": [{"header": header} for header in season_headers],
            "style": "Table Style Medium 2",
        },
    )
    season_data.freeze_panes(1, 0)
    season_data.set_column("A:B", 12)
    season_data.set_column("C:D", 32)
    season_data.set_column("E:E", 26)
    season_data.set_column("F:F", 10)
    season_data.set_column("G:G", 34)
    season_data.set_column("H:M", 16)

    # ------------------------------------------------------------------
    # Position Matrix
    # ------------------------------------------------------------------
    matrix = workbook.add_worksheet("Position Matrix")
    matrix.merge_range(0, 0, 0, 2 + len(races), "Season Finishing Position Matrix", title_format)
    matrix_headers = ["Rider", "Number", "Nationality"] + [
        race_short_name(race["race"], race["round"]) for race in races
    ]
    matrix.write_row(2, 0, matrix_headers, header_format)

    for rider_index, rider in enumerate(riders):
        row = 3 + rider_index
        matrix.write(row, 0, rider["name"], body_format)
        matrix.write(row, 1, rider["number"], center_format)
        matrix.write(row, 2, rider["nationality"], center_format)
        for race_index, race in enumerate(races):
            position = results[rider["key"]].get(race["round"])
            column = 3 + race_index
            if position is None:
                matrix.write_blank(row, column, None, missing_format)
            else:
                matrix.write_number(row, column, position, integer_format)

    results_start_row = 3
    results_end_row = 2 + len(riders)
    results_start_col = 3
    results_end_col = 2 + len(races)
    matrix.conditional_format(
        results_start_row, results_start_col, results_end_row, results_end_col,
        {"type": "cell", "criteria": "between", "minimum": 1, "maximum": 3,
         "format": workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})},
    )
    matrix.conditional_format(
        results_start_row, results_start_col, results_end_row, results_end_col,
        {"type": "cell", "criteria": "between", "minimum": 4, "maximum": 10,
         "format": workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500"})},
    )
    matrix.conditional_format(
        results_start_row, results_start_col, results_end_row, results_end_col,
        {"type": "cell", "criteria": ">", "value": 10,
         "format": workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})},
    )
    matrix.freeze_panes(3, 3)
    matrix.set_column("A:A", 26)
    matrix.set_column("B:C", 12)
    matrix.set_column(3, 2 + len(races), 16)

    # ------------------------------------------------------------------
    # Interactive Season Progress Bump Chart
    # ------------------------------------------------------------------
    dashboard = workbook.add_worksheet("Season Progress")
    helper_start_col = 4
    helper_end_col = helper_start_col + len(riders)
    dashboard.merge_range(
        0, 0, 0, max(8, helper_end_col),
        "MotoGP Season Finishing Position Progress",
        title_format,
    )
    dashboard.write("A2", f"Races loaded: {len(races)}")
    dashboard.write("B2", f"Season riders: {len(riders)}")
    dashboard.write_row("A3", ["Show", "Rider", "Number", "Color"], header_format)

    for rider_index, rider in enumerate(riders):
        row = 3 + rider_index
        if hasattr(dashboard, "insert_checkbox"):
            dashboard.insert_checkbox(row, 0, True)
        else:
            dashboard.write_boolean(row, 0, True)
        dashboard.write(row, 1, rider["name"], body_format)
        dashboard.write(row, 2, rider["number"], center_format)
        color_cell = workbook.add_format({
            "bg_color": colors[rider_index],
            "font_color": colors[rider_index],
            "border": 1,
        })
        dashboard.write(row, 3, colors[rider_index], color_cell)

    # Helper table: races in rows, riders in columns.
    dashboard.write_row(
        2,
        helper_start_col,
        ["Race"] + [rider["name"] for rider in riders],
        header_format,
    )

    for race_index, race in enumerate(races):
        helper_row = 3 + race_index
        dashboard.write(helper_row, helper_start_col, race_short_name(race["race"], race["round"]))
        for rider_index, rider in enumerate(riders):
            target_col = helper_start_col + 1 + rider_index
            position = results[rider["key"]].get(race["round"])
            if position is None:
                formula = "=NA()"
            else:
                formula = f"=IF($A${4 + rider_index},{position},NA())"
            dashboard.write_formula(helper_row, target_col, formula, None, "#N/A")

    chart = workbook.add_chart({"type": "line"})
    for rider_index, rider in enumerate(riders):
        value_col = helper_start_col + 1 + rider_index
        chart.add_series({
            "name": ["Season Progress", 2, value_col],
            "categories": [
                "Season Progress", 3, helper_start_col,
                2 + len(races), helper_start_col,
            ],
            "values": [
                "Season Progress", 3, value_col,
                2 + len(races), value_col,
            ],
            "line": {"color": colors[rider_index], "width": 2.0},
            "marker": {
                "type": "circle", "size": 5,
                "border": {"color": colors[rider_index]},
                "fill": {"color": colors[rider_index]},
            },
            "gap": 100,
        })

    chart.set_title({"name": "Finishing Position by Race"})
    chart.set_x_axis({
        "name": "Race",
        "text_axis": True,
        "label_position": "low",
        "num_font": {"rotation": -45, "size": 9},
    })
    chart.set_y_axis({
        "name": "Finish Position",
        "reverse": True,
        "min": 1,
        "max": max_grid,
        "major_unit": 1,
        "major_gridlines": {"visible": True, "line": {"color": "#D9E1F2"}},
    })
    chart.set_legend({"position": "bottom"})
    chart.set_size({"width": 1250, "height": 720})

    chart_start_row = max(32, 7 + len(riders))
    dashboard.insert_chart(chart_start_row, 0, chart)
    dashboard.freeze_panes(3, 0)
    dashboard.set_column("A:A", 10)
    dashboard.set_column("B:B", 27)
    dashboard.set_column("C:D", 12)
    dashboard.set_column(helper_start_col, helper_end_col, 16)

    workbook.close()


def main():
    parser = argparse.ArgumentParser(
        description="Build a MotoGP season finishing-position bump chart from race JSON files"
    )
    parser.add_argument("input_folder", help="Folder containing race JSON files")
    parser.add_argument("-o", "--output", default="motogp_season_bump_chart.xlsx")
    args = parser.parse_args()

    races = load_races(args.input_folder)
    build_workbook(races, args.output)
    print(
        f"Wrote {args.output}: {len(races)} races, "
        f"{len(season_riders(races))} season riders"
    )


if __name__ == "__main__":
    main()
