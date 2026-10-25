#!/usr/bin/env python3
"""Build a MotoGP Excel workbook from motogp_parser_generic.py JSON output."""
import argparse
import colorsys
import json
from pathlib import Path
import xlsxwriter


def safe_sheet(name):
    invalid = "[]:*?/\\"
    cleaned = "".join("_" if char in invalid else char for char in name)
    return cleaned.title()[:31]


def rider_colors(count):
    colors = []
    golden_ratio = 0.618033988749895
    hue = 0.58
    for index in range(count):
        hue = (hue + golden_ratio) % 1.0
        saturation = 0.68 if index % 2 == 0 else 0.82
        value = 0.72 if index % 3 == 0 else 0.88
        red, green, blue = colorsys.hsv_to_rgb(hue, saturation, value)
        colors.append(f"#{int(red * 255):02X}{int(green * 255):02X}{int(blue * 255):02X}")
    return colors


def write_optional_number(worksheet, row, col, value, cell_format):
    if value is None:
        worksheet.write_blank(row, col, None, cell_format)
    else:
        worksheet.write_number(row, col, value, cell_format)


def build(json_path, output_path):
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    riders = data["riders"]
    colors = rider_colors(len(riders))

    workbook = xlsxwriter.Workbook(output_path)
    workbook.set_properties({
        "title": f'{data.get("race", "MotoGP Race")} Lap Analysis',
        "comments": f'Generated from {data.get("source_pdf", "MotoGP Analysis PDF")}',
    })

    title = workbook.add_format({
        "bold": True, "font_size": 16, "font_color": "#FFFFFF",
        "bg_color": "#1F4E78", "align": "center", "valign": "vcenter",
    })
    header = workbook.add_format({
        "bold": True, "font_color": "#FFFFFF", "bg_color": "#4472C4",
        "border": 1, "align": "center", "valign": "vcenter",
    })
    num3 = workbook.add_format({"num_format": "0.000", "border": 1})
    num1 = workbook.add_format({"num_format": "0.0", "border": 1})
    integer = workbook.add_format({"num_format": "0", "border": 1})
    text = workbook.add_format({"border": 1})
    centered = workbook.add_format({"border": 1, "align": "center"})
    meta = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    time_format = workbook.add_format({"num_format": "m:ss.000", "border": 1})
    dashboard_time = workbook.add_format({"num_format": "m:ss.000"})
    partial_format = workbook.add_format({"bg_color": "#FFF2CC", "border": 1})

    # Individual rider worksheets.
    for rider in riders:
        worksheet = workbook.add_worksheet(safe_sheet(rider["name"]))
        worksheet.merge_range(
            "A1:M1",
            f'{rider["position"]}. {rider["name"]} #{rider["number"]}',
            title,
        )
        worksheet.write_row("A3", [
            "Team", rider["team"],
            "Manufacturer", rider["manufacturer"],
            "Nationality", rider.get("nationality", ""),
            "Total Laps", rider["total_laps"],
            "Valid Laps", rider["valid_laps"],
            "Front Tyre", rider.get("front_tyre", ""),
            "Rear Tyre", rider.get("rear_tyre", ""),
        ], meta)

        headers = [
            "Lap", "Lap Time", "Lap Seconds", "T1", "T2", "T3", "T4",
            "Speed km/h", "Cumulative Seconds", "Excel Time",
            "Cancelled", "Pit Lane", "Partial Timing",
        ]
        worksheet.write_row("A5", headers, header)

        for row, lap in enumerate(rider["laps"], 5):
            base_format = partial_format if lap.get("partial_timing") else integer
            worksheet.write_number(row, 0, lap["lap"], base_format)
            worksheet.write(row, 1, lap["lap_time"], text)
            worksheet.write_number(row, 2, lap["lap_seconds"], num3)
            for col, key in enumerate(["t1", "t2", "t3", "t4"], 3):
                write_optional_number(worksheet, row, col, lap.get(key), num3)
            write_optional_number(worksheet, row, 7, lap.get("speed_kmh"), num1)
            worksheet.write_number(row, 8, lap["cumulative_seconds"], num3)
            worksheet.write_number(row, 9, lap["lap_seconds"] / 86400.0, time_format)
            worksheet.write_boolean(row, 10, bool(lap.get("cancelled")), centered)
            worksheet.write_boolean(row, 11, bool(lap.get("pit_lane")), centered)
            worksheet.write_boolean(row, 12, bool(lap.get("partial_timing")), centered)

        worksheet.freeze_panes(5, 0)
        worksheet.autofilter(4, 0, 4 + len(rider["laps"]), len(headers) - 1)
        worksheet.set_column("A:A", 8)
        worksheet.set_column("B:B", 12)
        worksheet.set_column("C:J", 18)
        worksheet.set_column("K:M", 15)

    # Combined normalized race data.
    race_data = workbook.add_worksheet("Race Data")
    combined_headers = [
        "Rider", "Number", "Position", "Team", "Manufacturer", "Nationality",
        "Total Laps", "Valid Laps", "Lap", "Lap Time", "Lap Seconds",
        "T1", "T2", "T3", "T4", "Speed km/h", "Cumulative Seconds",
        "Cancelled", "Pit Lane", "Partial Timing",
    ]
    race_data.write_row(0, 0, combined_headers, header)
    output_row = 1
    for rider in riders:
        for lap in rider["laps"]:
            values = [
                rider["name"], rider["number"], rider["position"], rider["team"],
                rider["manufacturer"], rider.get("nationality", ""),
                rider["total_laps"], rider["valid_laps"], lap["lap"], lap["lap_time"],
                lap["lap_seconds"], lap.get("t1"), lap.get("t2"), lap.get("t3"),
                lap.get("t4"), lap.get("speed_kmh"), lap["cumulative_seconds"],
                bool(lap.get("cancelled")), bool(lap.get("pit_lane")),
                bool(lap.get("partial_timing")),
            ]
            race_data.write_row(output_row, 0, values)
            output_row += 1

    race_data.add_table(
        0, 0, output_row - 1, len(combined_headers) - 1,
        {
            "name": "RaceData",
            "columns": [{"header": item} for item in combined_headers],
            "style": "Table Style Medium 2",
        },
    )
    race_data.freeze_panes(1, 0)
    race_data.set_column("A:A", 25)
    race_data.set_column("B:C", 11)
    race_data.set_column("D:D", 34)
    race_data.set_column("E:T", 16)

    # Interactive dashboard.
    dashboard = workbook.add_worksheet("Race Progress")
    helper_start_col = 4
    chart_last_col = helper_start_col + len(riders)
    dashboard.merge_range(
        0, 0, 0, max(7, chart_last_col),
        f'{data.get("race", "MotoGP Race")} - Rider Lap Time Analysis',
        title,
    )
    dashboard.write("A2", data.get("circuit", ""))
    dashboard.write("B2", data.get("event_date", ""))
    dashboard.write_row("A3", ["Show", "Rider", "Color"], header)

    for index, rider in enumerate(riders):
        row = 3 + index
        if hasattr(dashboard, "insert_checkbox"):
            dashboard.insert_checkbox(row, 0, True)
        else:
            dashboard.write_boolean(row, 0, True)
        dashboard.write(row, 1, rider["name"])
        color_format = workbook.add_format({"bg_color": colors[index], "font_color": colors[index]})
        dashboard.write(row, 2, colors[index], color_format)

    dashboard.write_row(2, helper_start_col, ["Lap"] + [r["name"] for r in riders], header)
    max_laps = max(len(rider["laps"]) for rider in riders)

    for lap_index in range(max_laps):
        row = 3 + lap_index
        dashboard.write_number(row, helper_start_col, lap_index + 1)
        for rider_index, rider in enumerate(riders):
            target_col = helper_start_col + 1 + rider_index
            if lap_index < len(rider["laps"]):
                source = safe_sheet(rider["name"]).replace("'", "''")
                formula = (
                    f'=IF($A${4 + rider_index},'
                    f"'{source}'!J{6 + lap_index},NA())"
                )
            else:
                formula = "=NA()"
            dashboard.write_formula(row, target_col, formula, dashboard_time, "#N/A")

    chart = workbook.add_chart({"type": "line"})
    for rider_index, rider in enumerate(riders):
        value_col = helper_start_col + 1 + rider_index
        chart.add_series({
            "name": ["Race Progress", 2, value_col],
            "categories": ["Race Progress", 3, helper_start_col, 2 + max_laps, helper_start_col],
            "values": ["Race Progress", 3, value_col, 2 + max_laps, value_col],
            "line": {"color": colors[rider_index], "width": 2.0},
        })

    chart.set_title({"name": "Lap Time Progress"})
    chart.set_x_axis({"name": "Lap", "major_unit": 1})
    chart.set_y_axis({
        "name": "Lap Time", "num_format": "m:ss.000",
        "major_gridlines": {"visible": True},
    })
    chart.set_legend({"position": "bottom"})
    chart.set_size({"width": 1150, "height": 650})

    chart_start_row = max(31, 6 + len(riders))
    dashboard.insert_chart(chart_start_row, 0, chart)
    dashboard.freeze_panes(3, 0)
    dashboard.set_column("A:A", 10)
    dashboard.set_column("B:B", 26)
    dashboard.set_column("C:C", 14)
    dashboard.set_column(helper_start_col, chart_last_col, 14, dashboard_time)

    workbook.close()
    print(f"Wrote {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Build a MotoGP Excel workbook")
    parser.add_argument("json", help="JSON produced by motogp_parser_generic.py")
    parser.add_argument("-o", "--output", default="motogp_race.xlsx")
    args = parser.parse_args()
    build(args.json, args.output)


if __name__ == "__main__":
    main()
