#!/usr/bin/env python3
r"""
Create a MotoGP season workbook with probabilistic charts and an editable
championship scenario solver from JSON files produced by motogp_parser_generic.py.

The workbook intentionally separates:
1. Historical/model estimates from existing Grand Prix JSON data.
2. Official championship inputs entered manually, including Sprint-adjusted totals.
3. Deterministic scenario math for selected future GP/Sprint finishes.

Example:
    python motogp_probability_and_scenarios.py "C:\MotoGP\Season_JSON" \
        -o "C:\MotoGP\motogp_probability_scenarios.xlsx"
"""

import argparse
import colorsys
import json
import math
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import xlsxwriter


GP_POINTS = {1:25, 2:20, 3:16, 4:13, 5:11, 6:10, 7:9, 8:8, 9:7,
             10:6, 11:5, 12:4, 13:3, 14:2, 15:1}
SPRINT_POINTS = {1:12, 2:9, 3:7, 4:6, 5:5, 6:4, 7:3, 8:2, 9:1}
DATE_FORMATS = ("%B %d, %Y", "%b %d, %Y", "%Y-%m-%d", "%d %B %Y", "%d %b %Y")


def parse_event_date(value):
    if not value:
        raise ValueError("Missing event_date")
    cleaned = str(value).strip()
    match = re.search(
        r"(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),\s*(.+)$",
        cleaned, re.IGNORECASE,
    )
    if match:
        cleaned = match.group(1).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            pass
    raise ValueError(f"Unsupported event_date format: {value!r}")


def rider_key(rider):
    return f"{str(rider.get('number', '')).strip()}|{' '.join(str(rider.get('name','')).split()).upper()}"


def load_races(folder_path):
    folder = Path(folder_path)
    files = sorted(folder.glob("*.json"))
    if not files:
        raise FileNotFoundError(f"No JSON files found in {folder}")
    races, errors = [], []
    for file in files:
        try:
            data = json.loads(file.read_text(encoding="utf-8"))
            riders = data.get("riders")
            if not riders:
                raise ValueError("Missing riders")
            races.append({
                "file": file.name,
                "date": parse_event_date(data.get("event_date")),
                "race": data.get("race", file.stem),
                "circuit": data.get("circuit", ""),
                "riders": riders,
            })
        except Exception as exc:
            errors.append(f"{file.name}: {exc}")
    if errors:
        raise ValueError("Unable to load JSON files:\n- " + "\n- ".join(errors))
    races.sort(key=lambda race: (race["date"], race["race"], race["file"]))
    for index, race in enumerate(races, 1):
        race["round"] = index
    return races


def palette(count):
    result, hue = [], 0.17
    for index in range(count):
        hue = (hue + 0.618033988749895) % 1.0
        red, green, blue = colorsys.hsv_to_rgb(hue, 0.72 + 0.1 * (index % 2), 0.78 + 0.1 * (index % 3 > 0))
        result.append(f"#{int(red*255):02X}{int(green*255):02X}{int(blue*255):02X}")
    return result


def calculate_statistics(races, recent_window=5):
    entries = defaultdict(list)
    profiles = {}
    for race in races:
        leader_laps = max(int(r.get("total_laps", len(r.get("laps", [])))) for r in race["riders"])
        # Normalize median lap pace against the median of rider medians in the same race.
        medians = []
        per_rider_median = {}
        for rider in race["riders"]:
            valid = [lap["lap_seconds"] for lap in rider.get("laps", [])
                     if not lap.get("cancelled") and not lap.get("partial_timing") and lap.get("lap_seconds") is not None]
            if valid:
                valid = sorted(valid)
                mid = len(valid) // 2
                median = valid[mid] if len(valid) % 2 else (valid[mid-1] + valid[mid]) / 2
                per_rider_median[rider_key(rider)] = median
                medians.append(median)
        field_median = sorted(medians)[len(medians)//2] if medians else None

        for rider in race["riders"]:
            key = rider_key(rider)
            profiles.setdefault(key, {
                "key": key, "name": rider.get("name", "Unknown"), "number": rider.get("number"),
                "nationality": rider.get("nationality", ""), "manufacturer": rider.get("manufacturer", ""),
            })
            position = int(rider["position"])
            completed = int(rider.get("total_laps", len(rider.get("laps", []))))
            shortened = completed < leader_laps
            median = per_rider_median.get(key)
            pace_delta = ((median / field_median) - 1.0) if median and field_median else None
            entries[key].append({
                "round": race["round"], "position": position, "shortened": shortened,
                "gp_points": GP_POINTS.get(position, 0), "pace_delta": pace_delta,
            })

    statistics = []
    for key, profile in profiles.items():
        rows = entries[key]
        positions = [row["position"] for row in rows]
        recent = rows[-recent_window:]
        pace_values = [row["pace_delta"] for row in rows if row["pace_delta"] is not None]
        n = len(rows)
        wins = sum(row["position"] == 1 for row in rows)
        podiums = sum(row["position"] <= 3 for row in rows)
        top5 = sum(row["position"] <= 5 for row in rows)
        shortened = sum(row["shortened"] for row in rows)
        gp_points = sum(row["gp_points"] for row in rows)
        average_finish = sum(positions) / n
        recent_average = sum(row["position"] for row in recent) / len(recent)
        variance = sum((position-average_finish)**2 for position in positions) / n
        # Beta(1,1) smoothed empirical probabilities prevent 0% and 100% with small samples.
        win_probability = (wins + 1) / (n + 2)
        podium_probability = (podiums + 1) / (n + 2)
        top5_probability = (top5 + 1) / (n + 2)
        shortened_probability = (shortened + 1) / (n + 2)
        statistics.append({
            **profile, "races": n, "wins": wins, "podiums": podiums, "top5": top5,
            "gp_points": gp_points, "average_finish": average_finish,
            "recent_average": recent_average, "finish_stddev": math.sqrt(variance),
            "win_probability": win_probability, "podium_probability": podium_probability,
            "top5_probability": top5_probability,
            "shortened_probability": shortened_probability,
            "pace_delta": sum(pace_values)/len(pace_values) if pace_values else None,
        })
    statistics.sort(key=lambda row: (-row["gp_points"], row["average_finish"], row["name"]))
    return statistics


def build(races, output_path, recent_window=5):
    stats = calculate_statistics(races, recent_window)
    colors = palette(len(stats))
    wb = xlsxwriter.Workbook(output_path)
    wb.set_properties({"title":"MotoGP Probability and Championship Scenarios"})

    title = wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF","bg_color":"#1F4E78","align":"center"})
    section = wb.add_format({"bold":True,"font_size":12,"font_color":"#FFFFFF","bg_color":"#5B9BD5"})
    header = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#4472C4","border":1,"align":"center"})
    body = wb.add_format({"border":1})
    integer = wb.add_format({"border":1,"num_format":"0","align":"center"})
    decimal = wb.add_format({"border":1,"num_format":"0.00"})
    percent = wb.add_format({"border":1,"num_format":"0.0%"})
    input_fmt = wb.add_format({"border":1,"bg_color":"#FFF2CC","num_format":"0.00"})
    input_int = wb.add_format({"border":1,"bg_color":"#FFF2CC","num_format":"0","align":"center"})
    output_fmt = wb.add_format({"border":1,"bg_color":"#E2F0D9","num_format":"0.00"})
    output_int = wb.add_format({"border":1,"bg_color":"#E2F0D9","num_format":"0"})
    note = wb.add_format({"italic":True,"font_color":"#666666","text_wrap":True})

    # ------------------------- Model Data -------------------------
    model = wb.add_worksheet("Model Data")
    model_headers = ["Rider","Number","Nationality","Manufacturer","Races","Wins","Podiums","Top 5",
                     "GP Points from JSON","Avg Finish","Recent Avg Finish","Finish StdDev",
                     "Smoothed Win %","Smoothed Podium %","Smoothed Top 5 %","Shortened Race %","Normalized Pace Delta"]
    model.write_row(0,0,model_headers,header)
    for row_index, stat in enumerate(stats,1):
        values = [stat["name"],stat["number"],stat["nationality"],stat["manufacturer"],stat["races"],stat["wins"],
                  stat["podiums"],stat["top5"],stat["gp_points"],stat["average_finish"],stat["recent_average"],
                  stat["finish_stddev"],stat["win_probability"],stat["podium_probability"],stat["top5_probability"],
                  stat["shortened_probability"],stat["pace_delta"]]
        model.write_row(row_index,0,values)
    model.add_table(0,0,len(stats),len(model_headers)-1,{"name":"ModelData","columns":[{"header":h} for h in model_headers],"style":"Table Style Medium 2"})
    model.set_column("A:A",26); model.set_column("B:D",14); model.set_column("E:Q",18); model.freeze_panes(1,0)

    # --------------------- Probabilistic Charts -------------------
    prob = wb.add_worksheet("Probabilistic Charts")
    prob.merge_range("A1:N1","MotoGP Model Estimates",title)
    prob.write("A2","Editable model controls",section)
    controls = [
        ("Season-history weight",0.40,"Weight on full-season empirical rates"),
        ("Recent-form weight",0.40,"Weight on last N race finishing form"),
        ("Pace weight",0.20,"Weight on normalized median-lap pace"),
        ("Recent window",recent_window,"Number of latest entries already reflected in Model Data"),
        ("Probability temperature",8.0,"Higher values make title estimates less concentrated"),
        ("Remaining rounds",4,"Used in projected final points"),
        ("Max points per weekend",37,"25 GP plus 12 Sprint"),
    ]
    prob.write_row("A3",["Control","Value","Explanation"],header)
    for index,(label,value,explanation) in enumerate(controls,3):
        prob.write(index,0,label,body)
        prob.write(index,1,value,input_int if isinstance(value,int) else input_fmt)
        prob.write(index,2,explanation,body)
    prob.data_validation("B4:B6",{"validate":"decimal","criteria":"between","minimum":0,"maximum":1})
    prob.write("A11","Rider adjustments and calculated estimates",section)
    output_headers = ["Rider","Official Current Points","Form Adjustment","Participating","Historical Win %",
                      "Recent Form Score","Pace Score","Adjusted Win %","Podium %","Top 5 %","Expected GP Points",
                      "Projected Final Points","Title Probability Estimate"]
    prob.write_row("A12",output_headers,header)
    start_row = 12
    for index,stat in enumerate(stats):
        row = start_row + index
        excel_row = row + 1
        model_row = index + 2
        prob.write(row,0,stat["name"],body)
        prob.write_number(row,1,stat["gp_points"],input_int) # Replace with official GP+Sprint total.
        prob.write_number(row,2,1.0,input_fmt)
        if hasattr(prob,"insert_checkbox"): prob.insert_checkbox(row,3,True)
        else: prob.write_boolean(row,3,True)
        prob.write_formula(row,4,f"='Model Data'!M{model_row}",percent)
        # Higher score for a better (lower) recent average finish.
        prob.write_formula(row,5,f"=MAX(0,1-('Model Data'!K{model_row}-1)/MAX(1,MAX('Model Data'!$K$2:$K${len(stats)+1})-1))",percent)
        # Faster-than-field pace receives >50%, slower receives <50%.
        prob.write_formula(row,6,f'=IFERROR(1/(1+EXP(40*\'Model Data\'!Q{model_row})),0.5)',percent)
        # Weighted, adjusted, and bounded probability estimate.
        prob.write_formula(row,7,f'=IF(D{excel_row},MIN(0.95,MAX(0.001,C{excel_row}*($B$4*E{excel_row}+$B$5*F{excel_row}+$B$6*G{excel_row}))),0)',percent)
        prob.write_formula(row,8,f"='Model Data'!N{model_row}",percent)
        prob.write_formula(row,9,f"='Model Data'!O{model_row}",percent)
        prob.write_formula(row,10,f"=IF(D{excel_row},AVERAGE('Model Data'!I{model_row}/'Model Data'!E{model_row},SUMPRODUCT(E{excel_row},25,H{excel_row},25)),0)",decimal)
        prob.write_formula(row,11,f"=B{excel_row}+$B$8*K{excel_row}",output_fmt)
        # Softmax on projected final points. This is a transparent heuristic, not Monte Carlo.
        end_excel = start_row + len(stats)
        prob.write_formula(row,12,f"=EXP(L{excel_row}/$B$7)/SUMPRODUCT(EXP($L${start_row+1}:$L${end_excel}/$B$7))",percent)
    end_row = start_row + len(stats)-1
    prob.set_column("A:A",26); prob.set_column("B:M",18); prob.set_column("C:C",14)
    prob.freeze_panes(start_row,1)
    prob.write(end_row+3,0,"Methodology note:",section)
    prob.merge_range(end_row+4,0,end_row+6,12,
        "These are model estimates from historical GP classifications, recent finishing form, and normalized lap pace. "
        "Official Current Points is editable because the JSON files do not contain Sprint classifications. "
        "The title probability estimate is a softmax heuristic over projected final points, not betting odds or an official forecast.",note)

    win_chart = wb.add_chart({"type":"bar"})
    win_chart.add_series({"name":"Adjusted Next-Race Win %","categories":["Probabilistic Charts",start_row,0,end_row,0],
                          "values":["Probabilistic Charts",start_row,7,end_row,7],"fill":{"color":"#4472C4"}})
    win_chart.set_title({"name":"Estimated Next-Race Win Probability"}); win_chart.set_x_axis({"name":"Probability","num_format":"0%"})
    win_chart.set_y_axis({"reverse":True}); win_chart.set_legend({"none":True}); win_chart.set_size({"width":740,"height":500})
    prob.insert_chart("O3",win_chart)

    title_chart = wb.add_chart({"type":"bar"})
    title_chart.add_series({"name":"Title Probability Estimate","categories":["Probabilistic Charts",start_row,0,end_row,0],
                            "values":["Probabilistic Charts",start_row,12,end_row,12],"fill":{"color":"#70AD47"}})
    title_chart.set_title({"name":"Projected Championship Probability Estimate"}); title_chart.set_x_axis({"name":"Probability","num_format":"0%"})
    title_chart.set_y_axis({"reverse":True}); title_chart.set_legend({"none":True}); title_chart.set_size({"width":740,"height":500})
    prob.insert_chart("O30",title_chart)

    # ------------------- Championship Scenario Solver -------------------
    solver = wb.add_worksheet("Championship Solver")
    solver.merge_range("A1:N1","Editable Championship Scenario Solver",title)
    solver.write("A2","Points tables",section)
    solver.write_row("A3",["Position","GP Points","Sprint Points"],header)
    for position in range(1,23):
        row = 3 + position
        solver.write_number(row-1,0,position,integer)
        solver.write_number(row-1,1,GP_POINTS.get(position,0),integer)
        solver.write_number(row-1,2,SPRINT_POINTS.get(position,0),integer)
    solver.write("E2","Scenario inputs",section)
    solver.write_row("E3",["Input","Value"],header)
    input_rows = [
        ("Target rider",stats[0]["name"] if stats else ""),
        ("Target current points",stats[0]["gp_points"] if stats else 0),
        ("Primary rival",stats[1]["name"] if len(stats)>1 else ""),
        ("Rival current points",stats[1]["gp_points"] if len(stats)>1 else 0),
        ("Remaining rounds including scenario",4),
        ("Target next GP finish",1),
        ("Target next Sprint finish",1),
        ("Rival next GP finish",2),
        ("Rival next Sprint finish",2),
    ]
    for index,(label,value) in enumerate(input_rows,3):
        solver.write(index,4,label,body)
        solver.write(index,5,value,input_int if isinstance(value,int) else input_fmt)
    rider_list_end = len(stats)+1
    solver.data_validation("F4",{"validate":"list","source":f"='Model Data'!$A$2:$A${rider_list_end}"})
    solver.data_validation("F6",{"validate":"list","source":f"='Model Data'!$A$2:$A${rider_list_end}"})
    solver.data_validation("F9:F12",{"validate":"integer","criteria":"between","minimum":1,"maximum":30})
    solver.data_validation("F8",{"validate":"integer","criteria":"between","minimum":1,"maximum":30})

    solver.write("E14","Scenario outputs",section)
    scenario_outputs = [
        ("Target GP points","=IFERROR(VLOOKUP(F9,$A$4:$C$25,2,FALSE),0)"),
        ("Target Sprint points","=IFERROR(VLOOKUP(F10,$A$4:$C$25,3,FALSE),0)"),
        ("Rival GP points","=IFERROR(VLOOKUP(F11,$A$4:$C$25,2,FALSE),0)"),
        ("Rival Sprint points","=IFERROR(VLOOKUP(F12,$A$4:$C$25,3,FALSE),0)"),
        ("Target points after scenario","=F5+F15+F16"),
        ("Rival points after scenario","=F7+F17+F18"),
        ("Gap after scenario","=F19-F20"),
        ("Rounds remaining afterward","=MAX(0,F8-1)"),
        ("Maximum points still available","=F22*37"),
        ("Target mathematically alive?",'=IF(F19+F23>=F20,"YES","NO")'),
        ("Rival mathematically alive?",'=IF(F20+F23>=F19,"YES","NO")'),
        ("Required net gain per remaining round","=IF(F22=0,0,MAX(0,F20-F19+1)/F22)"),
    ]
    for index,(label,formula) in enumerate(scenario_outputs,14):
        solver.write(index,4,label,body)
        solver.write_formula(index,5,formula,output_fmt)

    # Minimum rival GP finishing place for a target-rider GP win to erase the current gap that round.
    solver.write("H2","Next GP win condition",section)
    solver.write_row("H3",["Rival GP Finish","Rival GP Points","Net GP Gain vs Target Win","Closes Current Gap?"],header)
    for position in range(1,23):
        row = 3 + position
        excel_row = row + 1
        solver.write_number(row,7,position,integer)
        solver.write_formula(row,8,f"=IFERROR(VLOOKUP(H{excel_row},$A$4:$C$25,2,FALSE),0)",integer)
        solver.write_formula(row,9,f"=25-I{excel_row}",integer)
        solver.write_formula(row,10,f'=IF(J{excel_row}>=$F$7-$F$5,"YES","NO")',body)

    solver.set_column("A:C",14); solver.set_column("E:E",34); solver.set_column("F:F",18); solver.set_column("H:K",22)
    solver.freeze_panes(3,0)
    solver.merge_range("E29:K31",
        "Use Official Current Points for championship-accurate scenarios. The historical points derived from the JSON files include Sunday GP classifications only; Sprint results must be included manually in the current-points cells.",note)

    scenario_chart = wb.add_chart({"type":"column"})
    scenario_chart.add_series({"name":"Current Points","categories":"='Championship Solver'!$E$4,$E$6","values":"='Championship Solver'!$F$5,$F$7","fill":{"color":"#A5A5A5"}})
    scenario_chart.add_series({"name":"After Scenario","categories":"='Championship Solver'!$E$4,$E$6","values":"='Championship Solver'!$F$19,$F$20","fill":{"color":"#ED7D31"}})
    scenario_chart.set_title({"name":"Current vs Scenario Championship Points"}); scenario_chart.set_y_axis({"name":"Points"})
    scenario_chart.set_legend({"position":"bottom"}); scenario_chart.set_size({"width":620,"height":360})
    solver.insert_chart("M3",scenario_chart)

    wb.close()


def main():
    parser = argparse.ArgumentParser(description="Build MotoGP probability charts and championship scenario solver")
    parser.add_argument("input_folder",help="Folder containing race JSON files")
    parser.add_argument("-o","--output",default="motogp_probability_scenarios.xlsx")
    parser.add_argument("--recent-window",type=int,default=5)
    args = parser.parse_args()
    races = load_races(args.input_folder)
    build(races,args.output,args.recent_window)
    print(f"Wrote {args.output}: {len(races)} races, {len(calculate_statistics(races,args.recent_window))} riders")


if __name__ == "__main__":
    main()
