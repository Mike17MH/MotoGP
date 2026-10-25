#!/usr/bin/env python3
"""POC parser for the first three riders in MotoGP Analysis.pdf."""
import argparse, json, re
from pathlib import Path
import pdfplumber

LAP_RE = re.compile(r"(?m)^\s*(\d{1,2})\s+(\d+'\d{2}\.\d{3})\s+(\d{2}\.\d{3})\s+(\d{2}\.\d{3})\s+(\d{2}\.\d{3})\s+(\d{2}\.\d{3})\s+(\d{3}\.\d)\s*$")
RIDERS = [
    {"name":"Marco BEZZECCHI","number":72,"position":1,"team":"Aprilia Racing","manufacturer":"APRILIA"},
    {"name":"Raul FERNANDEZ","number":25,"position":2,"team":"SuperFile Trackhouse MotoGP Team","manufacturer":"APRILIA"},
    {"name":"Jorge MARTIN","number":89,"position":3,"team":"Aprilia Racing","manufacturer":"APRILIA"},
]

def seconds(value):
    minutes, rest = value.split("'")
    return round(int(minutes) * 60 + float(rest), 3)

def rows(text):
    found = {}
    for m in LAP_RE.finditer(text or ""):
        lap = int(m.group(1))
        found[lap] = {
            "lap": lap, "lap_time": m.group(2).replace("'", ":"),
            "lap_seconds": seconds(m.group(2)),
            "t1": float(m.group(3)), "t2": float(m.group(4)),
            "t3": float(m.group(5)), "t4": float(m.group(6)),
            "speed_kmh": float(m.group(7)), "cancelled": False,
        }
    return found

def segment(text, start=None, end=None):
    if start:
        text = text.split(start, 1)[1]
    if end and end in text:
        text = text.split(end, 1)[0]
    return text

def parse(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        p1, p2 = pdf.pages[0], pdf.pages[1]
        left1 = p1.crop((30,180,300,730)).extract_text(x_tolerance=2,y_tolerance=3) or ""
        right1 = p1.crop((300,180,570,730)).extract_text(x_tolerance=2,y_tolerance=3) or ""
        left2 = p2.crop((30,55,300,730)).extract_text(x_tolerance=2,y_tolerance=3) or ""

    blocks = {
        "Marco BEZZECCHI": segment(left1, "Marco BEZZECCHI", "Raul FERNANDEZ"),
        "Raul FERNANDEZ": segment(left1, "Raul FERNANDEZ") + "\n" + segment(right1, None, "Jorge MARTIN"),
        "Jorge MARTIN": segment(right1, "Jorge MARTIN") + "\n" + segment(left2, None, "Ai OGURA"),
    }
    riders = []
    for meta in RIDERS:
        lap_map = rows(blocks[meta["name"]])
        laps = [lap_map[i] for i in sorted(lap_map)]
        if len(laps) != 27:
            raise ValueError(f'{meta["name"]}: expected 27 laps, extracted {len(laps)} ({sorted(lap_map)})')
        cumulative = 0.0
        for lap in laps:
            cumulative = round(cumulative + lap["lap_seconds"], 3)
            lap["cumulative_seconds"] = cumulative
        riders.append({**meta, "total_laps":27, "valid_laps":26, "laps":laps})
    return {
        "source_pdf": str(Path(pdf_path).name),
        "race": "Pertamina Grand Prix of Indonesia",
        "circuit": "Pertamina Mandalika Circuit",
        "poc_rider_count": len(riders),
        "riders": riders,
    }

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="Path to Analysis.pdf")
    ap.add_argument("-o", "--output", default="motogp_poc.json")
    args = ap.parse_args()
    data = parse(args.pdf)
    Path(args.output).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f'Wrote {args.output}: {len(data["riders"])} riders, {sum(len(r["laps"]) for r in data["riders"])} lap rows')
if __name__ == "__main__": main()
