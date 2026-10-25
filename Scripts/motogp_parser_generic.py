#!/usr/bin/env python3
"""Generic parser for MotoGP Chronological Analysis of Performances PDFs."""
import argparse
import json
import re
from pathlib import Path
import pdfplumber

MANUFACTURERS = "APRILIA|DUCATI|HONDA|KTM|YAMAHA"
HEADER_RE = re.compile(
    rf"(?m)^(?P<name>[A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ' .-]+?)\s+"
    rf"(?P<manufacturer>{MANUFACTURERS})\s+(?P<nation>[A-Z]{{3}})\s*$"
)
POSITION_RE = re.compile(r"(?m)^(?P<position>\d+)(?:st|nd|rd|th)\s+(?P<number>\d+)\s*$")
COUNTS_RE = re.compile(
    r"Runs=(?P<runs>\d+)\s+Total laps=(?P<total>\d+)\s+"
    r"Full laps=(?P<full>\d+)\s+Valid laps=(?P<valid>\d+)"
)
TYRES_RE = re.compile(
    r"Run\s+#\s+(?P<run>\d+)\s+Front Tyre\s+(?P<front>\S+)\s+Rear Tyre\s+(?P<rear>\S+)"
)
LAP_RE = re.compile(r"^\s*(?P<lap>\d{1,3})\s+(?P<time>\d+'\d{2}\.\d{3})\s*(?P<tail>.*)$")
NUMBER_RE = re.compile(r"\d{1,3}\.\d{1,3}")


def time_seconds(value):
    minutes, seconds = value.split("'")
    return round(int(minutes) * 60 + float(seconds), 3)


def extract_columns(pdf_path):
    chunks = []
    full_pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            full_pages.append(page.extract_text(x_tolerance=2, y_tolerance=3) or "")
            width, height = page.width, page.height
            top, bottom = 55, min(730, height - 90)
            boxes = [
                (30, top, width / 2 + 2, bottom),
                (width / 2 - 2, top, width - 25, bottom),
            ]
            for column_number, box in enumerate(boxes, 1):
                text = page.crop(box).extract_text(x_tolerance=2, y_tolerance=3) or ""
                chunks.append({
                    "page": page_number,
                    "column": column_number,
                    "text": text,
                })
    return chunks, full_pages


def detect_event(full_pages):
    first = full_pages[0]
    lines = [line.strip() for line in first.splitlines() if line.strip()]
    race = next((line for line in lines if "GRAND PRIX OF" in line), "MotoGP Grand Prix")
    circuit_match = re.search(r"(?m)^(.+?)\s+\d{4}\s*m\.?\s*$", first)
    circuit = circuit_match.group(1).strip() if circuit_match else ""
    date_match = re.search(r"(?m)^(.+?,\s+(?:Sunday|Saturday|Friday),\s+.+?\d{4})", first)
    event_date = date_match.group(1).strip() if date_match else ""
    return race, circuit, event_date


def joined_text(chunks):
    # Page/column headers may contain noise, but rider headers provide robust block boundaries.
    return "\n".join(chunk["text"] for chunk in chunks)


def parse_laps(block, total_laps):
    laps = {}
    for raw_line in block.splitlines():
        match = LAP_RE.match(raw_line)
        if not match:
            continue
        lap_no = int(match.group("lap"))
        if lap_no < 1 or lap_no > total_laps:
            continue
        lap_time_pdf = match.group("time")
        tail = match.group("tail")
        values = [float(value) for value in NUMBER_RE.findall(tail)]

        # Standard rows have T1, T2, T3, T4, speed. Interrupted laps can contain fewer values.
        sectors = [None, None, None, None]
        speed = None
        if len(values) >= 5:
            sectors = values[-5:-1]
            speed = values[-1]
        else:
            for index, value in enumerate(values[:4]):
                sectors[index] = value

        laps[lap_no] = {
            "lap": lap_no,
            "lap_time": lap_time_pdf.replace("'", ":"),
            "lap_seconds": time_seconds(lap_time_pdf),
            "t1": sectors[0],
            "t2": sectors[1],
            "t3": sectors[2],
            "t4": sectors[3],
            "speed_kmh": speed,
            "cancelled": "*" in raw_line,
            "pit_lane": bool(re.search(r"\bP\b", raw_line)),
            "partial_timing": len(values) < 5,
        }

    expected = list(range(1, total_laps + 1))
    missing = [lap for lap in expected if lap not in laps]
    if missing:
        raise ValueError(f"Missing lap rows: {missing}")

    cumulative = 0.0
    ordered = []
    for lap_no in expected:
        lap = laps[lap_no]
        cumulative = round(cumulative + lap["lap_seconds"], 3)
        lap["cumulative_seconds"] = cumulative
        ordered.append(lap)
    return ordered


def parse_rider_block(header_match, block):
    position_match = POSITION_RE.search(block)
    counts_match = COUNTS_RE.search(block)
    if not position_match or not counts_match:
        raise ValueError(f"Incomplete rider metadata for {header_match.group('name')}")

    lines = [line.strip() for line in block.splitlines() if line.strip()]
    position_line_index = next(
        index for index, line in enumerate(lines) if POSITION_RE.fullmatch(line)
    )
    team = lines[position_line_index + 1]

    tyres_match = TYRES_RE.search(block)
    total_laps = int(counts_match.group("total"))
    laps = parse_laps(block, total_laps)

    return {
        "name": header_match.group("name").strip(),
        "number": int(position_match.group("number")),
        "position": int(position_match.group("position")),
        "team": team,
        "manufacturer": header_match.group("manufacturer"),
        "nationality": header_match.group("nation"),
        "runs": int(counts_match.group("runs")),
        "total_laps": total_laps,
        "full_laps": int(counts_match.group("full")),
        "valid_laps": int(counts_match.group("valid")),
        "front_tyre": tyres_match.group("front") if tyres_match else None,
        "rear_tyre": tyres_match.group("rear") if tyres_match else None,
        "laps": laps,
    }


def parse(pdf_path):
    chunks, full_pages = extract_columns(pdf_path)
    text = joined_text(chunks)
    headers = list(HEADER_RE.finditer(text))
    if not headers:
        raise ValueError("No rider headers were detected")

    riders = []
    errors = []
    for index, header in enumerate(headers):
        end = headers[index + 1].start() if index + 1 < len(headers) else len(text)
        block = text[header.start():end]
        try:
            riders.append(parse_rider_block(header, block))
        except Exception as exc:
            errors.append(f"{header.group('name')}: {exc}")

    if errors:
        raise ValueError("Rider extraction failed:\n- " + "\n- ".join(errors))

    riders.sort(key=lambda item: item["position"])
    positions = [rider["position"] for rider in riders]
    expected_positions = list(range(1, len(riders) + 1))
    if positions != expected_positions:
        raise ValueError(f"Finishing positions are not continuous: {positions}")

    race, circuit, event_date = detect_event(full_pages)
    return {
        "source_pdf": Path(pdf_path).name,
        "race": race,
        "circuit": circuit,
        "event_date": event_date,
        "page_count": len(full_pages),
        "rider_count": len(riders),
        "lap_row_count": sum(len(rider["laps"]) for rider in riders),
        "riders": riders,
    }


def main():
    parser = argparse.ArgumentParser(description="Parse a MotoGP Chronological Analysis PDF")
    parser.add_argument("pdf", help="Path to the MotoGP Analysis PDF")
    parser.add_argument("-o", "--output", default="motogp_race.json")
    args = parser.parse_args()

    data = parse(args.pdf)
    Path(args.output).write_text(
        json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(
        f'Wrote {args.output}: {data["rider_count"]} riders, '
        f'{data["lap_row_count"]} lap rows'
    )


if __name__ == "__main__":
    main()
