#!/usr/bin/env python3
"""Extract all riders and lap records from the MotoGP chronological analysis PDF."""
import argparse, json, re
from pathlib import Path
import pdfplumber

RIDER_META = [
("Marco BEZZECCHI",72,1,"Aprilia Racing","APRILIA",27,26),
("Raul FERNANDEZ",25,2,"SuperFile Trackhouse MotoGP Team","APRILIA",27,26),
("Jorge MARTIN",89,3,"Aprilia Racing","APRILIA",27,26),
("Ai OGURA",79,4,"SuperFile Trackhouse MotoGP Team","APRILIA",27,26),
("Pedro ACOSTA",37,5,"Red Bull KTM Factory Racing","KTM",27,26),
("Marc MARQUEZ",93,6,"Ducati Lenovo Team","DUCATI",27,26),
("Alex MARQUEZ",73,7,"BK8 Gresini Racing MotoGP","DUCATI",27,25),
("Fabio DI GIANNANTONIO",49,8,"Pertamina Enduro VR46 Racing Team","DUCATI",27,25),
("Fabio QUARTARARO",20,9,"Monster Energy Yamaha MotoGP Team","YAMAHA",27,24),
("Enea BASTIANINI",23,10,"Red Bull KTM Tech3","KTM",27,26),
("Brad BINDER",33,11,"Red Bull KTM Factory Racing","KTM",27,26),
("Luca MARINI",10,12,"Honda HRC Castrol","HONDA",27,25),
("Johann ZARCO",5,13,"Castrol Honda LCR","HONDA",27,25),
("Franco MORBIDELLI",21,14,"Pertamina Enduro VR46 Racing Team","DUCATI",27,24),
("Alex RINS",42,15,"Monster Energy Yamaha MotoGP Team","YAMAHA",27,26),
("Diogo MOREIRA",11,16,"Pro Honda LCR","HONDA",27,23),
("Maverick VIÑALES",12,17,"Red Bull KTM Tech3","KTM",27,22),
("Toprak RAZGATLIOGLU",7,18,"Prima Pramac Yamaha MotoGP","YAMAHA",8,5),
("Francesco BAGNAIA",63,19,"Ducati Lenovo Team","DUCATI",5,4),
("Jack MILLER",43,20,"Prima Pramac Yamaha MotoGP","YAMAHA",5,4),
("Fermin ALDEGUER",54,21,"BK8 Gresini Racing MotoGP","DUCATI",1,0),
]

# Each rider can span more than one page/column. Coordinates match this official timing layout.
SEGMENTS = {
"Marco BEZZECCHI":[(1,"L","Marco BEZZECCHI","Raul FERNANDEZ")],
"Raul FERNANDEZ":[(1,"L","Raul FERNANDEZ",None),(1,"R",None,"Jorge MARTIN")],
"Jorge MARTIN":[(1,"R","Jorge MARTIN",None),(2,"L",None,"Ai OGURA")],
"Ai OGURA":[(2,"L","Ai OGURA","Pedro ACOSTA")],
"Pedro ACOSTA":[(2,"L","Pedro ACOSTA",None),(2,"R",None,"Marc MARQUEZ")],
"Marc MARQUEZ":[(2,"R","Marc MARQUEZ",None),(3,"L",None,"Alex MARQUEZ")],
"Alex MARQUEZ":[(3,"L","Alex MARQUEZ","Fabio DI GIANNANTONIO")],
"Fabio DI GIANNANTONIO":[(3,"L","Fabio DI GIANNANTONIO",None),(3,"R",None,"Fabio QUARTARARO")],
"Fabio QUARTARARO":[(3,"R","Fabio QUARTARARO","Enea BASTIANINI")],
"Enea BASTIANINI":[(3,"R","Enea BASTIANINI",None),(4,"L",None,"Brad BINDER")],
"Brad BINDER":[(4,"L","Brad BINDER",None)],
"Luca MARINI":[(4,"R","Luca MARINI","Johann ZARCO")],
"Johann ZARCO":[(4,"R","Johann ZARCO",None),(5,"L",None,"Franco MORBIDELLI")],
"Franco MORBIDELLI":[(5,"L","Franco MORBIDELLI","Alex RINS")],
"Alex RINS":[(5,"L","Alex RINS",None),(5,"R",None,"Diogo MOREIRA")],
"Diogo MOREIRA":[(5,"R","Diogo MOREIRA",None),(6,"L",None,"Maverick VIÑALES")],
"Maverick VIÑALES":[(6,"L","Maverick VIÑALES","Toprak RAZGATLIOGLU")],
"Toprak RAZGATLIOGLU":[(6,"L","Toprak RAZGATLIOGLU","Francesco BAGNAIA")],
"Francesco BAGNAIA":[(6,"L","Francesco BAGNAIA",None),(6,"R",None,"Jack MILLER")],
"Jack MILLER":[(6,"R","Jack MILLER","Fermin ALDEGUER")],
"Fermin ALDEGUER":[(6,"R","Fermin ALDEGUER",None)],
}

ROW_RE = re.compile(r"^\s*(\d{1,2})\s+(\d+'\d{2}\.\d{3})\s+(.+?)\s*$")
NUM_RE = re.compile(r"\d{2,3}\.\d{1,3}")

def lap_seconds(value):
    m, s = value.split("'")
    return round(int(m)*60 + float(s), 3)

def cut(text,start,end):
    if start:
        if start not in text: raise ValueError(f"Start marker not found: {start}")
        text=text.split(start,1)[1]
    if end:
        if end not in text: raise ValueError(f"End marker not found: {end}")
        text=text.split(end,1)[0]
    return text

def parse_rows(text):
    out={}
    for line in text.splitlines():
        m=ROW_RE.match(line)
        if not m: continue
        lap=int(m.group(1)); lt=m.group(2); tail=m.group(3)
        nums=[float(x) for x in NUM_RE.findall(tail)]
        if len(nums) < 5: continue
        # Last five numeric values are T1,T2,T3,T4,Speed, even when * or P markers occur.
        t1,t2,t3,t4,speed=nums[-5:]
        out[lap]={
            "lap":lap,"lap_time":lt.replace("'",":"),"lap_seconds":lap_seconds(lt),
            "t1":t1,"t2":t2,"t3":t3,"t4":t4,"speed_kmh":speed,
            "cancelled":"*" in line,"pit_lane":bool(re.search(r"\bP\b",line)),
        }
    return out

def parse(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 6: raise ValueError("Expected a six-page chronological analysis PDF")
        chunks={}
        for page_no in range(1,7):
            page=pdf.pages[page_no-1]
            for side,bbox in {"L":(30,55,300,730),"R":(300,55,570,730)}.items():
                chunks[(page_no,side)]=page.crop(bbox).extract_text(x_tolerance=2,y_tolerance=3) or ""
    riders=[]
    for name,number,position,team,mfr,total_laps,valid_laps in RIDER_META:
        text="\n".join(cut(chunks[(p,s)],start,end) for p,s,start,end in SEGMENTS[name])
        lap_map=parse_rows(text)
        expected=list(range(1,total_laps+1))
        if sorted(lap_map)!=expected:
            raise ValueError(f"{name}: expected laps {expected}, extracted {sorted(lap_map)}")
        cumulative=0.0
        laps=[]
        for lap_no in expected:
            lap=lap_map[lap_no]
            cumulative=round(cumulative+lap["lap_seconds"],3)
            lap["cumulative_seconds"]=cumulative
            laps.append(lap)
        riders.append({"name":name,"number":number,"position":position,"team":team,
                       "manufacturer":mfr,"total_laps":total_laps,"valid_laps":valid_laps,"laps":laps})
    return {"source_pdf":Path(pdf_path).name,"race":"Pertamina Grand Prix of Indonesia",
            "circuit":"Pertamina Mandalika Circuit","rider_count":len(riders),"riders":riders}

def main():
    ap=argparse.ArgumentParser(description="Extract all MotoGP riders to JSON")
    ap.add_argument("pdf"); ap.add_argument("-o","--output",default="motogp_all_riders.json")
    args=ap.parse_args(); data=parse(args.pdf)
    Path(args.output).write_text(json.dumps(data,indent=2,ensure_ascii=False),encoding="utf-8")
    print(f'Wrote {args.output}: {len(data["riders"])} riders, {sum(len(r["laps"]) for r in data["riders"])} lap rows')
if __name__=="__main__": main()
