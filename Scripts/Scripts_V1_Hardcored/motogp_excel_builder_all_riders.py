#!/usr/bin/env python3
"""Build a full MotoGP Excel workbook from the all-riders JSON."""
import argparse,json
from pathlib import Path
import xlsxwriter

COLORS=["#0070C0","#C00000","#00A651","#7030A0","#ED7D31","#5B9BD5","#A5A5A5",
"#FFC000","#4472C4","#70AD47","#264478","#9E480E","#636363","#997300","#255E91",
"#43682B","#6A5ACD","#DC143C","#008B8B","#FF1493","#8B4513"]

def safe_sheet(name): return name.title()[:31]

def build(json_path,output_path):
    data=json.loads(Path(json_path).read_text(encoding="utf-8")); wb=xlsxwriter.Workbook(output_path)
    wb.set_properties({"title":"MotoGP Lap Analysis","comments":"Generated from Analysis.pdf"})
    title=wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF","bg_color":"#1F4E78","align":"center"})
    header=wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#4472C4","border":1,"align":"center"})
    num3=wb.add_format({"num_format":"0.000","border":1}); num1=wb.add_format({"num_format":"0.0","border":1})
    integer=wb.add_format({"num_format":"0","border":1}); text=wb.add_format({"border":1})
    meta=wb.add_format({"bold":True,"bg_color":"#D9EAF7"}); time_fmt=wb.add_format({"num_format":"m:ss.000","border":1})
    dash_time=wb.add_format({"num_format":"m:ss.000"})

    for rider in data["riders"]:
        ws=wb.add_worksheet(safe_sheet(rider["name"])); ws.merge_range("A1:J1",f'{rider["name"]} #{rider["number"]} - Position {rider["position"]}',title)
        ws.write_row("A3",["Team",rider["team"],"Manufacturer",rider["manufacturer"],"Total Laps",rider["total_laps"],"Valid Laps",rider["valid_laps"]],meta)
        headers=["Lap","Lap Time","Lap Seconds","T1","T2","T3","T4","Speed km/h","Cumulative Seconds","Excel Time"]
        ws.write_row("A5",headers,header)
        for row,lap in enumerate(rider["laps"],5):
            ws.write_number(row,0,lap["lap"],integer); ws.write(row,1,lap["lap_time"],text)
            for col,key in enumerate(["lap_seconds","t1","t2","t3","t4"],2): ws.write_number(row,col,lap[key],num3)
            ws.write_number(row,7,lap["speed_kmh"],num1); ws.write_number(row,8,lap["cumulative_seconds"],num3)
            ws.write_number(row,9,lap["lap_seconds"]/86400.0,time_fmt)
        ws.freeze_panes(5,0); ws.set_column("A:A",8); ws.set_column("B:B",12); ws.set_column("C:J",18)
        ws.autofilter(4,0,4+len(rider["laps"]),9)

    allws=wb.add_worksheet("Race Data")
    all_headers=["Rider","Number","Position","Lap","Lap Time","Lap Seconds","T1","T2","T3","T4","Speed km/h","Cumulative Seconds","Cancelled","Pit Lane"]
    allws.write_row(0,0,all_headers,header); rr=1
    for rider in data["riders"]:
        for lap in rider["laps"]:
            allws.write_row(rr,0,[rider["name"],rider["number"],rider["position"],lap["lap"],lap["lap_time"],lap["lap_seconds"],lap["t1"],lap["t2"],lap["t3"],lap["t4"],lap["speed_kmh"],lap["cumulative_seconds"],lap["cancelled"],lap["pit_lane"]]); rr+=1
    allws.add_table(0,0,rr-1,len(all_headers)-1,{"name":"RaceData","columns":[{"header":h} for h in all_headers],"style":"Table Style Medium 2"})
    allws.freeze_panes(1,0); allws.set_column("A:A",24); allws.set_column("B:N",16)

    dash=wb.add_worksheet("Race Progress"); last_col=4+len(data["riders"])
    dash.merge_range(0,0,0,max(7,last_col),"MotoGP - Rider Lap Time Analysis",title)
    dash.write_row("A3",["Show","Rider","Color"],header)
    for i,rider in enumerate(data["riders"]):
        row=3+i
        if hasattr(dash,"insert_checkbox"): dash.insert_checkbox(row,0,True)
        else: dash.write_boolean(row,0,True)
        dash.write(row,1,rider["name"]); dash.write(row,2,COLORS[i%len(COLORS)])
    dash.write_row("E3",["Lap"]+[r["name"] for r in data["riders"]],header)
    max_laps=max(len(r["laps"]) for r in data["riders"])
    for i in range(max_laps):
        er=3+i; dash.write_number(er,4,i+1)
        for j,rider in enumerate(data["riders"]):
            src=safe_sheet(rider["name"]).replace("'","''")
            if i < len(rider["laps"]): formula=f'=IF($A${4+j},\'{src}\'!J{6+i},NA())'
            else: formula="=NA()"
            dash.write_formula(er,5+j,formula,None,"#N/A")
    dash.set_column(5,4+len(data["riders"]),14,dash_time)
    chart=wb.add_chart({"type":"line"})
    for j,rider in enumerate(data["riders"]):
        chart.add_series({"name":["Race Progress",2,5+j],"categories":["Race Progress",3,4,3+max_laps-1,4],
                          "values":["Race Progress",3,5+j,3+max_laps-1,5+j],
                          "line":{"color":COLORS[j%len(COLORS)],"width":2.0}})
    chart.set_title({"name":"Lap Time Progress"}); chart.set_x_axis({"name":"Lap"})
    chart.set_y_axis({"name":"Lap Time","num_format":"m:ss.000","major_gridlines":{"visible":True}})
    chart.set_legend({"position":"bottom"}); chart.set_size({"width":1100,"height":620})
    dash.insert_chart("A33",chart); dash.set_column("A:A",10); dash.set_column("B:B",26); dash.set_column("C:C",12); dash.freeze_panes(3,0)
    wb.close(); print(f"Wrote {output_path}")

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("json"); ap.add_argument("-o","--output",default="motogp_all_riders.xlsx")
    args=ap.parse_args(); build(args.json,args.output)
if __name__=="__main__": main()
