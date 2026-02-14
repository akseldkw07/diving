# from posixpath import split
import pandas as pd
from openpyxl import load_workbook as lw


def main():
    dive_csv_path = "/Users/Akseldkw/Desktop/Aksel/Diving/Activities.csv"
    dive_log_path = "/Users/Akseldkw/Desktop/Aksel/Diving/Scuba_Log.xlsx"
    wb = lw(dive_log_path)
    ws = wb["Log"]
    print(ws.max_row)
    print(len(ws["A"]))
    print(ws[f"A{ws.max_row}"].value)
    index_start = ws[f"A{ws.max_row}"].value + 1
    new_dives = pd.read_csv(dive_csv_path)

    columns = [
        "Date",
        "0",
        "1",
        "2",
        "3",
        "4",
        "Max Depth",
        "Min Water Temp",
        "5",
        "Dive Time In",
        "Gas Type",
        "Time",
        "Surface Interval",
        "7",
        "Avg HR",
        "Max HR",
    ]
    new_dives["Date"] = pd.to_datetime(new_dives["Date"])
    new_dives.sort_values(by=["Date"], ascending=True, inplace=True)
    new_dives["Surface Interval"] = pd.to_timedelta(
        new_dives["Surface Interval"].apply(lambda x: ("0:" + x) if (x.count(":") < 2) else x)
    ).astype(str)
    new_dives["Dive Time In"] = new_dives["Date"].dt.strftime("%I:%M %p")
    new_dives["Date"] = new_dives["Date"].dt.strftime("%-m/%-d/%Y")
    new_dives["Time"] = (pd.to_timedelta(new_dives["Time"]).dt.total_seconds() / 60).round(0)

    new_dives.reset_index(inplace=True)
    for index in range(8):
        new_dives[str(index)] = ""

    new_dives = new_dives[columns]
    new_dives.index += index_start
    new_dives.reset_index(inplace=True)
    print(new_dives.info())
    print(new_dives)

    with pd.ExcelWriter(path=dive_log_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        start_row = len(ws["A"])
        new_dives.to_excel(writer, sheet_name="Log", index=False, startrow=start_row, header=False)


if __name__ == "__main__":
    main()
