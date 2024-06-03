import argparse
import json
import os

import pandas as pd


def get_raw_story_json(input):
    with open(input, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["importGridList"]


def get_sheets_from_json(data):
    sheets = {}
    # For each Sheet
    for sheet in data:
        # Get Sheet Name
        # The "name" key is usually "Assets/PathTo/Sheets.xls:SheetName"
        sheet_name = sheet["name"].split(":")[-1]
        print("Reading sheet: " + sheet_name)
        # Create Sheet
        sheets[sheet_name] = []
        # Get all rows
        for row in sheet["rows"]:
            # Add row to Sheet
            sheets[sheet_name].append(row["strings"])
            # There's actually a `rowIndex` in rows,
            # but they are always in order, so we don't need it currently.
    return sheets


def write_sheets_to_xlsx(sheets, output="Output.xlsx"):
    # Create empty DataFrame
    df = pd.DataFrame()
    # Write an empty sheet to create file
    df.to_excel(output, sheet_name="Sheet1", index=False, header=False)
    with pd.ExcelWriter(
        output,
        engine="openpyxl",
        mode="a",  # Append sheets to the file
        if_sheet_exists="replace",
    ) as writer:
        for sheet_name in sheets:
            print("Writing sheet: " + sheet_name)
            df = pd.DataFrame(sheets[sheet_name])
            df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=str)
    args = parser.parse_args()

    abs_path = os.path.abspath(args.input)
    filename = os.path.basename(abs_path)
    pureFilename = os.path.splitext(filename)[0]

    data = get_raw_story_json(abs_path)
    sheets = get_sheets_from_json(data)
    write_sheets_to_xlsx(sheets, f"{pureFilename}.xlsx")

    print("Done!")
