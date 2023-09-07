import openpyxl
from main import parse_input_workbook_sheet, write_output
from datetime import datetime, date

def test_parse_input_workbook_sheet():
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append([date(2021,1,1), "", ""])
    sheet.append([date(2021,1,31), "Page Views", "Page Views"])
    sheet.append(["Site ID", date(2021,1,1), date(2021,1,2)])
    sheet.append(["site 1", 6, 4])
    parsed = parse_input_workbook_sheet(sheet)
    assert "site 1" in parsed.keys()
    assert parsed["site 1"][date(2021, 1, 1)]["Page Views"] == 6


def test_write_output():
    data = {"site 1": {datetime(2021, 1, 1): {"Page Views": 6, "Unique Visitors": 2, "Total Time Spent": 12, "Visits": 1, "Average Time Spent on Site": 0.5},
                       datetime(2021, 1, 2): {"Page Views": 5, "Unique Visitors": 3, "Total Time Spent": 13, "Visits": 2, "Average Time Spent on Site": 0.6}},
            "site 2": {datetime(2021, 1, 1): {"Page Views": 4, "Unique Visitors": 4, "Total Time Spent": 14, "Visits": 3, "Average Time Spent on Site": 0.7},
                       datetime(2021, 1, 2): {"Page Views": 3, "Unique Visitors": 5, "Total Time Spent": 15, "Visits": 4, "Average Time Spent on Site": 0.8}}}
    wb = write_output(data)
    sheet = wb["output"]
    assert sheet["A1"].value == "Day of Month"
    assert sheet["C2"].value == "site 1"
    assert sheet["D3"].value == 5