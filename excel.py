import openpyxl
import os

def generate_workbook(workbook_file_path):

    wb = openpyxl.Workbook()

    ws = wb.create_sheet("Test Cases")
    fill = openpyxl.styles.PatternFill("solid", fgColor="DDDDDD")
    ws.cell(row=1, column=1, value="Test Case ID").fill = fill
    ws.cell(row=1, column=2, value="Active").fill = fill
    ws.cell(row=1, column=2, value="Test Case Name").fill = fill
    ws.cell(row=1, column=2, value="Test Case Description").fill = fill

    # remove default sheet
    del wb["Sheet"]
    #wb.remove("Sheet")

    wb.save(workbook_file_path)

