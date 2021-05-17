import openpyxl
from openpyxl import workbook

book = openpyxl.load_workbook('Test2.xlsx')
sheet = book.active

sheet.insert_cols(6, 2)
sheet.insert_cols(9, 2)
sheet.insert_cols(16, 10)

sheet.cell(row=1, column=6).value = "ΔALC比上次"
sheet.cell(row=3, column=6).value = "=E3-E2"
sheet.cell(row=1, column=7).value = "ΔALC比C0"

sheet.cell(row=1, column=9).value = "ΔANC比上次"
sheet.cell(row=3, column=9).value = "=H3-H2"
sheet.cell(row=1, column=10).value = "ΔANC比C0"

sheet.cell(row=1, column=16).value = "NLR"
sheet.cell(row=2, column=16).value = "=H2/E2"
sheet.cell(row=1, column=17).value = "ΔNLR比上次"
sheet.cell(row=3, column=17).value = "=P3-P2"
sheet.cell(row=1, column=18).value = "ΔNLR比基线"

sheet.cell(row=1, column=19).value = "dNLR"
sheet.cell(row=2, column=19).value = "=H2/(D2-H2)"

sheet.cell(row=1, column=20).value = "LMR"
sheet.cell(row=2, column=20).value = "=E2/L2"

sheet.cell(row=1, column=21).value = "ΔLMR比上次"
sheet.cell(row=3, column=21).value = "=T3-T2"
sheet.cell(row=1, column=22).value = "ΔLMR比基线"

sheet.cell(row=1, column=23).value = "PLR"
sheet.cell(row=2, column=23).value = "=M2/E2"
sheet.cell(row=1, column=24).value = "ΔPLR比上次"
sheet.cell(row=3, column=24).value = "=W3-W2"
sheet.cell(row=1, column=25).value = "ΔPLR比基线"

sheet.cell(row=1, column=27).value = "ΔAMC比上次"
sheet.cell(row=3, column=27).value = "=L3-L2"
sheet.cell(row=1, column=28).value = "ΔAMC比C0"

sheet.insert_rows(2, 1)

book.save('result.xlsx')




