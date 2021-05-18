import openpyxl
from datetime import datetime
from openpyxl.utils import get_column_letter

print("输入文件名：")
fileName = input()

book = openpyxl.load_workbook(fileName + '.xlsx')
sheet = book.active

sheet.insert_cols(6, 2)
sheet.insert_cols(9, 2)
sheet.insert_cols(16, 10)

sheet.insert_rows(2, 1)

print("输入治疗日期：")
cureDateStr = input()

index = 0

cureDate = datetime.strptime(cureDateStr, '%Y/%m/%d')

for n in range(3, sheet.max_column):
    if sheet.cell(row=n, column=3).value <= cureDate <= sheet.cell(row=n + 1, column=3).value:
        index = n

sheet.cell(row=index, column=2).value = cureDate

for n in range(1, 12):
    sheet.cell(row=2, column=n + 3).value = n
for m in range(12, 25):
    sheet.cell(row=2, column=m + 4).value = m

for n in range(index - 1, 2, -1):
    sheet.cell(row=n, column=1).value = "C" + "-" + str(index - n)

for n in range(index, sheet.max_row + 1):
    sheet.cell(row=n, column=1).value = "C" + str(n - index + 1)

sheet.cell(row=1, column=6).value = "ΔALC比上次"
m = 0
for n in range(4, sheet.max_row + 1):
    if sheet.cell(row=n-1, column=5).value is not None:
        m = n-1

    if sheet.cell(row=n, column=5).value is None:
        sheet.cell(row=n, column=6).value = ""
    elif sheet.cell(row=n-1, column=5).value is None:
        sheet.cell(row=n, column=6).value = "=E" + str(n) + "-E" + str(m)
    else:
        sheet.cell(row=n, column=6).value = "=E" + str(n) + "-E" + str(n - 1)

sheet.cell(row=1, column=7).value = "ΔALC比C0"
for n in range(index + 1, sheet.max_row + 1):
    if sheet.cell(row=n, column=5).value is None:
        sheet.cell(row=n, column=7).value = ""
    else:
        sheet.cell(row=n, column=7).value = "=E" + str(n) + "-" + str(sheet.cell(row=index, column=5).value)

sheet.cell(row=1, column=9).value = "ΔANC比上次"
m = 0
for n in range(4, sheet.max_row + 1):
    if sheet.cell(row=n-1, column=8).value is not None:
        m = n-1

    if sheet.cell(row=n, column=8).value is None:
        sheet.cell(row=n, column=9).value = ""
    elif sheet.cell(row=n-1, column=8).value is None:
        sheet.cell(row=n, column=9).value = "=H" + str(n) + "-H" + str(m)
    else:
        sheet.cell(row=n, column=9).value = "=H" + str(n) + "-H" + str(n - 1)

sheet.cell(row=1, column=10).value = "ΔANC比C0"
for n in range(index + 1, sheet.max_row + 1):
    if sheet.cell(row=n, column=8).value is None:
        sheet.cell(row=n, column=10).value = ""
    else:
        sheet.cell(row=n, column=10).value = "=H" + str(n) + "-" + str(sheet.cell(row=index, column=8).value)

sheet.cell(row=1, column=16).value = "NLR"
for n in range(3, sheet.max_row + 1):
    if sheet.cell(row=n, column=5).value is None:
        sheet.cell(row=n, column=16).value = ""
    else:
        sheet.cell(row=n, column=16).value = "=H" + str(n) + "/E" + str(n)

sheet.cell(row=1, column=17).value = "ΔNLR比上次"
m = 0
for n in range(4, sheet.max_row + 1):
    if sheet.cell(row=n-1, column=16).value != "":
        m = n-1

    if sheet.cell(row=n, column=16).value == "":
        sheet.cell(row=n, column=17).value = ""
    elif sheet.cell(row=n-1, column=16).value == "":
        sheet.cell(row=n, column=17).value = "=P" + str(n) + "-P" + str(m)
    else:
        sheet.cell(row=n, column=17).value = "=P" + str(n) + "-P" + str(n - 1)

sheet.cell(row=1, column=18).value = "ΔNLR比基线"
for n in range(index + 1, sheet.max_row + 1):
    if sheet.cell(row=n, column=16).value == "":
        sheet.cell(row=n, column=18).value = ""
    else:
        sheet.cell(row=n, column=18).value = "=P" + str(n) + "-(H" + str(index) + "/E" + str(index) + ")"

sheet.cell(row=1, column=19).value = "dNLR"
for n in range(3, sheet.max_row + 1):
    if sheet.cell(row=n, column=4).value is None:
        sheet.cell(row=n, column=19).value = ""
    elif sheet.cell(row=n, column=8).value is None:
        sheet.cell(row=n, column=19).value = ""
    else:
        sheet.cell(row=n, column=19).value = "=H" + str(n) + "/(D" + str(n) + "-H" + str(n) + ")"

sheet.cell(row=1, column=20).value = "LMR"
for n in range(3, sheet.max_row + 1):
    if sheet.cell(row=n, column=12).value is None:
        sheet.cell(row=n, column=20).value = ""
    else:
        sheet.cell(row=n, column=20).value = "=E" + str(n) + "/L" + str(n)

m = 0
sheet.cell(row=1, column=21).value = "ΔLMR比上次"
for n in range(4, sheet.max_row + 1):
    if sheet.cell(row=n-1, column=20).value != "":
        m = n-1

    if sheet.cell(row=n, column=20).value == "":
        sheet.cell(row=n, column=21).value = ""
    elif sheet.cell(row=n-1, column=20).value == "":
        sheet.cell(row=n, column=21).value = "=T" + str(n) + "-T" + str(m)
    else:
        sheet.cell(row=n, column=21).value = "=T" + str(n) + "-T" + str(n - 1)

sheet.cell(row=1, column=22).value = "ΔLMR比基线"
for n in range(index + 1, sheet.max_row + 1):
    if sheet.cell(row=n, column=20).value == "":
        sheet.cell(row=n, column=22).value = ""
    else:
        sheet.cell(row=n, column=22).value = "=T" + str(n) + "-(E" + str(index) + "/L" + str(index) + ")"

sheet.cell(row=1, column=23).value = "PLR"
for n in range(3, sheet.max_row + 1):
    if sheet.cell(row=n, column=5).value is None:
        sheet.cell(row=n, column=23).value = ""
    else:
        sheet.cell(row=n, column=23).value = "=M" + str(n) + "/E" + str(n)

sheet.cell(row=1, column=24).value = "ΔPLR比上次"
m = 0
for n in range(4, sheet.max_row + 1):
    if sheet.cell(row=n-1, column=23).value != "":
        m = n-1

    if sheet.cell(row=n, column=23).value == "":
        sheet.cell(row=n, column=24).value = ""
    elif sheet.cell(row=n-1, column=23).value == "":
        sheet.cell(row=n, column=24).value = "=W" + str(n) + "-W" + str(m)
    else:
        sheet.cell(row=n, column=24).value = "=W" + str(n) + "-W" + str(n - 1)

sheet.cell(row=1, column=25).value = "ΔPLR比基线"
for n in range(index + 1, sheet.max_row + 1):
    if sheet.cell(row=n, column=23).value == "":
        sheet.cell(row=n, column=25).value = ""
    else:
        sheet.cell(row=n, column=25).value = "=W" + str(n) + "-(M" + str(index) + "/E" + str(index) + ")"

m = 0
sheet.cell(row=1, column=27).value = "ΔAMC比上次"
for n in range(4, sheet.max_row + 1):
    if sheet.cell(row=n-1, column=12).value is not None:
        m = n-1

    if sheet.cell(row=n, column=12).value is None:
        sheet.cell(row=n, column=27).value = ""
    elif sheet.cell(row=n-1, column=12).value is None:
        sheet.cell(row=n, column=27).value = "=L" + str(n) + "-L" + str(m)
    else:
        sheet.cell(row=n, column=27).value = "=L" + str(n) + "-L" + str(n - 1)

sheet.cell(row=1, column=28).value = "ΔAMC比C0"
for n in range(index + 1, sheet.max_row + 1):
    if sheet.cell(row=n, column=12).value is None:
        sheet.cell(row=n, column=28).value = ""
    else:
        sheet.cell(row=n, column=28).value = "=L" + str(n) + "-" + str(sheet.cell(row=index, column=12).value)

sheet.column_dimensions[get_column_letter(3)].width = 13
sheet.column_dimensions[get_column_letter(15)].width = 13

book.save(fileName + '结果' + '.xlsx')
