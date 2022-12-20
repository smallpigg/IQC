import xlrd
import xlwt
import xlsxwriter

path  = "文件记录更改通知单-Template.xls"

wb = xlrd.open_workbook(path)
ws = wb.sheet_by_index(0)

ws[4][0].value = "文件记录编号：AAA-IQC-222"

print(ws[4][0].value)

wb1 = xlsxwriter.Workbook(path)
ws1 = wb1.sheet_by_index(0)

