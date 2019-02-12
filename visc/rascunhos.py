import xlsxwriter

def planilha():
    workbook = xlsxwriter.Workbook('teste.xlsx')
    return workbook

def sheet(workbook):
    worksheet = workbook.add_worksheet('plan1')
    return worksheet

workbook = planilha()
sheet(workbook)
workbook.close()