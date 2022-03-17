from numpy import save
from openpyxl import load_workbook

import win32com.client

# Fills in ATE excel template of file type .xlsx
# results must be index according to the cells of the template
# that is, cells are filled in starting from the first index of results
# Adapted from 

def write_results(template_path, results_path, data_names, data_values):
    report = load_workbook(template_path)
    report.template = False
    sheet = report["A5 form"]

    i = 0
    for r in range(2, sheet.max_row+1):
        for c in range(3, sheet.max_column+1):
            val = sheet.cell(r, c).value
            if val != None and "$!" in str(val):
                for idx, name in enumerate(data_names):
                    if name == str(val):
                        if "." not in data_values[idx]:
                            sheet.cell(r, c).value = data_values[idx]
                        else:
                            sheet.cell(r, c).value = float(data_values[idx])
            i += 1

    # i = 0
    # for r in range(2, sheet.max_row+1):
    #     for c in range(3, sheet.max_column+1):
    #         val = sheet.cell(r, c).value

    #         if val != None and "$!" in str(val):
    #             if "." not in results[i]:
    #                 sheet.cell(r, c).value = results[i]

    #             else:
    #                 sheet.cell(r, c).value = float(results[i])
    #             i += 1


    report.save(results_path)


# Generates a PDF from a given results_path xlsx file
def generate_pdf(results_path, pdf_path):
    excel = win32com.client.Dispatch("Excel.Application")

    final_report = excel.Workbooks.Open(results_path)
    sheet_indeces = [1]

    final_report.WorkSheets(sheet_indeces).Select()
    final_report.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
    final_report.Close()


data_names = 	['$!Test_Date', '$!Tested_By', '$!Part_No', '$!6.4.3.', '$!6.5.5.']
data_values =	['March 17, 2022', 'Tristan Lee', '123', 'P', '0.0']

template = r'C:\Users\trist\Desktop\Desktop\UBC Engineering\Coop\Alpha\Work\Python\PyLab_PDF\templates\0100048-A5_R_1.xltx'
ex_result = r'C:\Users\trist\Desktop\Desktop\UBC Engineering\Coop\Alpha\Work\Python\PyLab_PDF\templates\results.xlsx'
pdf = r'C:\Users\trist\Desktop\Desktop\UBC Engineering\Coop\Alpha\Work\Python\PyLab_PDF\templates\0100048-A5_R_1.pdf'
results = ["Tristan Lee", "0100044-001", "123456789", "Power-ATE-BB", "March 16, 2022", "P", "P", "P", "P", "P", "54.00", "0.0", "P", "P",
           "P", "P", "58.01", "47.9", "54.08", "-0.3", "1.00", "-0.5", "52.00", "45.00", "=F36*F35", "50.00", "36.00", "P", "XLTX TEST"]

write_results(template, ex_result, data_names, data_values)
generate_pdf(ex_result, pdf)
