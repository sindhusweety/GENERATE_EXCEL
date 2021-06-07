#!/usr/bin/python3.9

#import required libraries
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.borders import Border, Side

#openpyxl handles excelsheet
wb = Workbook()
ws = wb.active


class Generate_Balance_Sheet:

    def __init__(self, **kwargs):
        self.kwargs = kwargs
        self.excel_data = list()

    def read_excels(self):
        excel_pathlist = [('/home/sindhukumari/PycharmProjects/Upwork/Download/Open_Status_Report.XLSX', 47, 6),
                          ('/home/sindhukumari/PycharmProjects/Upwork/Download/Balance_sheet.XLSX', 34, 2)]
        for path, r, c in excel_pathlist:
            wb_obj = load_workbook(path)
            sheet_obj = wb_obj.active

            cell_obj = sheet_obj.cell(row=r, column=c)
            self.excel_data.append(cell_obj.value)

    def generate_excel(self):
        # ***********SET BACKGROUND COLOR********************************#
        for rows in ws.iter_rows(min_row=1, max_row=34, min_col=1, max_col=8):
            for cell in rows:
                cell.fill = PatternFill(start_color='EEEEEF', end_color='EEEEEF', fill_type="solid")

        # **************SET BOLD & SIZE & STYLE*************************#
        fontStyle = Font(name='Tahoma', size=20, bold=True, italic=False, vertAlign=None,
                                 underline='none', strike=False, color='FF000000')
        ws.cell(row=1, column=3, value= "BALANCE SHEET RECONCILIATION").font = fontStyle

        ws['A3'] = "Entity ID"
        ws['A4'] = "Account Number"
        ws['A5'] = "Account Description"
        ws['A6'] = "Period"

        ws['A8'] = "General Ledger Balance"
        ws['A10'] = "Supporting or Subsidiary System Detail"
        ws['A12'] = "link"

        ws['D3'] = "Prepared By"
        ws['D4'] = "Reviewed By"
        ws['F3'] = "Date"
        ws['F4'] = "Date"

        ws['E15'] = "Total Supporting or Subsidiary System Balance"
        ws['E17'] = "Variance"

        ws['E27'] = "Total of Reconciling Items"
        ws['E29'] = "Unreconciled Balance (Should be 0)"
        ws['A31'] = "Additional Explanatory Notes"

        #***************** FIX THE CELL WIDTH & HEIGHT********************#
        ws.cell(row=1, column=1).value = ""
        ws.row_dimensions[1].height = 30
        ws.cell(row=2, column=1).value = ""
        ws.column_dimensions['A'].width = 40
        ws.cell(row=2, column=2).value = ""
        ws.column_dimensions['B'].width = 20
        ws.cell(row=2, column=3).value = ""
        ws.column_dimensions['C'].width = 15
        ws.cell(row=2, column=4).value = ""
        ws.column_dimensions['D'].width = 20
        ws.cell(row=2, column=5).value = ""
        ws.column_dimensions['E'].width = 25
        ws.cell(row=2, column=6).value = ""
        ws.column_dimensions['F'].width = 20
        ws.cell(row=2, column=7).value = ""
        ws.column_dimensions['G'].width = 30

        #******************SET COLOR**********************#
        graycolor = PatternFill(start_color='D0D0D0',
                              end_color='D0D0D0',
                              fill_type='solid')
        ws['G8'].fill = graycolor

        ws['E12'].fill = graycolor
        ws['E13'].fill = graycolor
        ws['E14'].fill = graycolor

        ws['G15'].fill = graycolor
        ws['G17'].fill = graycolor

        ws['E20'].fill = graycolor
        ws['E21'].fill = graycolor
        ws['E22'].fill = graycolor
        ws['E23'].fill = graycolor
        ws['E24'].fill = graycolor
        ws['E25'].fill = graycolor

        ws['G27'].fill = graycolor
        ws['G29'].fill = graycolor

        #**************SET BOLD*************************#
        ws['B1'].font = Font(bold=True)
        ws['A8'].font = Font(bold=True)
        ws['A10'].font = Font(bold=True)
        ws['E15'].font = Font(bold=True)
        ws['E27'].font = Font(bold=True)
        ws['E29'].font = Font(bold=True)
        ws['A31'].font = Font(bold=True)

        #*************SET BORDER*********************#
        thin_border = Border(bottom=Side(style='thin'))
        ws.cell(row=3, column=2).border = thin_border
        ws.cell(row=4, column=2).border = thin_border
        ws.cell(row=5, column=2).border = thin_border
        ws.cell(row=6, column=2).border = thin_border

        ws.cell(row=3, column=5).border = thin_border
        ws.cell(row=4, column=5).border = thin_border
        ws.cell(row=3, column=7).border = thin_border
        ws.cell(row=4, column=7).border = thin_border

        ws.cell(row=8, column=7).border = thin_border
        ws.cell(row=15, column=7).border = thin_border
        ws.cell(row=17, column=7).border = thin_border
        ws.cell(row=27, column=7).border = thin_border
        ws.cell(row=29, column=7).border = thin_border

        #*********************PASS VALUES**********************#
        ws['E12'] = self.excel_data[0] #OPEN REPORT
        ws['G15'] = self.excel_data[0] #OPEN REPORT
        ws['G8']  = self.excel_data[1] #BALANCE SHEET
        ws['G17'] = self.excel_data[1] - self.excel_data[0]
        wb.save("Balance-Sheet-Reconciliation.xlsx")

gObj = Generate_Balance_Sheet()
gObj.read_excels()
gObj.generate_excel()





