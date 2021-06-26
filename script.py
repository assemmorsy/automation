import openpyxl
workBook = openpyxl.load_workbook('sample.xlsx')
workSheet = workBook.active

class MasterSheet:
    def __init__(self):
        self.cases = []

    def load_cases(self,work_sheet):
        for row in range(2, work_sheet.max_row+1):
            case = {}
            for col in range(1, work_sheet.max_column + 1):
                case[work_sheet.cell(row=1, column=col).value] = work_sheet.cell(row=row, column=col).value
            self.cases.append(case)

    def mean_for_group(self , target_var , group_number):
        pass
