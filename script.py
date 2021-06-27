import openpyxl

class MasterSheet:
    def __init__(self, work_book):
        self.work_book = work_book
        self.groups = {}
        self.cases = {}
        self.load_group_names()

    def load_group_names(self):
        work_sheet = self.work_book['C']
        for row in range(2, work_sheet.max_row+1):
            self.groups[work_sheet.cell(row=row, column=1).value] = work_sheet.cell(row=row, column=2).value.strip()

    def load_cases(self):
        work_sheet = self.work_book['Master']

        for group in self.groups:
            self.cases[self.groups[group]] = []

        for row in range(2, work_sheet.max_row+1):
            case = {}
            for col in range(1, work_sheet.max_column + 1):
                case[work_sheet.cell(row=1, column=col).value] = work_sheet.cell(row=row, column=col).value
            self.cases[self.groups[case['G']]].append(case)

def main():
    workBook = openpyxl.load_workbook('sample_test.xlsx')
    x = MasterSheet(workBook)
    x.load_cases()
    for g in x.groups:
        print(len(x.cases[x.groups[g]]))

main()
