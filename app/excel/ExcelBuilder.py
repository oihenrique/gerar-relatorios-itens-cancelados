import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Alignment
from openpyxl.utils.cell import coordinate_from_string


class ExcelBuilder:
    def __init__(self, file_name):
        self.file_name = file_name
        self.workbook = openpyxl.load_workbook(f'{self.file_name}.xlsx')

    def add_column_to_excel(self, values):
        worksheet = self.workbook['Dados']

        for i, valor in enumerate(values):
            worksheet.cell(row=i + 2, column=16, value=values[i])

        self.save()

    def add_countif_formula(self, target_column, target_value, result_cell):
        worksheet = self.workbook['Dados']
        formula = f'=COUNTIF({get_column_letter(target_column)}:{get_column_letter(target_column)},{target_value})'

        cell = worksheet[result_cell]
        cell.value = formula
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.number_format = '0'

    def add_if_formula(self, sheet_name, target_column, target_row, condition_column, condition_row,
                       if_sucess, if_fail, result_cell):
        worksheet = self.workbook[sheet_name]
        formula = (f'=IF({get_column_letter(target_column)}{target_row}='
                   f'{get_column_letter(condition_column)}{condition_row},"{if_sucess}","{if_fail}")')

        cell = worksheet[result_cell]
        cell.value = formula
        cell.alignment = Alignment(horizontal='center', vertical='center')

        self.save()

    def add_title_header(self, sheet_name, header_titles):
        worksheet = self.workbook[sheet_name]

        for cell_coord, header_text in header_titles.items():
            col, row = coordinate_from_string(cell_coord)

            cell = worksheet.cell(row=row, column=column_index_from_string(col))
            cell.value = header_text

        self.save()

    def save(self):
        self.workbook.save(f'{self.file_name}.xlsx')

    def close(self):
        self.workbook.close()
