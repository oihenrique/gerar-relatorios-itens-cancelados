import openpyxl
import pyperclip as pc
import pandas as pd
import io


class ExcelProcessor:
    def __init__(self, file_name, filename_workbook=False):
        self.file_name = file_name
        if filename_workbook:
            self.workbook = openpyxl.load_workbook(f'{self.file_name}.xlsx')

    @staticmethod
    def read_clipboard_to_dataframe():
        data = pc.paste()
        return pd.read_csv(io.StringIO(data), sep='\t')

    def write_dataframe_to_excel(self, dataframe, sheet_name='Dados'):
        writer = pd.ExcelWriter(f'{self.file_name}.xlsx', engine='openpyxl')
        dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
        writer._save()
        self.workbook = openpyxl.load_workbook(f'{self.file_name}.xlsx')

    @staticmethod
    def copy_excel_data(source_file, source_sheet, target_file, target_sheet, start_row, end_row, start_col, end_col):
        df = pd.read_excel(source_file, sheet_name=source_sheet)
        selection = df.loc[start_row:end_row, start_col:end_col]

        new_spreadsheet = pd.DataFrame(selection)

        writer = pd.ExcelWriter(target_file, engine='openpyxl')
        new_spreadsheet.to_excel(writer, sheet_name=target_sheet, index=False)
        writer._save()

    def get_max_row(self, sheet_name):
        worksheet = self.workbook[sheet_name]
        return worksheet.max_row

    def save(self):
        self.workbook.save(f'{self.file_name}.xlsx')

    def close(self):
        self.workbook.close()
