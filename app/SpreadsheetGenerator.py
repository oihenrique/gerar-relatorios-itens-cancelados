import pandas as pd
from app.ExcelProcessor import ExcelProcessor
from app.ExcelStyler import ExcelStyler


class SpreadsheetGenerator:
    def __init__(self, stores):
        self.stores = stores

    def create_spreadsheet_from_clipboard(self, file_name):
        excel_processor = ExcelProcessor(file_name)

        clipboard_dataframe = excel_processor.read_clipboard_to_dataframe()
        excel_processor.write_dataframe_to_excel(clipboard_dataframe)
        excel_processor.add_column_to_excel(self.stores)

        excel_styler = ExcelStyler(file_name)

        for index, store in enumerate(self.stores, start=2):
            excel_processor.add_countif_formula(2, store, f'Q{index}')
            excel_styler.apply_number_format('Dados', 17, index, '0')

        excel_processor.save()
        excel_styler.close()
        excel_processor.close()

    @staticmethod
    def transfer_data(source_file, target_file, start_row, end_row, start_col, end_col):
        target_sheet_name = "Contagem"
        ExcelProcessor.copy_excel_data(source_file, 'Dados',
                                       target_file, target_sheet_name,
                                       start_row, end_row, start_col, end_col)

        excel_processor = ExcelProcessor(target_file[:-5], True)

        # Header data
        header_data = {
            'C1': 'Quant. Loja',
            'D1': 'Quant. Sistema',
            'E1': 'Verificação'
        }

        # Add title header and save
        excel_processor.add_title_header(target_sheet_name, header_data)

        tr = 2
        for i in range(1, excel_processor.get_max_row(target_sheet_name)):
            excel_processor.add_if_formula(target_sheet_name, 3, tr, 4, tr,
                                           "OK", "Divergência", f"E{tr}")
            tr += 1

        excel_processor.save()
        excel_processor.close()

        excel_styler = ExcelStyler(target_file[:-5])

        # Apply styles and set column width
        for column in range(1, 6):
            excel_styler.apply_style_to_header(target_sheet_name, column)
            excel_styler.set_column_width(target_sheet_name, column)
            excel_styler.set_column_alignment(target_sheet_name, column)

        excel_styler.set_column_alignment(target_sheet_name, 2, 'left')

        excel_styler.save()
        excel_styler.close()

    def generate_spreadsheets(self, data_file, date):
        ws = pd.read_excel(data_file, sheet_name=0, header=None)

        next_x = 0

        for row_index, store_number in enumerate(self.stores, start=2):
            store_cell = ws.iloc[row_index - 1, -1]
            quantity = (ws[1] == store_cell).sum()

            if store_cell == store_number:
                if store_number < 10:
                    sheet_name = f'{date}_Contagem_R0{store_number}'
                else:
                    sheet_name = f'{date}_Contagem_R{store_number}'

                start_index = next_x
                end_index = next_x + quantity - 1

                self.transfer_data(data_file, f'./excel/{sheet_name}.xlsx', start_index, end_index,
                                   'Código', 'Descrição')

                next_x = end_index + 1
