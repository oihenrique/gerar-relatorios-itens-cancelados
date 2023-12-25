from app.excel.SpreadsheetGenerator import SpreadsheetGenerator

stores = [2, 6, 8, 10, 11, 12, 15, 16, 20, 21, 22,
          23, 24, 25, 29, 30, 35, 37, 47, 48, 49,
          50, 51, 55, 57, 61, 63, 64, 65, 66, 67, 68, 69,
          70, 80, 81, 83]

generator = SpreadsheetGenerator(stores)

# generator.create_spreadsheet_from_clipboard('./data/24-12-23_Cancelados')

# generator.generate_spreadsheets('./data/24-12-23_Cancelados.xlsx', '24-12-23')

# send_email()
