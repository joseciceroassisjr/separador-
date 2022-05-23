import xlwings as xw
from pathlib import Path

#Ler o arquivo e criar a pasta de saída
EXCEL_FILE = "RTIP SIMPLIFICADO.xlsx"
OUTPUT_DIR = Path(__file__).parent / 'nome_do_arquivo_excel'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

#Fazer as cópias das worksheets (planilha de trabalho)
try: 
    excel_app = xw.App(visible=False)
    wb = excel_app.books.open(EXCEL_FILE)
    for sheet in wb.sheets:
        sheet.api.Copy()
        wb_new = xw.books.active
        wb_new.save(OUTPUT_DIR / f'{sheet.name}')
        wb_new.close()

finally:
    excel_app.quit()
