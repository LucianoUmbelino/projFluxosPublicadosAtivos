from openpyxl.worksheet.worksheet import Worksheet

def ajustar_largura_colunas(aba: Worksheet, colunas=['A', 'B', 'C']):
    for col in colunas:
        max_length = 0
        for cell in aba[col]:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        aba.column_dimensions[col].width = max_length + 2
