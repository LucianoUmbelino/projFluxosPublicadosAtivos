import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from datetime import datetime

# === VARIÁVEIS ===
mes = "Junho"

caminho_fluxos = r'C:\Projects\SEGER-GPP\E-Flow - Fluxos\Fluxos disponíveis para execução no E-Flow (produção) (2025-06-03).xlsx'
caminho_planilha_final = r'C:\Projects\SEGER-GPP\E-Flow - Fluxos\Fluxos_Publicados_Ativos.xlsx'

# === LEITURA DA PLANILHA DE FLUXOS ===
df = pd.read_excel(caminho_fluxos)

# Verifica se as colunas necessárias existem
if "Patriarca" not in df.columns or "Órgão" not in df.columns or "Nomes" not in df.columns:
    raise ValueError("A planilha deve conter as colunas 'Patriarca', 'Órgão' e 'Nomes'.")

# Remove linhas com valores nulos nas colunas necessárias
df = df.dropna(subset=["Patriarca", "Órgão", "Nomes"])

# Agrupa e conta todas as ocorrências por setor
df_resumo = df.groupby(["Patriarca", "Órgão"])["Nomes"].count().reset_index()
df_resumo.columns = ["Patriarca", "Setor", "Quantidade"]
df_resumo = df_resumo.sort_values(by=["Patriarca", "Setor"])

# === ABRIR A PLANILHA DE DESTINO ===
wb = load_workbook(caminho_planilha_final)

# Verifica se aba "Abril" existe para usar como modelo
if "Abril" not in wb.sheetnames:
    raise ValueError("A planilha deve conter uma aba chamada 'Abril' como modelo.")

# Verifica se a aba do mês existe
if mes in wb.sheetnames:
    # Exclui a aba existente
    del wb[mes]

# Copia a aba modelo
aba_modelo: Worksheet = wb["Maio"]
nova_aba: Worksheet = wb.copy_worksheet(aba_modelo)
nova_aba.title = mes

# Atualiza cabeçalhos
nova_aba["A4"] = f"Contagem de Fluxos Publicados e Ativos por Setor - Mês {mes}"
nova_aba["A5"] = f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
nova_aba["A6"] = "Patriarca"
nova_aba["B6"] = "Setor"
nova_aba["C6"] = "Quantidade"
nova_aba["E6"] = "Total"

# Remove dados anteriores
for row in nova_aba.iter_rows(min_row=7, max_row=nova_aba.max_row, min_col=1, max_col=5):
    for cell in row:
        cell.value = None

# Preenche dados a partir da linha 7
linha = 7
for _, row in df_resumo.iterrows():
    nova_aba[f"A{linha}"] = row["Patriarca"]
    nova_aba[f"B{linha}"] = row["Setor"]
    nova_aba[f"C{linha}"] = row["Quantidade"]
    linha += 1

# Soma total na coluna E
nova_aba["E7"] = f"=SUM(C7:C{linha-1})"

# Ajusta a largura das colunas A, B e C de acordo com o maior conteúdo
for col in ['A', 'B', 'C']:
    max_length = 0
    for cell in nova_aba[col]:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    nova_aba.column_dimensions[col].width = max_length + 2

# === SALVA O ARQUIVO ===
wb.save(caminho_planilha_final)
print(f"Aba '{mes}' criada/atualizada com sucesso no arquivo:\\n{caminho_planilha_final}")
