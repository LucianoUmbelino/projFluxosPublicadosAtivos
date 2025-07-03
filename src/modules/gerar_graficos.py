
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from datetime import datetime

# === VARIÁVEIS ===
caminho_planilha_final = r'C:\Projects\SEGER-GPP\E-Flow - Fluxos\Fluxos_Publicados_Ativos.xlsx'
nomes_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# === ABRIR PLANILHA DE DESTINO ===
wb = load_workbook(caminho_planilha_final)

# === COLETAR DADOS DE CADA ABA MENSAL ===
registros = []
for nome in wb.sheetnames:
    if nome in nomes_meses:
        aba = wb[nome]
        linha = 7
        while True:
            patriarca = aba[f"A{linha}"].value
            quantidade = aba[f"C{linha}"].value
            if not patriarca:
                break
            quantidade_valida = int(quantidade) if quantidade is not None else 0
            registros.append({
                "Mês": nome,
                "Patriarca": patriarca,
                "Quantidade": quantidade_valida
            })
            linha += 1

if not registros:
    raise ValueError("Nenhum dado encontrado nas abas mensais.")

df = pd.DataFrame(registros)

# === REMOVER ABA DE GRÁFICO ANTIGA SE EXISTIR ===
if "Graficos" in wb.sheetnames:
    del wb["Graficos"]

aba_graf = wb.create_sheet("Graficos")

# === TABELA PIVOT: LINHAS = MÊS, COLUNAS = PATRIARCA ===
df_pivot = df.pivot_table(index="Mês", columns="Patriarca", values="Quantidade", aggfunc="sum", fill_value=0)
df_pivot = df_pivot.sort_index()

# === ESCREVER TABELA NA PLANILHA ===
aba_graf.append(["Mês"] + list(df_pivot.columns))
for mes, row in df_pivot.iterrows():
    aba_graf.append([mes] + list(row.values))

# === CRIAR GRÁFICO DE BARRAS AGRUPADAS ===
chart = BarChart()
chart.type = "col"
chart.title = "Total de Fluxos Publicados por Patriarca por Mês"
chart.x_axis.title = "Mês"
chart.y_axis.title = "Total de Fluxos"

ncols = len(df_pivot.columns)
nrows = len(df_pivot.index)

data = Reference(aba_graf, min_col=2, min_row=1, max_col=1 + ncols, max_row=1 + nrows)
cats = Reference(aba_graf, min_col=1, min_row=2, max_row=1 + nrows)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

aba_graf.add_chart(chart, "B8")

wb.save(caminho_planilha_final)
print("✅ Aba 'Graficos' criada com sucesso com base nas abas mensais.")
