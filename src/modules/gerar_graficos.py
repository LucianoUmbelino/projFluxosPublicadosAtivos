import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from io import BytesIO
from datetime import datetime
from config.settings import CAMINHO_PLANILHA_FINAL

# Mapeamento de nomes completos para siglas
mapa_siglas = {
    "Janeiro": "Jan", "Fevereiro": "Fev", "Março": "Mar", "Abril": "Abr",
    "Maio": "Mai", "Junho": "Jun", "Julho": "Jul", "Agosto": "Ago",
    "Setembro": "Set", "Outubro": "Out", "Novembro": "Nov", "Dezembro": "Dez"
}
meses = list(mapa_siglas.keys())
ordem_meses = list(mapa_siglas.values())

def gerar_grafico_linha_evolucao_mensal_de_Fluxos_publicados(caminho_arquivo: str):
    # Inicializa a lista final
    fluxos_publicados = []

    # Lê a estrutura da planilha
    excel_file = pd.ExcelFile(caminho_arquivo)
    abas = excel_file.sheet_names

    # Para cada mês, extrai o somatório da coluna 'Quantidade'
    for mes in meses:
        if mes in abas:
            df = pd.read_excel(caminho_arquivo, sheet_name=mes, skiprows=6, header=None)
            # A coluna de índice 2 contém os valores da Quantidade
            quantidade = pd.to_numeric(df.iloc[:, 2], errors="coerce").dropna()
            total = int(quantidade.sum())
        else:
            total = 0
        fluxos_publicados.append(total)

    # Gera o gráfico de linha
    fig, ax = plt.subplots(figsize=(10, 5))

    plt.plot(ordem_meses, fluxos_publicados, marker="o", linestyle="-", color="blue")

    plt.xlabel("Meses")
    plt.ylabel("Total de Fluxos Publicados")
    plt.title(f"Evolução Mensal de Fluxos Publicados em {datetime.now().year}")
    plt.grid(True)
    plt.yticks(range(0, max(fluxos_publicados) + 50, 50))

    plt.tight_layout()

    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    plt.close(fig)
    buffer.seek(0)

    wb = load_workbook(caminho_arquivo)
    if "Graficos" in wb.sheetnames:
        del wb["Graficos"]
    aba_graficos = wb.create_sheet("Graficos")
    img = XLImage(buffer)
    aba_graficos.add_image(img, "A1")
    wb.save(caminho_arquivo)

def gerar_graficos_gerais():
    gerar_grafico_linha_evolucao_mensal_de_Fluxos_publicados(CAMINHO_PLANILHA_FINAL)
    print("✅ Aba 'Graficos' e gráfico de linha criados com sucesso.")

