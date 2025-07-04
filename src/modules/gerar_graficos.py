import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from config.settings import CAMINHO_PLANILHA_FINAL

# Mapeamento de nomes completos para siglas
mapa_siglas = {
    "Janeiro": "Jan", "Fevereiro": "Fev", "Março": "Mar", "Abril": "Abr",
    "Maio": "Mai", "Junho": "Jun", "Julho": "Jul", "Agosto": "Ago",
    "Setembro": "Set", "Outubro": "Out", "Novembro": "Nov", "Dezembro": "Dez"
}
ordem_meses = list(mapa_siglas.values())

def gerar_graficos_gerais():
    wb = load_workbook(CAMINHO_PLANILHA_FINAL)

    # Coletar dados das abas mensais
    registros = []
    for nome in wb.sheetnames:
        if nome in mapa_siglas:
            aba = wb[nome]
            linha = 7
            while True:
                patriarca = aba[f"A{linha}"].value
                quantidade = aba[f"C{linha}"].value
                if not patriarca:
                    break
                quantidade_valida = int(quantidade) if quantidade is not None else 0
                registros.append({
                    "Mês": mapa_siglas[nome],
                    "Patriarca": patriarca,
                    "Quantidade": quantidade_valida
                })
                linha += 1

    if not registros:
        raise ValueError("Nenhum dado encontrado nas abas mensais.")

    df = pd.DataFrame(registros)

    # Remover aba antiga de gráficos
    if "Graficos" in wb.sheetnames:
        del wb["Graficos"]
    aba_graf = wb.create_sheet("Graficos")

    # Criar aba de dados para o gráfico
    if "Dados para o Gráfico" in wb.sheetnames:
        del wb["Dados para o Gráfico"]
    aba_dados = wb.create_sheet("Dados para o Gráfico")

    # Total por mês com todos os meses do ano
    totais_mes = df.groupby("Mês")["Quantidade"].sum().reindex(ordem_meses, fill_value=0)

    # Escrever dados na aba de dados
    aba_dados.append(["Mês", "Total"])
    for mes, total in totais_mes.items():
        aba_dados.append([mes, total])

    # Determinar intervalo de meses com dados
    meses_com_dados = totais_mes[totais_mes > 0]
    if not meses_com_dados.empty:
        primeiro_indice = ordem_meses.index(meses_com_dados.index[0])
        ultimo_indice = ordem_meses.index(meses_com_dados.index[-1])
    else:
        primeiro_indice = 0
        ultimo_indice = 11

    # Gráfico de linha
    chart = LineChart()
    chart.title = "Total de Fluxos Publicados por Mês"
    chart.y_axis.title = "Total"
    chart.x_axis.title = "Mês"
    chart.y_axis.majorUnit = 50
    chart.y_axis.minorGridlines = None

    data = Reference(aba_dados, min_col=2, min_row=1 + primeiro_indice + 1, max_row=1 + ultimo_indice + 1)
    cats = Reference(aba_dados, min_col=1, min_row=1 + primeiro_indice + 1, max_row=1 + ultimo_indice + 1)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)

    # Adicionar rótulos de dados
    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True

    aba_graf.add_chart(chart, "B2")

    wb.save(CAMINHO_PLANILHA_FINAL)
    print("✅ Aba 'Graficos' e 'Dados para o Gráfico' criadas com sucesso com gráfico de linha ajustado.")

