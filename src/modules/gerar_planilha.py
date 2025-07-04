import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill, Alignment
from datetime import datetime
from config.settings import CAMINHO_PLANILHA_FINAL, ABA_MODELO
from utils.excel_helpers import ajustar_largura_colunas


def padronizar_coluna_nome(df):
    if "Nome" in df.columns:
        return df
    elif "Nomes" in df.columns:
        return df.rename(columns={"Nomes": "Nome"})
    else:
        raise ValueError("A planilha deve conter a coluna 'Nome' ou 'Nomes'.")


def gerar_fluxo_mensal(mes: str, caminho_fluxos: str):
    # === LEITURA DA PLANILHA DE FLUXOS ===
    df = pd.read_excel(caminho_fluxos, engine="openpyxl")
    df = padronizar_coluna_nome(df)

    colunas_necessarias = ["Patriarca", "Órgão", "Nome"]
    if not all(col in df.columns for col in colunas_necessarias):
        raise ValueError("A planilha deve conter as colunas 'Patriarca', 'Órgão' e 'Nome'.")

    df = df.dropna(subset=colunas_necessarias)

    df_resumo = df.groupby(["Patriarca", "Órgão"])["Nome"].count().reset_index()
    df_resumo.columns = ["Patriarca", "Setor", "Quantidade"]
    df_resumo = df_resumo.sort_values(by=["Patriarca", "Setor"])

    # === ABRIR A PLANILHA DE DESTINO ===
    wb = load_workbook(CAMINHO_PLANILHA_FINAL)

    if ABA_MODELO not in wb.sheetnames:
        raise ValueError(f"A planilha deve conter uma aba chamada '{ABA_MODELO}' como modelo.")

    if mes in wb.sheetnames:
        del wb[mes]

    aba_modelo: Worksheet = wb[ABA_MODELO]
    nova_aba: Worksheet = wb.copy_worksheet(aba_modelo)
    nova_aba.title = mes

    # Inserir nova aba antes da aba "Graficos", se existir
    if "Graficos" in wb.sheetnames:
        idx_graficos = wb.sheetnames.index("Graficos")
        wb._sheets.remove(nova_aba)
        wb._sheets.insert(idx_graficos, nova_aba)

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

    # Estilos base
    fill_branco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_azul = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    alinhamento_esquerda = Alignment(horizontal="left")

    # Preenche dados a partir da linha 7
    linha = 7
    for i, (_, row) in enumerate(df_resumo.iterrows()):
        cor = fill_branco if i % 2 == 0 else fill_azul
        nova_aba[f"A{linha}"] = row["Patriarca"]
        nova_aba[f"B{linha}"] = row["Setor"]
        nova_aba[f"C{linha}"] = row["Quantidade"]

        for col in ["A", "B", "C"]:
            cell = nova_aba[f"{col}{linha}"]
            cell.fill = cor
            if col == "B":
                cell.alignment = alinhamento_esquerda

        linha += 1

    # Soma total na coluna E
    nova_aba["E7"] = f"=SUM(C7:C{linha-1})"

    # Ajusta largura das colunas
    ajustar_largura_colunas(nova_aba, colunas=["A", "B", "C"])

    # Salva
    wb.save(CAMINHO_PLANILHA_FINAL)
    print(f"✅ Aba '{mes}' criada/atualizada com sucesso no arquivo:\n{CAMINHO_PLANILHA_FINAL}")

