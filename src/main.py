import os
import sys
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from modules.gerar_planilha import gerar_fluxo_mensal
from modules.gerar_graficos import gerar_graficos_gerais
from config.settings import CAMINHO_INICIAL, CAMINHO_PLANILHA_FINAL

def criar_atalho_na_area_de_trabalho(destino: Path):
    try:
        import pythoncom
        from win32com.client import Dispatch

        desktop = Path(os.path.join(os.environ["USERPROFILE"], "Desktop"))
        atalho = desktop / "E-Flow Fluxos.lnk"

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(str(atalho))
        shortcut.Targetpath = str(destino)
        shortcut.WorkingDirectory = str(destino)
        shortcut.IconLocation = "explorer.exe, 0"
        shortcut.save()
        print(f"📌 Atalho criado na área de trabalho: {atalho}")
    except Exception as e:
        print(f"⚠️ Não foi possível criar o atalho na área de trabalho: {e}")

def validar_pasta_e_planilha():
    if not CAMINHO_INICIAL.exists():
        print(f"📁 Pasta '{CAMINHO_INICIAL}' não encontrada. Criando...")
        CAMINHO_INICIAL.mkdir(parents=True, exist_ok=True)
        criar_atalho_na_area_de_trabalho(CAMINHO_INICIAL)

    arquivos_validos = list(CAMINHO_INICIAL.glob("*Fluxos disponíveis para execução no E-Flow (produção)*.xlsx"))
    if not arquivos_validos:
        print(f"❌ A planilha 'Fluxos disponíveis para execução no E-Flow (produção)' não foi encontrada na pasta '{CAMINHO_INICIAL}'.")
        sys.exit(1)

    if not CAMINHO_PLANILHA_FINAL.exists():
        print(f"❌ A planilha de saida 'Fluxos_Publicados_Ativos.xlsx' não foi encontrada na pasta '{CAMINHO_INICIAL}'.")
        sys.exit(1)

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha de entrada",
        initialdir=CAMINHO_INICIAL,
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.destroy()
    return caminho

def main():
    validar_pasta_e_planilha()

    meses_validos = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    mes = input("Informe o mês a ser processado (ex: Junho): ").strip().capitalize()
    if mes not in meses_validos:
        print("❌ Mês inválido. Por favor, digite o nome completo de um mês válido.")
        return

    caminho_entrada = selecionar_arquivo()

    if not caminho_entrada:
        print("❌ Nenhum arquivo selecionado. Encerrando.")
        return

    gerar_fluxo_mensal(mes, caminho_entrada)
    gerar_graficos_gerais()
    print("✅ Processamento finalizado com sucesso.")

if __name__ == "__main__":
    main()

