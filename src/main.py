import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog

# Importações dos módulos do projeto
from src.modules.gerar_planilha import gerar_fluxo_mensal
from src.modules.gerar_graficos import gerar_graficos_gerais
from src.config.settings import CAMINHO_INICIAL, CAMINHO_PLANILHA_FINAL
from src.utils.path_helpers import resource_path

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

        # Define o ícone do atalho usando o caminho empacotado
        icone = resource_path('assets/images/Logo_GPP_Azul-64X64.ico')
        shortcut.IconLocation = str(icone)

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
        print(f"❌ A planilha de saída 'Fluxos_Publicados_Ativos.xlsx' não foi encontrada na pasta '{CAMINHO_INICIAL}'.")
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

def exibe_mensagem_terminal(titulo, msg: str):
    janelaMsgs = tk.Tk()
    janelaMsgs.withdraw()

    messagebox.showinfo(titulo, msg)

    janelaMsgs.destroy()

def main():
    validar_pasta_e_planilha()

    meses_validos = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    msgMes = tk.Tk()
    msgMes.withdraw()  # Esconde a janela principal

    mes = simpledialog.askstring("Seleção do Mês de Processamento", "Informe o mês a ser processado (ex: Junho):")
    msgMes.destroy()

    if mes:
        mes = mes.strip().capitalize()

    if mes not in meses_validos:
        exibe_mensagem_terminal("Parâmetro Inválido","Mês inválido. Por favor, digite o nome completo de um mês válido.")
        return

    caminho_entrada = selecionar_arquivo()

    if not caminho_entrada:
        exibe_mensagem_terminal("Arquivo não encontrado","Nenhum arquivo de Fluxos Publicados selecionado.")
        return

    gerar_fluxo_mensal(mes, caminho_entrada)
    gerar_graficos_gerais()
    exibe_mensagem_terminal("Sucesso no Processamento","Processamento finalizado com sucesso.")

if __name__ == "__main__":
    main()

