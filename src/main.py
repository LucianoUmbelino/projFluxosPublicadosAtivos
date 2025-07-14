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

import pythoncom
from win32com.client import Dispatch

mensagens = []

def criar_atalho_na_area_de_trabalho(destino: Path):
    try:
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
        mensagens.append(f"Atalho criado na área de trabalho: {atalho}" + "\n")
    except Exception as e:
        mensagens.append(f"Não foi possível criar o atalho na área de trabalho: {e}" + "\n")

def validar_pasta_e_planilha():
    if not CAMINHO_INICIAL.exists():
        mensagens.append(f"Pasta '{CAMINHO_INICIAL}' não encontrada. Criando..." + "\n")
        CAMINHO_INICIAL.mkdir(parents=True, exist_ok=True)
        criar_atalho_na_area_de_trabalho(CAMINHO_INICIAL)

    arquivos_validos = list(CAMINHO_INICIAL.glob("*Fluxos disponíveis para execução no E-Flow (produção)*.xlsx"))
    if not arquivos_validos:
        mensagens.append(f"A planilha 'Fluxos disponíveis para execução no E-Flow (produção)' não foi encontrada na pasta '{CAMINHO_INICIAL}'." + "\n")
        sys.exit(1)

    if not CAMINHO_PLANILHA_FINAL.exists():
        mensagens.append(f"A planilha de saída 'Fluxos_Publicados_Ativos.xlsx' não foi encontrada na pasta '{CAMINHO_INICIAL}'." + "\n")
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
    janela_msgs = tk.Tk()
    janela_msgs.withdraw()

    messagebox.showinfo(titulo, msg)

    janela_msgs.destroy()

def main():
    validar_pasta_e_planilha()

    meses_validos = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    msg_mes = tk.Tk()
    msg_mes.withdraw()  # Esconde a janela principal

    mes = simpledialog.askstring("Seleção do Mês de Processamento", "Informe o mês a ser processado (ex: Junho):")
    msg_mes.destroy()

    if mes:
        mes = mes.strip().capitalize()

    if mes not in meses_validos:
        exibe_mensagem_terminal("Parâmetro Inválido","Mês inválido. Por favor, digite o nome completo de um mês válido.")
        return

    caminho_entrada = selecionar_arquivo()

    if not caminho_entrada:
        exibe_mensagem_terminal("Arquivo não encontrado","Nenhum arquivo de Fluxos Publicados selecionado.")
        return

    resultado_fluxo = gerar_fluxo_mensal(mes, caminho_entrada)
    if isinstance(resultado_fluxo, dict):
        status = resultado_fluxo.get("status", "").strip()
        mensagem = resultado_fluxo.get("mensagem", "").strip()
        mensagens.append(f"{status}\n{mensagem}\n")

        if "falha" in status.lower():
            exibe_mensagem_terminal("Falha no Processamento", "\n".join(mensagens))
            return

    resultado_graficos = gerar_graficos_gerais()
    if isinstance(resultado_graficos, list):
        houve_falha = False
        for resultado in resultado_graficos:
            if isinstance(resultado, dict):
                status = resultado.get("status", "").lower()
                mensagem = resultado.get("mensagem", "") + "\n"
                mensagens.append(mensagem)
                if "falha" in status:
                    houve_falha = True

        if houve_falha:
            exibe_mensagem_terminal("Falha no Processamento", "\n".join(mensagens))
            return

    mensagens.append("Sucesso no Processamento" + "\n" + "Processamento finalizado com sucesso.\n")
    exibe_mensagem_terminal("Resumo do Processamento", "\n".join(mensagens))


if __name__ == "__main__":
    main()