import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog

# Módulos do projeto
from src.modules.gerar_planilha import gerar_fluxo_mensal
from src.modules.gerar_graficos import gerar_graficos_gerais
from src.config.settings import CAMINHO_INICIAL, CAMINHO_PLANILHA_FINAL
from src.utils.path_helpers import resource_path
from src.utils.logger_helper import LoggerFluxo

import win32com.client
from win32com.client import Dispatch

logger = LoggerFluxo()


def minimizar_todas_janelas():
    shell = win32com.client.Dispatch("Shell.Application")
    shell.MinimizeAll()


def criar_atalho_na_area_de_trabalho(destino: Path):
    try:
        desktop = Path(os.path.join(os.environ["USERPROFILE"], "Desktop"))
        atalho = desktop / "E-Flow Fluxos.lnk"

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(str(atalho))
        shortcut.Targetpath = str(destino)
        shortcut.WorkingDirectory = str(destino)

        icone = resource_path('assets/images/FolderCheck.ico')
        shortcut.IconLocation = str(icone)

        shortcut.save()
        logger.info(f"Atalho criado na área de trabalho: {atalho}")
    except Exception as e:
        logger.error(f"Não foi possível criar o atalho na área de trabalho: {e}")


def validar_pasta_e_planilha():
    try:
        if not CAMINHO_INICIAL.exists():
            logger.warn(f"Pasta '{CAMINHO_INICIAL}' não encontrada. Criando...")
            CAMINHO_INICIAL.mkdir(parents=True, exist_ok=True)
            criar_atalho_na_area_de_trabalho(CAMINHO_INICIAL)

        arquivos_validos = list(CAMINHO_INICIAL.glob("*Fluxos disponíveis para execução no E-Flow (produção)*.xlsx"))
        if not arquivos_validos:
            logger.error(f"A planilha 'Fluxos disponíveis para execução no E-Flow (produção)' não foi encontrada na pasta '{CAMINHO_INICIAL}'.")
            logger.mostrar_mensagem("Erro na Validação")
            # logger.salvar_em_arquivo(Path("fluxo_execucao.log"))
            sys.exit(1)

        if not CAMINHO_PLANILHA_FINAL.exists():
            logger.error(f"A planilha de saída 'Fluxos_Publicados_Ativos.xlsx' não foi encontrada na pasta '{CAMINHO_INICIAL}'.")
            logger.mostrar_mensagem("Erro na Validação")
            # logger.salvar_em_arquivo(Path("fluxo_execucao.log"))
            sys.exit(1)

    except Exception as e:
        logger.error(f"Erro ao validar pasta e planilhas: {e}")
        logger.mostrar_mensagem("Erro na Validação")
        # logger.salvar_em_arquivo(Path("fluxo_execucao.log"))
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
    minimizar_todas_janelas()
    validar_pasta_e_planilha()

    meses_validos = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    msg_mes = tk.Tk()
    msg_mes.withdraw()
    mes = simpledialog.askstring("Seleção do Mês de Processamento", "Informe o mês a ser processado (ex: Junho):")
    msg_mes.destroy()

    if mes:
        mes = mes.strip().capitalize()

    if mes not in meses_validos:
        logger.error("Mês inválido. Por favor, digite o nome completo de um mês válido.")
        logger.mostrar_mensagem("Parâmetro Inválido")
        return

    caminho_entrada = selecionar_arquivo()

    if not caminho_entrada:
        logger.warn("Nenhum arquivo de Fluxos Publicados selecionado.")
        logger.mostrar_mensagem("Arquivo não encontrado")
        return

    resultado_fluxo = gerar_fluxo_mensal(mes, caminho_entrada)
    if isinstance(resultado_fluxo, dict):
        status = resultado_fluxo.get("status", "").strip()
        mensagem = resultado_fluxo.get("mensagem", "").strip()
        logger.info(status)
        logger.info(mensagem)

        if "falha" in status.lower():
            logger.mostrar_mensagem("Falha no Processamento")
            return

    resultado_graficos = gerar_graficos_gerais()
    if isinstance(resultado_graficos, list):
        houve_falha = False
        for resultado in resultado_graficos:
            if isinstance(resultado, dict):
                status = resultado.get("status", "").lower()
                mensagem = resultado.get("mensagem", "")
                logger.info(mensagem)
                if "falha" in status:
                    houve_falha = True

        if houve_falha:
            logger.mostrar_mensagem("Falha no Processamento")
            return

    logger.info("Processamento finalizado com sucesso.")
    logger.mostrar_mensagem("Resumo do Processamento")
    # logger.salvar_em_arquivo(Path("fluxo_execucao.log"))


if __name__ == "__main__":
    main()
