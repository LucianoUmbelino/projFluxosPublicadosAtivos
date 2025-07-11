import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog

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
        icone = resource_path('assets/images/Logo_GPP_Azul-64X64.ico')
        shortcut.IconLocation = str(icone)
        shortcut.save()
        mensagens.append(f"📌 Atalho criado na área de trabalho: {atalho}")
    except Exception as e:
        mensagens.append(f"⚠️ Não foi possível criar o atalho na área de trabalho: {e}")

def validar_pasta_e_planilha():
    if not CAMINHO_INICIAL.exists():
        mensagens.append(f"📁 Pasta '{CAMINHO_INICIAL}' não encontrada. Criando...")
        CAMINHO_INICIAL.mkdir(parents=True, exist_ok=True)
        criar_atalho_na_area_de_trabalho(CAMINHO_INICIAL)

    arquivos_validos = list(CAMINHO_INICIAL.glob("*Fluxos disponíveis para execução no E-Flow (produção)*.xlsx"))
    if not arquivos_validos:
        mensagens.append(f"❌ A planilha 'Fluxos disponíveis para execução no E-Flow (produção)' não foi encontrada na pasta '{CAMINHO_INICIAL}'.")
        exibe_mensagem_terminal("Erro", "\n".join(mensagens))
        sys.exit(1)

    if not CAMINHO_PLANILHA_FINAL.exists():
        mensagens.append(f"❌ A planilha de saída 'Fluxos_Publicados_Ativos.xlsx' não foi encontrada na pasta '{CAMINHO_INICIAL}'.")
        exibe_mensagem_terminal("Erro", "\n".join(mensagens))
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
    global mensagens
    mensagens = []

    validar_pasta_e_planilha()

    meses_validos = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    msgMes = tk.Tk()
    msgMes.withdraw()
    mes = simpledialog.askstring("Seleção do Mês de Processamento", "Informe o mês a ser processado (ex: Junho):")
    msgMes.destroy()

    if mes:
        mes = mes.strip().capitalize()
        if mes not in meses_validos:
            exibe_mensagem_terminal("Parâmetro Inválido", "Mês inválido. Por favor, digite o nome completo de um mês válido.")
            return
    else:
        exibe_mensagem_terminal("Entrada Inválida", "Nenhum mês foi informado.")
        return

    caminho_entrada = selecionar_arquivo()
    if not caminho_entrada:
        exibe_mensagem_terminal("Arquivo não encontrado", "Nenhum arquivo de Fluxos Publicados selecionado.")
        return

    resultado_fluxo = gerar_fluxo_mensal(mes, caminho_entrada)
    if isinstance(resultado_fluxo, dict):
        status = resultado_fluxo.get("status", "").lower()
        if "falha" in status:
            mensagens.append(resultado_fluxo.get("mensagem", "Erro desconhecido."))
            exibe_mensagem_terminal("Falha no Processamento", "\n".join(mensagens))
            return
        elif "sucesso" in status:
            mensagens.append(resultado_fluxo.get("mensagem", ""))

    resultado_fluxo = gerar_graficos_gerais()
    if isinstance(resultado_fluxo, dict):
        status = resultado_fluxo.get("status", "").lower()
        if "falha" in status:
            mensagens.append(resultado_fluxo.get("mensagem", "Erro ao gerar gráficos."))
            exibe_mensagem_terminal("Falha no Processamento", "\n".join(mensagens))
            return
        elif "sucesso" in status:
            mensagens.append(resultado_fluxo.get("mensagem", ""))

    mensagens.append("✅ Sucesso no Processamento: Processamento finalizado com sucesso.")
    exibe_mensagem_terminal("Resumo do Processamento", "\n".join(mensagens))

if __name__ == "__main__":
    main()
