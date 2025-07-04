import tkinter as tk
from tkinter import filedialog
from modules.gerar_planilha import gerar_fluxo_mensal
from modules.gerar_graficos import gerar_graficos_gerais
from config.settings import CAMINHO_INICIAL

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)  # Garante que a janela fique em primeiro plano
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha de entrada",
        initialdir=CAMINHO_INICIAL,
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.destroy()
    return caminho

def main():
    mes = input("Informe o mês a ser processado (ex: Junho): ").capitalize()
    caminho_entrada = selecionar_arquivo()

    if not caminho_entrada:
        print("❌ Nenhum arquivo selecionado. Encerrando.")
        return

    gerar_fluxo_mensal(mes, caminho_entrada)
    gerar_graficos_gerais()
    print("✅ Processamento finalizado com sucesso.")

if __name__ == "__main__":
    main()

