# src/utils/logger_helper.py

from tkinter import messagebox, Tk
from pathlib import Path
import datetime

class LoggerFluxo:
    def __init__(self):
        self.mensagens = []

    def info(self, msg: str):
        self.mensagens.append(f"[INFO] {msg}")

    def error(self, msg: str):
        self.mensagens.append(f"[ERRO] {msg}")

    def warn(self, msg: str):
        self.mensagens.append(f"[ALERTA] {msg}")

    def mostrar_mensagem(self, titulo="Resumo do Processamento"):
        texto = "\n".join(self.mensagens)
        root = Tk()
        root.withdraw()
        messagebox.showinfo(titulo, texto)
        root.destroy()

    def salvar_em_arquivo(self, caminho: Path):
        try:
            with open(caminho, "a", encoding="utf-8") as f:
                data = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f.write(f"\n[{data}]\n")
                for msg in self.mensagens:
                    f.write(msg + "\n")
        except Exception as e:
            # Em vez de print, registre o erro internamente (opcionalmente mostre)
            self.error(f"Erro ao salvar log: {e}")
            # self.mostrar_mensagem("Erro ao salvar o arquivo de log")
