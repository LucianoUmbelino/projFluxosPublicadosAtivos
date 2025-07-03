from modules.gerar_planilha import gerar_fluxo_mensal
from modules.gerar_graficos import gerar_graficos_gerais

def main():
    mes = "Junho"  # ou pegue automaticamente via datetime
    gerar_fluxo_mensal(mes)
    gerar_graficos_gerais()
    print("✅ Processamento finalizado com sucesso.")

if __name__ == "__main__":
    main()
