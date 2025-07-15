import os
from pathlib import Path

desktop_path = Path(os.path.join(os.environ["USERPROFILE"], "Desktop"))

# Caminho baseado na Área de Trabalho do usuário logado
CAMINHO_INICIAL = desktop_path / 'E-Flow - Fluxos'
CAMINHO_PLANILHA_FINAL = CAMINHO_INICIAL / 'Fluxos_Publicados_Ativos.xlsx'

# Nome da aba modelo usada como referência
ABA_MODELO = "Maio"

