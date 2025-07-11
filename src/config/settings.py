from pathlib import Path
from utils.path_helpers import resource_path

# Caminho base para salvar a planilha final (fora do executável)
CAMINHO_INICIAL = Path(r'C:\SEGER-GPP\E-Flow - Fluxos')

# Caminho completo para a planilha final
CAMINHO_PLANILHA_FINAL = CAMINHO_INICIAL / 'Fluxos_Publicados_Ativos.xlsx'

# Nome da aba modelo usada como referência
ABA_MODELO = "Maio"

# Exemplo de uso de recurso empacotado (se necessário)
# caminho_logo = resource_path('assets/images/Logo_GPP_Azul-64X64.ico')
