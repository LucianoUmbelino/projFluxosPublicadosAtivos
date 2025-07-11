import sys
from pathlib import Path

def resource_path(relative_path: str) -> Path:
    """Retorna o caminho absoluto para recursos, compatível com PyInstaller."""
    try:
        # Quando empacotado com PyInstaller
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        # Quando rodando como script normal
        base_path = Path(__file__).resolve().parent.parent
    return base_path / relative_path
