from datetime import datetime
import os


# Diretório base dos CSVs
DATA_DIR: str = os.getenv("TOX_DATA_DIR", r"D:\Progamação\ToxRepresentatives")

# Laboratório de amostras cegas a excluir do cadastro
EXCLUDED_LAB_ID: str = os.getenv("EXCLUDED_LAB_ID", "5aa61aeeef23e80010b1224e")

# Ano padrão do dashboard (coletas)
DEFAULT_YEAR: int = int(os.getenv("DEFAULT_YEAR", "2025"))

# Janela padrão para considerar atividade de coletas (em dias)
DEFAULT_ACTIVITY_WINDOW_DAYS: int = int(os.getenv("DEFAULT_ACTIVITY_WINDOW_DAYS", "15"))


def get_current_datetime() -> datetime:
    return datetime.now()


