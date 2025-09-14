from datetime import datetime
import os

# Tentar importar streamlit para acessar secrets
try:
    import streamlit as st
    HAS_STREAMLIT = True
except ImportError:
    HAS_STREAMLIT = False


def _get_config_value(key: str, default: str) -> str:
    """Obtém valor de configuração do Streamlit secrets ou variável de ambiente"""
    if HAS_STREAMLIT:
        try:
            return st.secrets.get(key, default)
        except Exception:
            pass
    return os.getenv(key, default)


# Diretório base dos CSVs
DATA_DIR: str = _get_config_value("TOX_DATA_DIR", r"D:\OneDrive - Synvia Group\Data Analysis\ToxRepresentatives")

# Laboratório de amostras cegas a excluir do cadastro
EXCLUDED_LAB_ID: str = _get_config_value("EXCLUDED_LAB_ID", "5aa61aeeef23e80010b1224e")

# Ano padrão do dashboard (coletas)
DEFAULT_YEAR: int = int(_get_config_value("DEFAULT_YEAR", "2025"))

# Janela padrão para considerar atividade de coletas (em dias)
DEFAULT_ACTIVITY_WINDOW_DAYS: int = int(_get_config_value("DEFAULT_ACTIVITY_WINDOW_DAYS", "15"))


# Azure AD (login corporativo)
AZURE_TENANT_ID: str = _get_config_value("AZURE_TENANT_ID", "")
AZURE_CLIENT_ID: str = _get_config_value("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET: str = _get_config_value("AZURE_CLIENT_SECRET", "")
AZURE_REDIRECT_URI: str = _get_config_value("AZURE_REDIRECT_URI", "")
AZURE_REQUIRE_LOGIN: bool = _get_config_value("AZURE_REQUIRE_LOGIN", "true").lower() in ["1", "true", "yes"]


def get_current_datetime() -> datetime:
    return datetime.now()


