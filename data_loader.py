from __future__ import annotations

import os
import re
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
from typing import Tuple

import ast
import json

from config import DATA_DIR, EXCLUDED_LAB_ID


def _ensure_columns(df: pd.DataFrame, columns_with_defaults: dict) -> pd.DataFrame:
    for col, default in columns_with_defaults.items():
        if col not in df.columns:
            df[col] = default
    return df


def _normalize_cnpj(cnpj: str | None) -> str:
    """
    Normaliza CNPJ para formato brasileiro (14 dígitos com zeros à esquerda)
    """
    if pd.isna(cnpj) or cnpj is None:
        return ""
    
    # Remove caracteres não numéricos
    cnpj_clean = re.sub(r'[^\d]', '', str(cnpj))
    
    # Garante 14 dígitos com zeros à esquerda
    return cnpj_clean.zfill(14)


def load_csvs() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    reps_path = os.path.join(DATA_DIR, "representatives.csv")
    labs_path = os.path.join(DATA_DIR, "laboratories.csv")
    gath_path = os.path.join(DATA_DIR, "gatherings.csv")

    df_reps = pd.read_csv(reps_path, low_memory=False)
    df_labs = pd.read_csv(labs_path, low_memory=False)
    df_gatherings = pd.read_csv(gath_path, low_memory=False)

    # Conversão de datas considerando timezone do banco (UTC-3)
    # Estratégia: interpretar timestamps como UTC e converter para America/Sao_Paulo
    def _to_brt(series: pd.Series) -> pd.Series:
        # Garante timezone-aware em UTC e converte para BRT removendo tz
        s = pd.to_datetime(series, errors="coerce", utc=True)
        try:
            s = s.dt.tz_convert("America/Sao_Paulo").dt.tz_localize(None)
        except Exception:
            # Fallback: se não conseguir converter, retorna a série já parseada
            s = s.dt.tz_localize(None)
        return s

    # Tipos e datas (labs)
    for col in ["createdAt", "updatedAt", "exclusionDate"]:
        if col in df_labs.columns:
            df_labs[col] = _to_brt(df_labs[col])

    # Datas (gatherings)
    if "createdAt" in df_gatherings.columns:
        df_gatherings["createdAt"] = _to_brt(df_gatherings["createdAt"])

    # Normalização de IDs (garante chaves consistentes mesmo que o CSV traga formatos variados)
    def _norm_oid(val: object) -> object:
        if pd.isna(val):
            return None
        s = str(val)
        m = re.search(r"[a-fA-F0-9]{24}", s)
        return m.group(0) if m else s

    for col in ["_id"]:
        if col in df_reps.columns:
            df_reps[col] = df_reps[col].apply(_norm_oid)
    for col in ["_id", "_representative"]:
        if col in df_labs.columns:
            df_labs[col] = df_labs[col].apply(_norm_oid)
    for col in ["_laboratory"]:
        if col in df_gatherings.columns:
            df_gatherings[col] = df_gatherings[col].apply(_norm_oid)

    # Normalizar CNPJ dos laboratórios
    if "cnpj" in df_labs.columns:
        df_labs["cnpj"] = df_labs["cnpj"].apply(_normalize_cnpj)

    # Excluir lab de amostras cegas
    if "_id" in df_labs.columns:
        df_labs = df_labs[df_labs["_id"] != EXCLUDED_LAB_ID]

    # Garantir colunas nas coletas
    df_gatherings = _ensure_columns(
        df_gatherings, {"active": True, "test": False, "disabledInReport": False}
    )
    
    # Limpar dados problemáticos
    df_labs = df_labs.replace({"fantasyName": "nan"}, pd.NA)
    df_labs = df_labs.replace({"name_rep": "nan"}, pd.NA)
    
    # Enriquecer com dados de localização (mover para depois da definição da função)
    # df_labs = enrich_labs_with_location(df_labs)
    
    return df_reps, df_labs, df_gatherings


def clean_representative_name(name: str | None) -> str:
    """
    Limpa o nome do representante removendo prefixos técnicos
    """
    if name is None or pd.isna(name):
        return "Sem Representante"
    
    name_str = str(name).strip()
    
    # Lista de prefixos para remover
    prefixes_to_remove = [
        "EXT-", "EXT -", "INT-", "INT -",
        "CAEPTOX - ", "CAEPTOX-", "CAEPTOX – ", "CAEPTOX–",
        "TLMK - ", "TLMK", "TMLK - ",
        "CAEPTOX -", "CAEPTOX –"
    ]
    
    # Remover prefixos
    for prefix in prefixes_to_remove:
        if name_str.upper().startswith(prefix.upper()):
            name_str = name_str[len(prefix):].strip()
            break
    
    # Se ficou vazio após remoção, retornar nome original
    if not name_str:
        return str(name)
    
    return name_str


def categorize_rep(name: str | None) -> str:
    """
    Categorização exatamente igual ao Power BI:
    - Interno: se name (uppercase) começar com qualquer one dos prefixos
    - Externo: caso contrário
    """
    if name is None:
        return "Externo"
    
    upper_name = str(name).upper()
    
    # Prefixos exatamente como no Power BI
    if (upper_name.startswith("INT-") or
        upper_name.startswith("INT -") or
        upper_name.startswith("CAEPTOX - ") or
        upper_name.startswith("CAEPTOX-") or
        upper_name.startswith("CAEPTOX – ") or  # en dash
        upper_name.startswith("CAEPTOX–") or    # em dash
        upper_name.startswith("TLMK - ") or
        upper_name.startswith("TLMK") or
        upper_name.startswith("TMLK - ")):
        return "Interno"
    else:
        return "Externo"


def enrich_labs_with_reps(df_reps: pd.DataFrame, df_labs: pd.DataFrame) -> pd.DataFrame:
    df_reps = df_reps.copy()
    df_labs = df_labs.copy()
    
    # Adicionar categoria e nome limpo
    df_reps["Categoria"] = df_reps["name"].apply(categorize_rep)
    df_reps["name_clean"] = df_reps["name"].apply(clean_representative_name)
    
    # Verificar se as colunas foram criadas
    if "name_clean" not in df_reps.columns:
        print("ERRO: Coluna name_clean não foi criada em df_reps")
        df_reps["name_clean"] = df_reps["name"].fillna("Sem Representante")
    
    reps_lookup = (
        df_reps.set_index("_id")[["name", "Categoria", "name_clean"]].rename(columns={"name": "name_rep", "name_clean": "name_rep_clean"})
    )
    
    # Verificar se as colunas estão no lookup
    if "name_rep_clean" not in reps_lookup.columns:
        print("ERRO: Coluna name_rep_clean não está no reps_lookup")
        reps_lookup["name_rep_clean"] = reps_lookup["name_rep"].fillna("Sem Representante")
    
    df_labs = df_labs.join(reps_lookup, on="_representative", how="left")
    
    # Garantir que a coluna name_rep_clean existe
    if "name_rep_clean" not in df_labs.columns:
        print("ERRO: Coluna name_rep_clean não foi adicionada ao df_labs")
        df_labs["name_rep_clean"] = df_labs["name_rep"].fillna("Sem Representante")
    
    return df_reps, df_labs


def merge_gatherings_with_labs(
    df_gatherings: pd.DataFrame, df_labs: pd.DataFrame
) -> pd.DataFrame:
    # Verificar quais colunas estão disponíveis no df_labs
    available_columns = ["_id", "fantasyName", "active", "approved", "exclusionDate", "name_rep", "name_rep_clean", "Categoria", "cnpj"]
    
    # Adicionar colunas de localização se existirem
    location_columns = ["state_code", "state_name", "city"]
    for col in location_columns:
        if col in df_labs.columns:
            available_columns.append(col)
    
    merged = df_gatherings.merge(
        df_labs[available_columns],
        left_on="_laboratory",
        right_on="_id",
        suffixes=("", "_lab"),
        how="left",
    )
    # Tratar valores nulos/vazios
    merged["name_rep"] = merged["name_rep"].fillna("Sem Representante")
    merged["name_rep_clean"] = merged["name_rep_clean"].fillna("Sem Representante")
    merged["Categoria"] = merged["Categoria"].fillna("Externo")
    
    # Limpar valores "nan" que podem ter passado
    merged["name_rep"] = merged["name_rep"].replace("nan", "Sem Representante")
    merged["name_rep_clean"] = merged["name_rep_clean"].replace("nan", "Sem Representante")
    merged["Categoria"] = merged["Categoria"].replace("nan", "Externo")
    
    return merged


def extract_location_data(address_str: str) -> tuple:
    """
    Extrai estado e cidade do campo address
    """
    if pd.isna(address_str) or address_str == '':
        return None, None
    
    try:
        # Tentar parsear como JSON primeiro
        if isinstance(address_str, str):
            address_data = json.loads(address_str)
        else:
            address_data = address_str
        
        # Extrair estado e cidade
        state_code = None
        city = None
        
        if 'state' in address_data:
            if isinstance(address_data['state'], dict):
                state_code = address_data['state'].get('code')
            else:
                state_code = address_data['state']
        
        if 'city' in address_data:
            city = address_data['city']
        
        return state_code, city
        
    except (json.JSONDecodeError, KeyError, TypeError):
        # Fallback: tentar extrair de string
        try:
            if isinstance(address_str, str) and 'state' in address_str:
                # Buscar padrão de estado
                import re
                state_match = re.search(r"'code':\s*'([A-Z]{2})'", address_str)
                city_match = re.search(r"'city':\s*'([^']+)'", address_str)
                
                state_code = state_match.group(1) if state_match else None
                city = city_match.group(1) if city_match else None
                
                return state_code, city
        except:
            pass
        
        return None, None


def enrich_labs_with_location(df_labs: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona colunas de estado e cidade aos laboratórios
    """
    df_labs = df_labs.copy()
    
    # Extrair estado e cidade do campo address
    location_data = df_labs['address'].apply(extract_location_data)
    df_labs['state_code'] = [data[0] if data else None for data in location_data]
    df_labs['city'] = [data[1] if data else None for data in location_data]
    
    # Mapeamento de códigos de estado para nomes completos
    state_mapping = {
        'AC': 'Acre', 'AL': 'Alagoas', 'AP': 'Amapá', 'AM': 'Amazonas',
        'BA': 'Bahia', 'CE': 'Ceará', 'DF': 'Distrito Federal', 'ES': 'Espírito Santo',
        'GO': 'Goiás', 'MA': 'Maranhão', 'MT': 'Mato Grosso', 'MS': 'Mato Grosso do Sul',
        'MG': 'Minas Gerais', 'PA': 'Pará', 'PB': 'Paraíba', 'PR': 'Paraná',
        'PE': 'Pernambuco', 'PI': 'Piauí', 'RJ': 'Rio de Janeiro', 'RN': 'Rio Grande do Norte',
        'RS': 'Rio Grande do Sul', 'RO': 'Rondônia', 'RR': 'Roraima', 'SC': 'Santa Catarina',
        'SP': 'São Paulo', 'SE': 'Sergipe', 'TO': 'Tocantins'
    }
    
    df_labs['state_name'] = df_labs['state_code'].map(state_mapping)
    
    return df_labs


