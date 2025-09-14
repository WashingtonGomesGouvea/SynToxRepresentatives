from __future__ import annotations

import pandas as pd
from datetime import datetime, timedelta
from typing import Tuple, Dict

from config import DEFAULT_ACTIVITY_WINDOW_DAYS


def filter_active_gatherings(df: pd.DataFrame, exclude_test: bool, exclude_disabled: bool) -> pd.DataFrame:
    # Usar active (da coleta) - todas as coletas têm active=True
    out = df[df["active"] == True]
    
    if exclude_test and "test" in out.columns:
        out = out[out["test"] == False]
    if exclude_disabled and "disabledInReport" in out.columns:
        out = out[out["disabledInReport"] == False]
    return out


def compute_credenciamento(df_labs: pd.DataFrame, current_date: datetime) -> pd.DataFrame:
    """
    Interpretação revisada de credenciamento, alinhada à operação:
    - Credenciado: laboratorio ativo e sem exclusão efetiva (exclusionDate nulo OU no futuro).
    - Descredenciado: caso contrário (inclui exclusão no passado OU active=false).
    Nota: campo 'approved' não determina cancelamento; é estado operacional.
    """
    labs = df_labs.copy()
    # Credenciado: ativo no sistema. Outros casos (active=False) considerados descredenciados/cancelados.
    labs["is_credenciado"] = labs.get("active", True) == True
    return labs


def compute_coleta_status(
    df_labs_cred: pd.DataFrame,
    df_gatherings_active: pd.DataFrame,
    current_date: datetime,
    activity_window_days: int = DEFAULT_ACTIVITY_WINDOW_DAYS,
) -> Tuple[pd.DataFrame, int, int, pd.DataFrame]:
    labs = df_labs_cred.copy()
    
    # Inicializar colunas para todos os labs
    labs["ativo_coleta"] = False
    labs["days_since_last"] = 999
    labs["ultima_coleta"] = pd.NaT
    labs["ultima_coleta_str"] = "Sem coletas"
    # Garantir coluna de exibição mesmo quando não houver coletas
    labs["days_since_last_display"] = labs["days_since_last"].apply(
        lambda x: str(x) if x != 999 else "-"
    )
    
    if df_gatherings_active.empty:
        last_collection = pd.DataFrame(columns=["_laboratory", "createdAt"])  # vazio
    else:
        # Garantir que createdAt seja datetime antes de calcular o máximo
        df_gatherings_clean = df_gatherings_active.copy()
        if "createdAt" in df_gatherings_clean.columns:
            df_gatherings_clean["createdAt"] = pd.to_datetime(df_gatherings_clean["createdAt"], errors="coerce")
            # Remover linhas onde a conversão falhou (valores inválidos)
            df_gatherings_clean = df_gatherings_clean.dropna(subset=["createdAt"])
        
        # Calcular última coleta por laboratório
        last_collection = df_gatherings_clean.groupby("_laboratory")["createdAt"].max().reset_index()
        
        # Merge com labs (usar suffixes para evitar conflito)
        labs = labs.merge(
            last_collection[["_laboratory", "createdAt"]],
            left_on="_id",
            right_on="_laboratory",
            how="left",
            suffixes=("_lab", "_coleta")
        )
        
        # Calcular dias desde a última coleta (robusto a timezone/horário futuro no mesmo dia)
        def _days_diff_safe(dt: pd.Timestamp) -> int:
            if pd.isna(dt):
                return 999
            diff_days = (current_date - dt).days
            # Se negativo (ex.: hora da coleta "no futuro" por fuso), normaliza para 0
            return max(diff_days, 0)

        labs["days_since_last"] = labs["createdAt_coleta"].apply(_days_diff_safe)
        
        # Atualizar última coleta e status
        labs["ultima_coleta"] = labs["createdAt_coleta"]
        labs["ultima_coleta_str"] = labs["ultima_coleta"].apply(
            lambda x: x.strftime("%d/%m/%Y %H:%M") if pd.notna(x) else "Sem coletas"
        )
        
        # Criar coluna para exibição dos dias (mostrar "-" para labs sem coletas)
        labs["days_since_last_display"] = labs["days_since_last"].apply(
            lambda x: str(x) if x != 999 else "-"
        )
        
        # Definir status de atividade baseado nos dias
        labs["ativo_coleta"] = labs["days_since_last"] <= activity_window_days

    ativos = int(labs["ativo_coleta"].sum())
    inativos = int(len(labs) - ativos)
    return labs, ativos, inativos, last_collection


def aggregate_volumes(df_gatherings_active: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    g = df_gatherings_active.copy()
    
    # Garantir que createdAt seja datetime antes de usar .dt accessor
    if "createdAt" in g.columns:
        g["createdAt"] = pd.to_datetime(g["createdAt"], errors="coerce")
        # Remover linhas onde a conversão falhou (valores inválidos)
        g = g.dropna(subset=["createdAt"])
    
    g["month"] = g["createdAt"].dt.to_period("M")
    g["week"] = g["createdAt"].dt.to_period("W")
    
    # Agregação mensal por categoria
    monthly = g.groupby(["month", "Categoria"]).size().reset_index(name="Volume")
    monthly["month"] = monthly["month"].astype(str)

    # Agregação semanal por categoria
    weekly = g.groupby(["week", "Categoria"]).size().reset_index(name="Volume")
    weekly["week"] = weekly["week"].astype(str)
    
    return weekly, monthly


def compute_kpis(monthly: pd.DataFrame) -> Dict[str, int]:
    """
    KPIs mensais consolidados. Se houver coluna 'Categoria', consolida por mês (soma das categorias)
    para evitar duplicidade ao calcular mínimo, máximo e média.
    """
    if monthly.empty or 'Volume' not in monthly.columns:
        return {"total": 0, "max": 0, "min": 0, "avg": 0}

    df = monthly.copy()
    if 'Categoria' in df.columns:
        df_total = df.groupby('month', as_index=False)['Volume'].sum()
        series = df_total['Volume']
    else:
        series = df['Volume']

    if series.empty:
        return {"total": 0, "max": 0, "min": 0, "avg": 0}

    total = int(series.fillna(0).sum())
    max_m = int(series.max()) if pd.notna(series.max()) else 0
    min_m = int(series.min()) if pd.notna(series.min()) else 0
    avg_val = series.mean()
    try:
        avg_m = int(avg_val) if pd.notna(avg_val) else 0
    except Exception:
        avg_m = 0

    return {"total": total, "max": max_m, "min": min_m, "avg": avg_m}


def build_rankings(
    df_gatherings_active: pd.DataFrame,
    df_labs_enriched: pd.DataFrame,
    last_collection: pd.DataFrame,
    current_date: datetime,
    activity_window_days: int,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # Filtrar apenas coletas de labs que existem
    labs_ids_validos = df_labs_enriched['_id'].unique()
    df_gatherings_valid = df_gatherings_active[df_gatherings_active['_laboratory'].isin(labs_ids_validos)]
    
    # Usar name_rep_clean se disponível, senão name_rep
    rep_name_col = 'name_rep_clean' if 'name_rep_clean' in df_gatherings_valid.columns else 'name_rep'
    
    ranking_reps = (
        df_gatherings_valid[df_gatherings_valid[rep_name_col].notna()]
        .groupby([rep_name_col, "Categoria"]).size().reset_index(name="Volume")
        .sort_values("Volume", ascending=False)
    )
    
    # Renomear coluna para manter consistência
    ranking_reps = ranking_reps.rename(columns={rep_name_col: "name_rep"})
    
    ranking_labs = (
        df_gatherings_valid.groupby("_laboratory").size().reset_index(name="Volume")
        .sort_values("Volume", ascending=False)
    )
    
    # Incluir name_rep_clean se disponível
    lab_columns = ["_id", "fantasyName", "cnpj", "is_credenciado", "name_rep", "Categoria"]
    if 'name_rep_clean' in df_labs_enriched.columns:
        lab_columns.append("name_rep_clean")
    
    ranking_labs = ranking_labs.merge(
        df_labs_enriched[lab_columns],
        left_on="_laboratory",
        right_on="_id",
        how="left",
    ).drop(columns=["_id"])

    if not last_collection.empty:
        ranking_labs = ranking_labs.merge(
            last_collection[["_laboratory", "createdAt"]],
            on="_laboratory",
            how="left",
        )
    else:
        ranking_labs["createdAt"] = pd.NaT

    ranking_labs = ranking_labs.rename(columns={"createdAt": "ultima_coleta"})
    # Formatar data da última coleta para exibição
    ranking_labs["ultima_coleta_str"] = ranking_labs["ultima_coleta"].apply(
        lambda x: x.strftime("%d/%m/%Y %H:%M") if pd.notna(x) else "Sem coletas"
    )
    ranking_labs["status_coleta"] = ranking_labs["ultima_coleta"].apply(
        lambda x: "Ativo"
        if pd.notna(x) and (current_date - x).days <= activity_window_days
        else "Inativo" if pd.notna(x) else "Sem Coletas"
    )

    return ranking_reps, ranking_labs


def compute_representative_metrics(
    df_gatherings_active: pd.DataFrame,
    df_labs_status: pd.DataFrame,
    current_date: datetime,
    activity_window_days: int = DEFAULT_ACTIVITY_WINDOW_DAYS,
) -> pd.DataFrame:
    """
    Calcula métricas de performance por representante para gestão comercial
    """
    # Usar name_rep_clean se disponível, senão name_rep
    rep_name_col = 'name_rep_clean' if 'name_rep_clean' in df_gatherings_active.columns else 'name_rep'
    
    # Agrupar coletas por representante
    rep_metrics = df_gatherings_active.groupby([rep_name_col, 'Categoria']).agg({
        '_laboratory': 'nunique',  # Labs únicos que coletaram
        'createdAt': 'count'       # Total de coletas
    }).reset_index()
    rep_metrics.columns = [rep_name_col, 'Categoria', 'labs_ativos', 'total_coletas']
    
    # Calcular labs credenciados por representante
    lab_rep_col = 'name_rep_clean' if 'name_rep_clean' in df_labs_status.columns else 'name_rep'
    labs_por_rep = df_labs_status.groupby([lab_rep_col, 'Categoria']).agg({
        '_id': 'count',  # Total de labs credenciados
        'is_credenciado': 'sum'  # Labs credenciados
    }).reset_index()
    labs_por_rep.columns = [lab_rep_col, 'Categoria', 'total_labs', 'labs_credenciados']
    
    # Merge das métricas
    rep_metrics = rep_metrics.merge(labs_por_rep, on=[rep_name_col, 'Categoria'], how='outer')
    rep_metrics = rep_metrics.fillna(0)
    
    # Calcular labs inativos (credenciados mas não coletando)
    rep_metrics['labs_inativos'] = rep_metrics['labs_credenciados'] - rep_metrics['labs_ativos']
    
    # Calcular taxa de ativação
    rep_metrics['taxa_ativacao'] = (
        rep_metrics['labs_ativos'] / rep_metrics['labs_credenciados'] * 100
    ).fillna(0)
    
    # Calcular produtividade (coletas por lab ativo)
    rep_metrics['produtividade'] = (
        rep_metrics['total_coletas'] / rep_metrics['labs_ativos']
    ).fillna(0)
    
    # Renomear coluna para manter consistência
    rep_metrics = rep_metrics.rename(columns={rep_name_col: "name_rep"})
    
    # Ordenar por total de coletas
    rep_metrics = rep_metrics.sort_values('total_coletas', ascending=False)
    
    return rep_metrics


def compute_new_accreditations(
    df_labs_status: pd.DataFrame,
    current_date: datetime,
    months_back: int = 3
) -> pd.DataFrame:
    """
    Identifica novos credenciamentos nos últimos meses
    """
    # Calcular data limite
    from datetime import timedelta
    date_limit = current_date - timedelta(days=months_back * 30)
    
    # Garantir que createdAt seja datetime antes de usar operações de data
    df_labs_clean = df_labs_status.copy()
    if "createdAt" in df_labs_clean.columns:
        df_labs_clean["createdAt"] = pd.to_datetime(df_labs_clean["createdAt"], errors="coerce")
        # Remover linhas onde a conversão falhou (valores inválidos)
        df_labs_clean = df_labs_clean.dropna(subset=["createdAt"])
    
    # Filtrar labs credenciados recentemente
    new_labs = df_labs_clean[
        (df_labs_clean['is_credenciado'] == True) &
        (df_labs_clean['createdAt'] >= date_limit)
    ].copy()
    
    if new_labs.empty:
        return pd.DataFrame(columns=['name_rep', 'name_rep_clean', 'Categoria', 'fantasyName', 'cnpj', 'createdAt', 'data_credenciamento', 'dias_credenciado'])
    
    # Calcular dias desde credenciamento
    new_labs['dias_credenciado'] = (current_date - new_labs['createdAt']).dt.days
    
    # Formatar data do credenciamento para exibição
    new_labs['data_credenciamento'] = new_labs['createdAt'].apply(
        lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else "Data não disponível"
    )
    
    # Ordenar por data de credenciamento
    new_labs = new_labs.sort_values('createdAt', ascending=False)
    
    # Incluir name_rep_clean se disponível
    columns_to_return = ['name_rep', 'Categoria', 'fantasyName', 'cnpj', 'createdAt', 'data_credenciamento', 'dias_credenciado']
    if 'name_rep_clean' in new_labs.columns:
        columns_to_return.insert(1, 'name_rep_clean')
    
    return new_labs[columns_to_return]


def compute_inactive_labs_alert(
    df_labs_status: pd.DataFrame,
    df_gatherings_merged: pd.DataFrame,
    current_date: datetime,
    inactivity_threshold_days: int = 30
) -> pd.DataFrame:
    """
    Identifica labs que pararam de coletar (alerta para cobrança)
    Busca a última coleta de todos os anos, não apenas 2025
    """
    # Labs credenciados
    labs_cred = df_labs_status[df_labs_status['is_credenciado'] == True]
    
    # Última coleta por lab (de todos os anos)
    if df_gatherings_merged.empty:
        last_collections = pd.DataFrame(columns=['_laboratory', 'createdAt'])
    else:
        # Filtrar apenas coletas ativas (não test, não disabled)
        df_gatherings_active_all_years = filter_active_gatherings(df_gatherings_merged, exclude_test=False, exclude_disabled=True)
        
        # Garantir que createdAt seja datetime antes de calcular o máximo
        if "createdAt" in df_gatherings_active_all_years.columns:
            df_gatherings_active_all_years = df_gatherings_active_all_years.copy()
            df_gatherings_active_all_years["createdAt"] = pd.to_datetime(df_gatherings_active_all_years["createdAt"], errors="coerce")
            # Remover linhas onde a conversão falhou (valores inválidos)
            df_gatherings_active_all_years = df_gatherings_active_all_years.dropna(subset=["createdAt"])
        
        last_collections = df_gatherings_active_all_years.groupby('_laboratory')['createdAt'].max().reset_index()
    
    # Merge com labs credenciados (usar suffixes para evitar conflito)
    labs_with_last_collection = labs_cred.merge(
        last_collections,
        left_on='_id',
        right_on='_laboratory',
        how='left',
        suffixes=('_lab', '_coleta')
    )
    
    # Calcular dias desde última coleta (usar a coluna createdAt_coleta)
    def _days_diff_safe(dt: pd.Timestamp) -> int:
        if pd.isna(dt):
            return 999
        diff_days = (current_date - dt).days
        return max(diff_days, 0)

    labs_with_last_collection['dias_sem_coletar'] = labs_with_last_collection['createdAt_coleta'].apply(_days_diff_safe)
    
    # Formatar data da última coleta para exibição
    labs_with_last_collection['ultima_coleta_str'] = labs_with_last_collection['createdAt_coleta'].apply(
        lambda x: x.strftime("%d/%m/%Y %H:%M") if pd.notna(x) else "Sem coletas"
    )
    
    # Criar coluna de exibição para dias sem coletar
    labs_with_last_collection['dias_sem_coletar_display'] = labs_with_last_collection['dias_sem_coletar'].apply(
        lambda x: str(x) if x != 999 else "-"
    )
    
    # Filtrar labs inativos (sem coleta há mais de X dias)
    inactive_labs = labs_with_last_collection[
        labs_with_last_collection['dias_sem_coletar'] > inactivity_threshold_days
    ].copy()
    
    # Ordenar por dias sem coletar
    inactive_labs = inactive_labs.sort_values('dias_sem_coletar', ascending=False)
    
    # Incluir name_rep_clean se disponível
    columns_to_return = ['name_rep', 'Categoria', 'fantasyName', 'cnpj', 'ultima_coleta_str', 'dias_sem_coletar_display', 'dias_sem_coletar']
    if 'name_rep_clean' in inactive_labs.columns:
        columns_to_return.insert(1, 'name_rep_clean')
    
    return inactive_labs[columns_to_return]


def compute_category_summary(
    df_gatherings_active: pd.DataFrame,
    df_labs_status: pd.DataFrame
) -> dict:
    """
    Resumo por categoria (Interno vs Externo)
    """
    # Coletas por categoria
    coletas_categoria = df_gatherings_active.groupby('Categoria').agg({
        '_laboratory': 'nunique',
        'createdAt': 'count'
    }).reset_index()
    coletas_categoria.columns = ['Categoria', 'labs_ativos', 'total_coletas']
    
    # Labs credenciados por categoria
    labs_categoria = df_labs_status.groupby('Categoria').agg({
        '_id': 'count',
        'is_credenciado': 'sum'
    }).reset_index()
    labs_categoria.columns = ['Categoria', 'total_labs', 'labs_credenciados']
    
    # Merge
    summary = coletas_categoria.merge(labs_categoria, on='Categoria', how='outer')
    summary = summary.fillna(0)
    
    # Calcular métricas
    summary['taxa_ativacao'] = (summary['labs_ativos'] / summary['labs_credenciados'] * 100).fillna(0)
    summary['produtividade'] = (summary['total_coletas'] / summary['labs_ativos']).fillna(0)
    
    return summary.to_dict('records')


def compute_geographic_metrics(
    df_gatherings_active: pd.DataFrame,
    df_labs_status: pd.DataFrame,
    current_date: datetime,
    activity_window_days: int = DEFAULT_ACTIVITY_WINDOW_DAYS,
) -> pd.DataFrame:
    """
    Calcula métricas de performance por estado
    """
    # Verificar se as colunas necessárias existem
    required_cols = ['state_code', 'state_name']
    missing_cols = [col for col in required_cols if col not in df_gatherings_active.columns]
    
    if missing_cols:
        # Se não há dados geográficos, retornar DataFrame vazio
        return pd.DataFrame(columns=['state_code', 'state_name', 'labs_ativos', 'total_coletas', 
                                   'total_labs', 'labs_credenciados', 'labs_inativos', 
                                   'taxa_ativacao', 'produtividade'])
    
    # Filtrar apenas registros com dados de localização válidos
    df_gatherings_with_location = df_gatherings_active.dropna(subset=['state_code', 'state_name'])
    df_labs_with_location = df_labs_status.dropna(subset=['state_code', 'state_name'])
    
    if df_gatherings_with_location.empty:
        return pd.DataFrame(columns=['state_code', 'state_name', 'labs_ativos', 'total_coletas', 
                                   'total_labs', 'labs_credenciados', 'labs_inativos', 
                                   'taxa_ativacao', 'produtividade'])
    
    # Agrupar coletas por estado
    state_metrics = df_gatherings_with_location.groupby(['state_code', 'state_name']).agg({
        '_laboratory': 'nunique',  # Labs únicos que coletaram
        'createdAt': 'count'       # Total de coletas
    }).reset_index()
    state_metrics.columns = ['state_code', 'state_name', 'labs_ativos', 'total_coletas']
    
    # Calcular labs credenciados por estado
    labs_por_estado = df_labs_with_location.groupby(['state_code', 'state_name']).agg({
        '_id': 'count',  # Total de labs credenciados
        'is_credenciado': 'sum'  # Labs credenciados
    }).reset_index()
    labs_por_estado.columns = ['state_code', 'state_name', 'total_labs', 'labs_credenciados']
    
    # Merge das métricas
    state_metrics = state_metrics.merge(labs_por_estado, on=['state_code', 'state_name'], how='outer')
    state_metrics = state_metrics.fillna(0)
    
    # Calcular labs inativos (credenciados mas não coletando)
    state_metrics['labs_inativos'] = state_metrics['labs_credenciados'] - state_metrics['labs_ativos']
    
    # Calcular taxa de ativação
    state_metrics['taxa_ativacao'] = (
        state_metrics['labs_ativos'] / state_metrics['labs_credenciados'] * 100
    ).fillna(0)
    
    # Calcular produtividade (coletas por lab ativo)
    state_metrics['produtividade'] = (
        state_metrics['total_coletas'] / state_metrics['labs_ativos']
    ).fillna(0)
    
    # Ordenar por total de coletas
    state_metrics = state_metrics.sort_values('total_coletas', ascending=False)
    
    return state_metrics


def compute_city_metrics(
    df_gatherings_active: pd.DataFrame,
    df_labs_status: pd.DataFrame,
    current_date: datetime,
    activity_window_days: int = DEFAULT_ACTIVITY_WINDOW_DAYS,
) -> pd.DataFrame:
    """
    Calcula métricas de performance por cidade
    """
    # Verificar se as colunas necessárias existem
    required_cols = ['state_code', 'state_name', 'city']
    missing_cols = [col for col in required_cols if col not in df_gatherings_active.columns]
    
    if missing_cols:
        # Se não há dados geográficos, retornar DataFrame vazio
        return pd.DataFrame(columns=['state_code', 'state_name', 'city', 'labs_ativos', 'total_coletas', 
                                   'total_labs', 'labs_credenciados', 'labs_inativos', 
                                   'taxa_ativacao', 'produtividade'])
    
    # Filtrar apenas registros com dados de localização válidos
    df_gatherings_with_location = df_gatherings_active.dropna(subset=['state_code', 'state_name', 'city'])
    df_labs_with_location = df_labs_status.dropna(subset=['state_code', 'state_name', 'city'])
    
    if df_gatherings_with_location.empty:
        return pd.DataFrame(columns=['state_code', 'state_name', 'city', 'labs_ativos', 'total_coletas', 
                                   'total_labs', 'labs_credenciados', 'labs_inativos', 
                                   'taxa_ativacao', 'produtividade'])
    
    # Agrupar coletas por cidade
    city_metrics = df_gatherings_with_location.groupby(['state_code', 'state_name', 'city']).agg({
        '_laboratory': 'nunique',  # Labs únicos que coletaram
        'createdAt': 'count'       # Total de coletas
    }).reset_index()
    city_metrics.columns = ['state_code', 'state_name', 'city', 'labs_ativos', 'total_coletas']
    
    # Calcular labs credenciados por cidade
    labs_por_cidade = df_labs_with_location.groupby(['state_code', 'state_name', 'city']).agg({
        '_id': 'count',  # Total de labs credenciados
        'is_credenciado': 'sum'  # Labs credenciados
    }).reset_index()
    labs_por_cidade.columns = ['state_code', 'state_name', 'city', 'total_labs', 'labs_credenciados']
    
    # Merge das métricas
    city_metrics = city_metrics.merge(labs_por_cidade, on=['state_code', 'state_name', 'city'], how='outer')
    city_metrics = city_metrics.fillna(0)
    
    # Calcular labs inativos
    city_metrics['labs_inativos'] = city_metrics['labs_credenciados'] - city_metrics['labs_ativos']
    
    # Calcular taxa de ativação
    city_metrics['taxa_ativacao'] = (
        city_metrics['labs_ativos'] / city_metrics['labs_credenciados'] * 100
    ).fillna(0)
    
    # Calcular produtividade
    city_metrics['produtividade'] = (
        city_metrics['total_coletas'] / city_metrics['labs_ativos']
    ).fillna(0)
    
    # Ordenar por total de coletas
    city_metrics = city_metrics.sort_values('total_coletas', ascending=False)
    
    return city_metrics


