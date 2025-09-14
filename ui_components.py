from __future__ import annotations

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from typing import Dict
import locale

# Configurar locale para formatação brasileira
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        pass


def get_state_coordinates(state_code: str) -> tuple:
    """Retorna as coordenadas (lat, lon) do centro de cada estado brasileiro"""
    coordinates = {
        'AC': (-8.77, -70.55),  # Acre
        'AL': (-9.71, -35.73),  # Alagoas
        'AP': (0.90, -52.00),   # Amapá
        'AM': (-3.42, -65.86),  # Amazonas
        'BA': (-12.97, -38.50), # Bahia
        'CE': (-3.72, -38.54),  # Ceará
        'DF': (-15.78, -47.92), # Distrito Federal
        'ES': (-20.31, -40.31), # Espírito Santo
        'GO': (-16.64, -49.31), # Goiás
        'MA': (-2.53, -44.30),  # Maranhão
        'MT': (-15.60, -56.10), # Mato Grosso
        'MS': (-20.44, -54.64), # Mato Grosso do Sul
        'MG': (-19.92, -43.93), # Minas Gerais
        'PA': (-1.45, -48.50),  # Pará
        'PB': (-7.12, -34.86),  # Paraíba
        'PR': (-25.42, -49.27), # Paraná
        'PE': (-8.05, -34.92),  # Pernambuco
        'PI': (-5.09, -42.80),  # Piauí
        'RJ': (-22.91, -43.20), # Rio de Janeiro
        'RN': (-5.79, -35.21),  # Rio Grande do Norte
        'RS': (-30.03, -51.23), # Rio Grande do Sul
        'RO': (-8.76, -63.90),  # Rondônia
        'RR': (2.82, -60.67),   # Roraima
        'SC': (-27.59, -48.55), # Santa Catarina
        'SP': (-23.55, -46.64), # São Paulo
        'SE': (-10.90, -37.07), # Sergipe
        'TO': (-10.17, -48.33), # Tocantins
    }
    return coordinates.get(state_code, None)


def format_number_br(value: int) -> str:
    """Formata números no padrão brasileiro (1.234.567)"""
    try:
        return f"{value:,}".replace(",", ".")
    except:
        return str(value)


def kpi_cards(kpis: Dict[str, int], cred: int, descred: int, ativos: int, inativos: int, window_days: int):
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    # Formatar números para exibição brasileira
    col1.metric("Total de Coletas", format_number_br(kpis["total"]))
    col2.metric("Máximo Mensal", format_number_br(kpis["max"]))
    col3.metric("Mínimo Mensal", format_number_br(kpis["min"]))
    col4.metric("Média Mensal", format_number_br(kpis["avg"]))
    col5.metric("Labs Credenciados", format_number_br(cred))
    col6.metric(f"Labs Ativos ({window_days}d)", format_number_br(ativos))
    
    st.caption(f"Labs Inativos ({window_days}d): {format_number_br(inativos)} — Descredenciados: {format_number_br(descred)}")


def line_chart_monthly(monthly: pd.DataFrame):
    st.subheader("📊 Volume Mensal por Categoria")
    
    if monthly.empty:
        st.info("Nenhum dado mensal disponível.")
        return
    
    # Separar dados por categoria (Interno vs Externo)
    if 'Categoria' in monthly.columns:
        # Criar gráfico com três linhas (Internos, Externos e Total)
        fig = go.Figure()
        
        # Dados para Internos
        interno_data = monthly[monthly['Categoria'] == 'Interno']
        if not interno_data.empty:
            fig.add_trace(go.Scatter(
                x=interno_data['month'],
                y=interno_data['Volume'],
                mode='lines+markers',
                name='Representantes Internos',
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=8)
            ))
        
        # Dados para Externos
        externo_data = monthly[monthly['Categoria'] == 'Externo']
        if not externo_data.empty:
            fig.add_trace(go.Scatter(
                x=externo_data['month'],
                y=externo_data['Volume'],
                mode='lines+markers',
                name='Representantes Externos',
                line=dict(color='#ff7f0e', width=3),
                marker=dict(size=8)
            ))
        
        # Adicionar linha do TOTAL (soma de Internos + Externos)
        if not interno_data.empty and not externo_data.empty:
            # Criar DataFrame com totais
            total_data = interno_data.copy()
            total_data = total_data.merge(
                externo_data[['month', 'Volume']], 
                on='month', 
                suffixes=('_interno', '_externo')
            )
            total_data['Volume_total'] = total_data['Volume_interno'] + total_data['Volume_externo']
            
            fig.add_trace(go.Scatter(
                x=total_data['month'],
                y=total_data['Volume_total'],
                mode='lines+markers',
                name='TOTAL (Internos + Externos)',
                line=dict(color='#2ca02c', width=4, dash='dash'),
                marker=dict(size=10, symbol='diamond')
            ))
        
        fig.update_layout(
            title="Volume de Coletas por Mês",
            xaxis_title="Mês",
            yaxis_title="Quantidade de Coletas",
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        # Formatar tooltips em português
        fig.update_traces(
            hovertemplate="<b>%{fullData.name}</b><br>" +
                         "Mês: %{x}<br>" +
                         "Coletas: %{y:,.0f}<extra></extra>"
        )
        
    else:
        # Fallback para dados sem categoria
        fig = px.line(monthly, x="month", y="Volume", markers=True, title="Volume por Mês")
        fig.update_layout(xaxis_title="Mês", yaxis_title="Quantidade de Coletas")
        fig.update_traces(
            hovertemplate="<b>Volume Mensal</b><br>" +
                         "Mês: %{x}<br>" +
                         "Coletas: %{y:,.0f}<extra></extra>"
        )
    
    st.plotly_chart(fig, use_container_width=True)


def line_chart_weekly(weekly: pd.DataFrame):
    st.subheader("📈 Volume Semanal por Categoria")
    
    if weekly.empty:
        st.info("Nenhum dado semanal disponível.")
        return
    
    # Separar dados por categoria
    if 'Categoria' in weekly.columns:
        fig = go.Figure()
        
        # Dados para Internos
        interno_data = weekly[weekly['Categoria'] == 'Interno']
        if not interno_data.empty:
            fig.add_trace(go.Scatter(
                x=interno_data['week'],
                y=interno_data['Volume'],
                mode='lines+markers',
                name='Representantes Internos',
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=6)
            ))
        
        # Dados para Externos
        externo_data = weekly[weekly['Categoria'] == 'Externo']
        if not externo_data.empty:
            fig.add_trace(go.Scatter(
                x=externo_data['week'],
                y=externo_data['Volume'],
                mode='lines+markers',
                name='Representantes Externos',
                line=dict(color='#ff7f0e', width=3),
                marker=dict(size=6)
            ))
        
        # Adicionar linha do TOTAL (soma de Internos + Externos)
        if not interno_data.empty and not externo_data.empty:
            # Criar DataFrame com totais
            total_data = interno_data.copy()
            total_data = total_data.merge(
                externo_data[['week', 'Volume']], 
                on='week', 
                suffixes=('_interno', '_externo')
            )
            total_data['Volume_total'] = total_data['Volume_interno'] + total_data['Volume_externo']
            
            fig.add_trace(go.Scatter(
                x=total_data['week'],
                y=total_data['Volume_total'],
                mode='lines+markers',
                name='TOTAL (Internos + Externos)',
                line=dict(color='#2ca02c', width=4, dash='dash'),
                marker=dict(size=8, symbol='diamond')
            ))
        
        fig.update_layout(
            title="Volume de Coletas por Semana",
            xaxis_title="Semana",
            yaxis_title="Quantidade de Coletas",
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        fig.update_traces(
            hovertemplate="<b>%{fullData.name}</b><br>" +
                         "Semana: %{x}<br>" +
                         "Coletas: %{y:,.0f}<extra></extra>"
        )
        
    else:
        # Fallback para dados sem categoria
        fig = px.line(weekly, x="week", y="Volume", markers=True, title="Volume por Semana")
        fig.update_layout(xaxis_title="Semana", yaxis_title="Quantidade de Coletas")
        fig.update_traces(
            hovertemplate="<b>Volume Semanal</b><br>" +
                         "Semana: %{x}<br>" +
                         "Coletas: %{y:,.0f}<extra></extra>"
        )
    
    st.plotly_chart(fig, use_container_width=True)


def table(df: pd.DataFrame, header: str):
    st.subheader(header)
    
    # Formatar números na tabela
    df_display = df.copy()
    
    # Identificar colunas numéricas para formatação
    numeric_columns = df_display.select_dtypes(include=['int64', 'float64']).columns
    
    for col in numeric_columns:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(
                lambda x: format_number_br(int(x)) if pd.notna(x) and x != 999 else x
            )
    
    st.dataframe(df_display, use_container_width=True)


def download_button(df: pd.DataFrame, label: str, filename: str):
    st.download_button(
        label=label,
        data=df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8"),
        file_name=filename,
        mime="text/csv",
        use_container_width=True,
    )


def performance_dashboard(rep_metrics: pd.DataFrame, category_summary: list):
    """Dashboard específico para gestão comercial"""
    st.header("🎯 Dashboard de Performance Comercial")
    
    # KPIs principais
    col1, col2, col3, col4 = st.columns(4)
    
    for summary in category_summary:
        if summary['Categoria'] == 'Interno':
            with col1:
                st.metric(
                    "Internos - Labs Ativos", 
                    f"{int(summary['labs_ativos']):,}".replace(",", ".")
                )
            with col2:
                st.metric(
                    "Internos - Total Coletas", 
                    f"{int(summary['total_coletas']):,}".replace(",", ".")
                )
        elif summary['Categoria'] == 'Externo':
            with col3:
                st.metric(
                    "Externos - Labs Ativos", 
                    f"{int(summary['labs_ativos']):,}".replace(",", ".")
                )
            with col4:
                st.metric(
                    "Externos - Total Coletas", 
                    f"{int(summary['total_coletas']):,}".replace(",", ".")
                )
    
    # Gráfico de pizza para distribuição
    if category_summary:
        fig = px.pie(
            values=[s['total_coletas'] for s in category_summary],
            names=[s['Categoria'] for s in category_summary],
            title="Distribuição de Coletas por Categoria",
            color_discrete_map={'Interno': '#1f77b4', 'Externo': '#ff7f0e'}
        )
        fig.update_traces(
            textposition='inside',
            textinfo='percent+label',
            hovertemplate="<b>%{label}</b><br>" +
                         "Coletas: %{value:,.0f}<br>" +
                         "Percentual: %{percent}<extra></extra>"
        )
        st.plotly_chart(fig, use_container_width=True)


def representative_table(df: pd.DataFrame, title: str):
    """Tabela específica para representantes com formatação melhorada"""
    st.subheader(title)
    
    # Formatar colunas específicas
    df_display = df.copy()
    
    # Formatar números
    numeric_cols = ['Volume', 'labs_credenciados', 'labs_ativos', 'labs_inativos', 
                   'total_coletas', 'taxa_ativacao', 'produtividade']
    
    for col in numeric_cols:
        if col in df_display.columns:
            if col in ['taxa_ativacao']:
                df_display[col] = df_display[col].apply(
                    lambda x: f"{x:.1f}%" if pd.notna(x) else "-"
                )
            elif col in ['produtividade']:
                df_display[col] = df_display[col].apply(
                    lambda x: f"{x:.1f}" if pd.notna(x) else "-"
                )
            else:
                df_display[col] = df_display[col].apply(
                    lambda x: format_number_br(int(x)) if pd.notna(x) and x != 999 else "-"
                )
    
    st.dataframe(df_display, use_container_width=True)


def geographic_dashboard(state_metrics: pd.DataFrame, city_metrics: pd.DataFrame):
    """Dashboard específico para análise geográfica"""
    st.header("🗺️ Análise Geográfica - Performance por Estado")
    
    if state_metrics.empty:
        st.info("Nenhum dado geográfico disponível. Verificando dados de localização...")
        
        # Debug: mostrar informações sobre os dados
        st.write("**Debug Info:**")
        st.write(f"- State metrics shape: {state_metrics.shape}")
        st.write(f"- City metrics shape: {city_metrics.shape}")
        st.write(f"- State metrics columns: {list(state_metrics.columns)}")
        return
    
    # KPIs principais por estado
    col1, col2, col3, col4 = st.columns(4)
    
    # Top estados
    top_states = state_metrics.head(4)
    
    for i, (_, state) in enumerate(top_states.iterrows()):
        with [col1, col2, col3, col4][i]:
            st.metric(
                f"{state['state_name']} - Total Coletas", 
                f"{int(state['total_coletas']):,}".replace(",", ".")
            )
    
    # Gráfico de barras - Top 10 estados por volume
    st.subheader("📊 Top 10 Estados por Volume de Coletas")
    
    top_10_states = state_metrics.head(10)
    
    fig = px.bar(
        top_10_states,
        x='state_name',
        y='total_coletas',
        title="Volume de Coletas por Estado",
        color='total_coletas',
        color_continuous_scale='viridis'
    )
    
    fig.update_layout(
        xaxis_title="Estado",
        yaxis_title="Total de Coletas",
        xaxis={'categoryorder':'total descending'},
        showlegend=False
    )
    
    fig.update_traces(
        hovertemplate="<b>%{x}</b><br>" +
                     "Coletas: %{y:,.0f}<extra></extra>"
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Gráfico de pizza - Distribuição por estado
    st.subheader("🥧 Distribuição de Coletas por Estado")
    
    fig_pie = px.pie(
        state_metrics,
        values='total_coletas',
        names='state_name',
        title="Distribuição Percentual por Estado"
    )
    
    fig_pie.update_traces(
        textposition='inside',
        textinfo='percent+label',
        hovertemplate="<b>%{label}</b><br>" +
                     "Coletas: %{value:,.0f}<br>" +
                     "Percentual: %{percent}<extra></extra>"
    )
    
    st.plotly_chart(fig_pie, use_container_width=True)
    
    # Mapa de calor (choropleth) - Estados por produtividade
    st.subheader("🔥 Mapa de Produtividade por Estado")
    
    # Preparar dados para o mapa
    map_data = state_metrics.copy()
    map_data['produtividade_formatted'] = map_data['produtividade'].apply(
        lambda x: f"{x:.1f}" if pd.notna(x) and x > 0 else "0.0"
    )
    
    # Criar mapa usando plotly - versão alternativa para estados brasileiros
    try:
        # Criar mapa usando scattergeo para melhor compatibilidade
        fig_map = go.Figure()
        
        # Adicionar scattergeo para cada estado
        fig_map.add_trace(go.Scattergeo(
            lon=map_data['state_code'].apply(lambda x: get_state_coordinates(x)[1] if get_state_coordinates(x) else None),
            lat=map_data['state_code'].apply(lambda x: get_state_coordinates(x)[0] if get_state_coordinates(x) else None),
            text=map_data['state_name'],
            mode='markers',
            marker=dict(
                size=map_data['produtividade'].apply(lambda x: max(10, min(50, x/10))),  # Tamanho baseado na produtividade
                color=map_data['produtividade'],
                colorscale='RdYlGn',
                showscale=True,
                colorbar=dict(title="Produtividade (coletas/lab)")
            ),
            hovertemplate="<b>%{text}</b><br>" +
                         "Produtividade: %{marker.color:.1f} coletas/lab<br>" +
                         "<extra></extra>"
        ))
        
        fig_map.update_layout(
            title="Produtividade por Estado (coletas/lab ativo)",
            geo=dict(
                scope='south america',
                showland=True,
                landcolor='lightgray',
                showocean=True,
                oceancolor='lightblue',
                showcountries=True,
                countrycolor='white',
                showcoastlines=True,
                coastlinecolor='white',
                center=dict(lat=-15.7801, lon=-47.9292),
                projection_scale=2
            ),
            height=600
        )
        
        st.plotly_chart(fig_map, use_container_width=True)
        
    except Exception as e:
        st.error(f"Erro ao criar mapa: {str(e)}")
        st.info("Exibindo dados em formato de tabela em vez do mapa.")
        
        # Fallback: mostrar dados em tabela
        st.dataframe(map_data[['state_name', 'produtividade', 'total_coletas', 'labs_ativos']])
    
    # Tabela detalhada por estado
    st.subheader("📋 Métricas Detalhadas por Estado")
    geographic_table(state_metrics, "Performance por Estado")
    
    # Análise por cidade (top 20)
    if not city_metrics.empty:
        st.subheader("🏙️ Top 20 Cidades por Volume")
        top_cities = city_metrics.head(20)
        geographic_table(top_cities, "Performance por Cidade")


def geographic_table(df: pd.DataFrame, title: str):
    """Tabela específica para dados geográficos com formatação melhorada"""
    st.subheader(title)
    
    # Formatar colunas específicas
    df_display = df.copy()
    
    # Formatar números
    numeric_cols = ['total_coletas', 'labs_credenciados', 'labs_ativos', 'labs_inativos', 
                   'taxa_ativacao', 'produtividade']
    
    for col in numeric_cols:
        if col in df_display.columns:
            if col in ['taxa_ativacao']:
                df_display[col] = df_display[col].apply(
                    lambda x: f"{x:.1f}%" if pd.notna(x) else "-"
                )
            elif col in ['produtividade']:
                df_display[col] = df_display[col].apply(
                    lambda x: f"{x:.1f}" if pd.notna(x) else "-"
                )
            else:
                df_display[col] = df_display[col].apply(
                    lambda x: format_number_br(int(x)) if pd.notna(x) and x != 999 else "-"
                )
    
    # Renomear colunas para português
    column_mapping = {
        'state_name': 'Estado',
        'city': 'Cidade',
        'total_coletas': 'Total de Coletas',
        'labs_credenciados': 'Labs Credenciados',
        'labs_ativos': 'Labs Ativos',
        'labs_inativos': 'Labs Inativos',
        'taxa_ativacao': 'Taxa de Ativação (%)',
        'produtividade': 'Produtividade (coletas/lab)',
    }
    
    df_display = df_display.rename(columns=column_mapping)
    
    # Selecionar apenas colunas relevantes para exibição (evitar duplicatas)
    display_cols = []
    
    # Adicionar colunas na ordem correta, verificando se existem
    if 'Estado' in df_display.columns:
        display_cols.append('Estado')
    if 'Cidade' in df_display.columns:
        display_cols.append('Cidade')
    
    # Adicionar outras colunas que existem no DataFrame
    for col in column_mapping.values():
        if col in df_display.columns and col not in display_cols:
            display_cols.append(col)
    
    # Verificar se há colunas para exibir
    if display_cols:
        st.dataframe(df_display[display_cols], use_container_width=True)
    else:
        st.info("Nenhuma coluna válida encontrada para exibição.")


