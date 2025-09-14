# 🏢 Dashboard de Gestão Comercial - ToxRepresentatives

Dashboard interativo para gestão comercial de representantes e laboratórios, desenvolvido especificamente para gestores comerciais que lidam com representantes internos e externos.

## 🎯 Foco no Gestor Comercial

Este dashboard foi redesenhado pensando nas necessidades específicas do gestor comercial:

- **Visualização clara** de performance por categoria (Internos vs Externos)
- **Números formatados** no padrão brasileiro (1.234.567)
- **Nomes limpos** dos representantes (sem prefixos técnicos)
- **Interface em português** com labels claros
- **Layout otimizado** para tomada de decisões comerciais

## ✨ Principais Melhorias

### 📊 Gráficos Separados por Categoria
- **Volume Mensal**: Linhas separadas para Representantes Internos e Externos
- **Volume Semanal**: Diferenciação visual por categoria
- **Tooltips em português** com formatação brasileira

### 🏷️ Formatação Brasileira
- **Números**: 1.234.567 em vez de "1,234k"
- **Labels claros**: "Quantidade de Coletas" em vez de "Volume"
- **Headers em português**: "Nome do Representante" em vez de "name_rep"

### 👥 Nomes Limpos dos Representantes
- Remove prefixos técnicos: "EXT-", "INT-", "CAEPTOX -"
- Exemplo: "EXT- GLAUDYSON BARBOZA DE MOURA" → "GLAUDYSON BARBOZA DE MOURA"
- Facilita identificação e comunicação

### 🎯 Layout Reorganizado
1. **📊 Dashboard Geral** - Visão macro do negócio
2. **🎯 Gestão Comercial** - Performance dos representantes
3. **🏆 Ranking** - Top performers
4. **🏥 Laboratórios** - Status operacional
5. **⚠️ Alertas** - Labs inativos

## 🚀 Como Executar

### Pré-requisitos
```bash
pip install -r requirements.txt
```

### Execução
```bash
streamlit run app.py
```

## 📋 Funcionalidades

### Dashboard Geral
- KPIs principais com formatação brasileira
- Gráfico mensal com linhas separadas por categoria
- Métricas de credenciamento e atividade

### Gestão Comercial
- **Dashboard de Performance**: KPIs por categoria com gráfico de pizza
- **Performance Detalhada**: Tabela com métricas por representante
- **Novos Credenciamentos**: Acompanhamento de novos labs

### Ranking
- **Top Representantes**: Ranking por volume de coletas
- **Top Laboratórios**: Performance dos labs
- **Download de relatórios** em CSV

### Visão Geral
- **Análise Semanal**: Tendências semanais por categoria (movida para Visão Geral)
- **Detalhamento Mensal**: Tabela/Gráfico com números brasileiros

### Laboratórios
- **Status Operacional**: Credenciamento e atividade
- **Última Coleta**: Dias desde a última atividade

### Alertas
- **Labs Inativos**: Identificação de labs sem coleta
- **Agrupamento por Representante**: Para ações comerciais
- **Relatórios de Alerta**: Para cobrança e follow-up

## 🔧 Configurações

### Filtros Disponíveis
- **Ano**: Seleção do período
- **Janela de Atividade**: Dias para considerar "ativo" (7-60 dias)
- **Tipo de Representante**: Todos, Interno, Externo
- **Representantes Específicos**: Seleção múltipla
- **Busca por Laboratório/CNPJ**: Campo com sugestões

### Alertas Configuráveis
- **Threshold de Inatividade**: 15-90 dias (padrão: 30)
- **Relatórios Automáticos**: CSV com encoding UTF-8

## 📊 Métricas Comerciais

### Por Representante
- **Labs Credenciados**: Total de laboratórios
- **Labs Ativos**: Laboratórios com coleta recente
- **Taxa de Ativação**: % de labs ativos
- **Produtividade**: Coletas por lab ativo

### Por Categoria
- **Internos vs Externos**: Comparação de performance
- **Distribuição**: Gráfico de pizza das coletas
- **Tendências**: Evolução temporal

## 🎨 Interface

### Cores e Símbolos
- **Internos**: Azul (#1f77b4)
- **Externos**: Laranja (#ff7f0e)
- **Ícones**: Emojis para facilitar navegação
- **Formatação**: Números com separadores brasileiros

### Responsividade
- **Layout Wide**: Otimizado para telas grandes
- **Tabelas Responsivas**: Scroll horizontal quando necessário
- **Gráficos Interativos**: Zoom, pan, hover

## 📈 Exemplos de Uso

### Para Gestores Comerciais
1. **Acompanhamento Diário**: Dashboard Geral
2. **Análise de Performance**: Gestão Comercial
3. **Identificação de Oportunidades**: Ranking
4. **Ações Corretivas**: Alertas

### Para Representantes
1. **Auto-avaliação**: Ranking individual
2. **Comparação**: Performance vs outros
3. **Metas**: Evolução temporal

## 🔄 Atualizações

### Versão 2.0 - Foco Comercial
- ✅ Gráficos separados por categoria
- ✅ Formatação brasileira
- ✅ Nomes limpos dos representantes
- ✅ Interface em português
- ✅ Layout reorganizado
- ✅ Dashboard de performance
- ✅ Alertas configuráveis

## 📞 Suporte

Para dúvidas ou sugestões sobre o dashboard, entre em contato com a equipe de desenvolvimento.

---

**Desenvolvido com foco na experiência do gestor comercial** 🎯
