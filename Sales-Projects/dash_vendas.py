"""
Projeto: Farol de Vendas - Dashboard Interativo
* @copyrigth    Sávio Silas <svosilas@gmail.com> - DEV Portal Vidros
* @date         11 Março 2024
* @file         das_vendas.py
* @brief
    Este script é responsável por criar um dashboard interativo de vendas utilizando Dash,
    uma framework Python para construção de aplicações web.

    O dashboard oferece as seguintes funcionalidades:
    - Filtragem de dados por vendedor.
    - Visualização da meta geral e realização em tempo real.
    - Projeções de vendas com base nos dados atuais.
    - Pontuação destacada do vendedor com base no desempenho.
    - Gráficos interativos de vendas por categoria e localidade.
    - Tabelas de faturamento detalhadas por subcategorias de produtos.
    Callbacks são utilizados para atualizar os componentes do aplicativo em resposta às interações do usuário.
    A estrutura do layout é organizada utilizando dash_bootstrap_components
    O dashboard é iniciado em um servidor local e pode ser acessado via navegador web para
    uma interação ao vivo com os dados.
"""

import dash, base64
import dash_bootstrap_components as dbc
from dash import html, dcc, Input, Output, State, dcc, dash_table
import pandas as pd
import mysql.connector
import numpy as np
from dash.exceptions import PreventUpdate
import plotly.graph_objs as go
from datetime import datetime, timedelta
from io import BytesIO
from dash import dash_table
from dash_table.Format import Format, Scheme, Symbol, Group
from dash.dependencies import Input, Output, State, MATCH, ALL
from pandas.tseries.offsets import MonthEnd, BDay

# Consulta
df = fetch_data()

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Define o start_date como o primeiro dia do mês atual
start_date_realizado = pd.to_datetime('today').normalize().replace(day=1)
# Define o end_date como o dia atual
end_date_realizado = pd.to_datetime('today').normalize()

def calc_realizado(df_):
    df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')
    df_filtered = df_[df_['Data_Pedido'].between(start_date_realizado, end_date_realizado)]
    df_filtered_unique = df_filtered.drop_duplicates(subset='Id_Pedido', keep='first')

    return df_filtered_unique['TOTAL'].sum()

def get_vendedor_names(df):
    options = [{'label': name, 'value': name} for name in df['Vendedor'].unique()]
    options.insert(0, {'label': 'TODOS OS VENDEDORES', 'value': 'TODOS OS VENDEDORES'})  # Insere na primeira posição
    return options

# No layout do app, defina o valor padrão para o dropdown
dcc.Dropdown(
    id='vendedor-dropdown',
    options=get_vendedor_names(df),
    value='TODOS OS VENDEDORES',  # Valor padrão
    clearable=False,
)

def dias_uteis_ate_ontem(start_date, end_date):
    # Converte as datas para o formato 'datetime64[D]'
    start_date_fmt = np.datetime64(start_date, 'D')
    # Ajusta end_date para incluir 'ontem' na contagem, adicionando um dia
    end_date_ajustado = end_date + np.timedelta64(1, 'D')
    end_date_fmt = np.datetime64(end_date_ajustado, 'D')
    
    # Retorna a contagem de dias úteis entre start_date e end_date (incluindo 'ontem')
    return np.busday_count(start_date_fmt, end_date_fmt)

def total_dias_uteis_no_mes(start_date_realizado):
    # Assegura que start_date esteja no primeiro dia do mês
    start_date_realizado = start_date_realizado.replace(day=1)
    start_date_fmt = np.datetime64(start_date_realizado, 'D')
    # Calcula o último dia do mês
    last_day_of_month = start_date_realizado + pd.offsets.MonthEnd(1)
    last_day_of_month_fmt = np.datetime64(last_day_of_month, 'D')

    return np.busday_count(start_date_fmt, last_day_of_month_fmt + np.timedelta64(1, 'D'))

def calc_projecao_geral(valor_realizado):
    dias_uteis_ate_ontem_ = dias_uteis_ate_ontem(start_date_realizado, end_date_realizado - pd.Timedelta(days=1))
    total_dias = total_dias_uteis_no_mes(start_date_realizado)
    if dias_uteis_ate_ontem_ > 0:  # Evitar divisão por zero
        return (valor_realizado / dias_uteis_ate_ontem_) * total_dias
    return 0

@app.callback(
    [Output("meta-geral-input", "value"),
     Output("meta-value-store", "data")],
    [Input("meta-geral-input", "n_submit")],
    [State("meta-geral-input", "value")]
)
def formatar_valor_meta(n_submit, valor):
    if n_submit > 0:
        try:
            # Converter o valor para float e armazenar sem formatação
            valor_numerico = float(valor.replace(",", "").replace(".", "").replace("R$", "").strip()) / 100
            valor_formatado = f"R$ {valor_numerico:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return valor_formatado, valor_numerico  # Retorna o valor formatado e armazena o valor numérico
        except ValueError:
            # Caso o valor não possa ser convertido para float, retorne o valor atual e None para o armazenamento
            return valor, None
    return valor, None  # Se n_submit for 0, apenas retorne o valor atual e None

@app.callback(
    Output('escalation-card', 'children'), 
    [Input('meta-value-store', 'data')]  # Usa o valor numérico armazenado
)
def update_escalation_card(meta_value):
    if meta_value is None:
        raise PreventUpdate
    # 95% e 90% da meta
    meta_95 = meta_value * 0.95
    meta_90 = meta_value * 0.90

    # Formatar os valores calculados como moeda
    formatted_meta_95 = f"R$ {meta_95:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    formatted_meta_90 = f"R$ {meta_90:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    # Construir e retornar o conteúdo atualizado do card de escalonamento
    return [dbc.CardBody([
        html.H5("ESCALONAMENTO", className="card-title"),
        html.P(f"95% da Meta: {formatted_meta_95}", className="card-text"),
        html.P(f"90% da Meta: {formatted_meta_90}", className="card-text"),
    ])]

@app.callback(
    Output("meta-vendedor-input", "value"),
    [Input("meta-vendedor-input", "n_submit")],
    [State("meta-vendedor-input", "value")]
)
def formatar_meta_vendedor(n_submit, valor):
    if n_submit > 0:
        try:
            # Removendo caracteres de formatação e convertendo para float
            valor_numerico = float(valor.replace("R$", "").replace(".", "").replace(",", ".").strip())
            # Formatar o valor como moeda
            valor_formatado = "R$ {:,.2f}".format(valor_numerico).replace(",", "X").replace(".", ",").replace("X", ".")
            return valor_formatado
        except ValueError:
            # Caso o valor não possa ser convertido para float, retornar o valor atual
            return valor
    return valor

@app.callback(
    Output("realizado-vendedor", "children"),
    [Input("vendedor-dropdown", "value")]
)
def calcular_realizado_vendedor(vendedor_selecionado):
    if not vendedor_selecionado:
        # Se nenhum vendedor estiver selecionado, não há o que calcular
        return "R$ 0,00"

    # Garantindo que a coluna 'Data_Pedido' esteja no formato correto
    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y', errors='coerce')

    # Filtrando o DataFrame pelo vendedor selecionado e pelo período definido
    df_filtrado = df[(df['Vendedor'] == vendedor_selecionado) & 
                     (df['Data_Pedido'] >= start_date_realizado) & 
                     (df['Data_Pedido'] <= end_date_realizado)]

    # Removendo duplicatas com base na coluna 'Id_Pedido', mantendo a primeira ocorrência
    df_filtrado_unico = df_filtrado.drop_duplicates(subset='Id_Pedido', keep='first')

    # Se df_filtrado_unico estiver vazio após os filtros, retorna R$ 0,00
    if df_filtrado_unico.empty:
        return "R$ 0,00"

    # Calcular o valor realizado somando a coluna 'TOTAL'
    valor_realizado_vendedor = df_filtrado_unico['TOTAL'].sum()

    # Formatar o valor como moeda
    valor_formatado = "R$ {:,.2f}".format(valor_realizado_vendedor).replace(",", "X").replace(".", ",").replace("X", ".")

    return valor_formatado

def calc_projecao_vendedor(df_, vendedor_selecionado):
    # Garantindo que a coluna 'Data_Pedido' esteja no formato correto
    df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')

    # Filtrando o DataFrame pelo vendedor selecionado e pelo período
    df_filtrado = df_[(df_['Vendedor'] == vendedor_selecionado) & (df_['Data_Pedido'] >= start_date_realizado) & (df_['Data_Pedido'] <= end_date_realizado)]

    # Removendo duplicatas com base na coluna 'Id_Pedido', mantendo a primeira ocorrência
    df_filtrado_unico = df_filtrado.drop_duplicates(subset='Id_Pedido', keep='first')

    # Calculando o valor realizado pelo vendedor selecionado
    valor_realizado_vendedor = df_filtrado_unico['TOTAL'].sum()

    # Calculando os dias úteis até ontem e o total de dias úteis no mês
    dias_uteis_ate_ontem_ = dias_uteis_ate_ontem(start_date_realizado, end_date_realizado - pd.Timedelta(days=1))
    total_dias_uteis = total_dias_uteis_no_mes(start_date_realizado)

    # Calculando a projeção
    if dias_uteis_ate_ontem_ > 0:  # Evitar divisão por zero
        projecao = (valor_realizado_vendedor / dias_uteis_ate_ontem_) * total_dias_uteis
    else:
        projecao = 0

    return projecao

@app.callback(
    Output("projecao-vendedor", "children"),
    [Input("vendedor-dropdown", "value")]
)
def atualizar_projecao_vendedor(vendedor_selecionado):
    if not vendedor_selecionado:
        raise PreventUpdate

    projecao = calc_projecao_vendedor(df, vendedor_selecionado)
    # Formatar a projeção como moeda
    projecao_formatada = "R$ {:,.2f}".format(projecao).replace(",", "X").replace(".", ",").replace("X", ".")
    
    return projecao_formatada

#################### GRÁFICO DE LINHA
def aggregate_daily_sales(df_, vendedor_selecionado):
    end_date = pd.to_datetime('today').normalize()
    start_date = end_date.replace(day=1)
    
    df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')
    
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df_filtered = df_[(df_['Data_Pedido'].between(start_date, end_date)) & (df_['Vendedor'] == vendedor_selecionado)]
    else:
        df_filtered = df_[df_['Data_Pedido'].between(start_date, end_date)]
        
    df_filtered_unique = df_filtered.drop_duplicates(subset='Id_Pedido', keep='first')
    
    daily_sales = df_filtered_unique.groupby('Data_Pedido')['TOTAL'].sum().reset_index()
    
    return daily_sales

@app.callback(
    Output('right-chart', 'figure'),  # Assumindo que o ID do gráfico é 'right-chart'
    [Input('vendedor-dropdown', 'value')]
)
def update_line_chart(vendedor_selecionado):
    filtered_data = aggregate_daily_sales(df, vendedor_selecionado)
    return generate_line_chart(filtered_data)

# Função para gerar o gráfico de linha com os dados agregados
def generate_line_chart(agg_df):
    trace = go.Scatter(
        x=agg_df['Data_Pedido'],
        y=agg_df['TOTAL'],
        mode='lines+markers',
        name='Faturamento',
        line=dict(shape='spline', smoothing=1.3, width=4),
        marker=dict(color='#5B84BC', size=10, line=dict(color='white', width=2)),
        hovertemplate='Dia %{x}<br>R$ %{y:,.2f}<extra></extra>',
        fill='tozeroy',
        fillcolor='rgba(135, 206, 250, 0.2)'
    )

    layout = go.Layout(
        title='FATURAMENTO POR DIA',
        xaxis=dict(title='', tickformat='%d/%m/%y'),
        yaxis=dict(title='', showgrid=True, gridcolor='lightgrey'),
        hovermode='closest',
        plot_bgcolor="white",
        autosize=True,  # Garante que o layout será ajustado automaticamente
        margin=go.layout.Margin(l=30, r=30, b=50, t=50)  # Margens ajustadas para encaixar no card
    )

    return {'data': [trace], 'layout': layout}
    
# Define o end_date como o dia atual
end_date_3_meses = pd.to_datetime('today').normalize()
# Define o start_date como o primeiro dia do mês, 3 meses antes do mês atual
start_date_3_meses = (end_date_3_meses - pd.DateOffset(months=2)).replace(day=1)

#################### GRÁFICO DE PILHA
def apply_discount(row):
    if row['Tipo_Desconto'] == 'Porcentagem':
        return max(0, row['total_produto'] - (row['total_produto'] * row['Desconto'] / 100))
    elif row['Tipo_Desconto'] == 'Reais':
        valor_frete = row.get('Valor_Frete', 0)
        return ((row['total_produto'] * row['Desconto']) / ((row['TOTAL'] - valor_frete) + row['Desconto']) - row['total_produto']) * (-1)
    return row['total_produto'] 

# Função para calcular as somas por mês e categoria
def calcular_somas(df, categorias, inicio_mes, fim_mes, vendedor_selecionado):
    # Assegurar que 'Data_Pedido' esteja no formato correto de datetime
    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'])

    # Fazendo a comparação usando datas de datetime explicitamente
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df_filtrado = df[(df['Data_Pedido'] >= inicio_mes) & 
                         (df['Data_Pedido'] <= fim_mes) & 
                         (df['Grupo'].isin(categorias)) & 
                         (df['Vendedor'] == vendedor_selecionado)]
    else:
        df_filtrado = df[(df['Data_Pedido'] >= inicio_mes) & 
                         (df['Data_Pedido'] <= fim_mes) & 
                         (df['Grupo'].isin(categorias))]

    df_filtrado['total_produto_com_desconto'] = df_filtrado.apply(apply_discount, axis=1)
    return df_filtrado['total_produto_com_desconto'].sum()

@app.callback(
    Output('stack-chart', 'figure'),
    [Input('vendedor-dropdown', 'value')]
)
def update_stack_chart(vendedor_selecionado):
    # Calculando as somas para cada mês e categoria com o filtro do vendedor
    somas_agregadas = [calcular_somas(df, categorias_agregadas, inicio, fim, vendedor_selecionado) for inicio, fim in meses]
    somas_vidro = [calcular_somas(df, categoria_vidro, inicio, fim, vendedor_selecionado) for inicio, fim in meses]
    
    # Criando o novo gráfico de pilha empilhadas com os dados atualizados
    fig = go.Figure(data=[
        go.Bar(name='Vidro', x=nomes_meses, y=somas_vidro, marker_color='blue'),
        go.Bar(name='Agregadas', x=nomes_meses, y=somas_agregadas, marker_color='green')
    ])
    fig.update_layout(barmode='stack', title='Vendas por Categoria nos Últimos 3 Meses')

    return fig

#################### Card Venda por Localidade 
def calcular_vendas_por_localidade(df, vendedor_selecionado=None):
    df = df.drop_duplicates(subset='Id_Pedido', keep='first')
    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y')
    
    mes_atual = datetime.now().month
    ano_atual = datetime.now().year
    df_mes_atual = df[(df['Data_Pedido'].dt.month == mes_atual) & (df['Data_Pedido'].dt.year == ano_atual)]
    
    if vendedor_selecionado and vendedor_selecionado != "TODOS OS VENDEDORES":
        df_mes_atual = df_mes_atual[df_mes_atual['Vendedor'] == vendedor_selecionado]
    
    vendas_capital = df_mes_atual[df_mes_atual['Cidade'].str.lower() == 'manaus']['TOTAL'].sum()
    vendas_interior = df_mes_atual[df_mes_atual['Cidade'].str.lower() != 'manaus']['TOTAL'].sum()
    
    return vendas_capital, vendas_interior

@app.callback(
    [Output('vendas-capital', 'children'),
     Output('vendas-interior', 'children')],
    [Input('vendedor-dropdown', 'value')]
)
def update_vendas_por_localidade(vendedor_selecionado):
    vendas_capital, vendas_interior = calcular_vendas_por_localidade(df, vendedor_selecionado)
    return [f"R$ {vendas_capital:,.2f}", f"R$ {vendas_interior:,.2f}"]

# Callback para atualizar o card de "Clientes Atendidos - Grupos Específicos"
@app.callback(
    [Output('clientes-atendidos-vidro', 'children'),
     Output('clientes-atendidos-agregados', 'children'),
     Output('clientes-atendidos-temperados', 'children')],
    [Input('vendedor-dropdown', 'value')]
)
def update_clientes_atendidos(vendedor_selecionado):
    clientes_vidro, clientes_agregados, clientes_temperado = contar_clientes_grupos(df, vendedor_selecionado)
    return [
        f"QTD. {clientes_vidro}",
        f"QTD. {clientes_agregados}",
        f"QTD. {clientes_temperado}"
    ]

# #################### Card qtd clientes atendidos
def contar_clientes_grupos(df, vendedor_selecionado=None):
    # Lista dos grupos desejados para 'Agregados'
    grupos_agregados = ['ACESSÓRIOS', 'ALUMÍNIO', 'FERRAGEM', 'KIT PARA BOX PADRÃO', 'SILICONE']

    # Converter 'Data_Pedido' para datetime, se ainda não for
    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y')

    # Filtrar o DataFrame pelo mês e ano atual
    mes_atual = datetime.now().month
    ano_atual = datetime.now().year

    if vendedor_selecionado and vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]

    df_mes_atual = df[(df['Data_Pedido'].dt.month == mes_atual) & (df['Data_Pedido'].dt.year == ano_atual)]

    # Filtrar por grupos 'Agregados'
    df_agregados = df_mes_atual[df_mes_atual['Grupo'].isin(grupos_agregados)]
    clientes_agregados = df_agregados['Matriz_Cliente'].nunique()

    # Filtrar por grupo 'VIDRO' na matriz (considerado 'Temperado')
    df_temperado = df_mes_atual[(df_mes_atual['Grupo'] == 'VIDRO') & (df_mes_atual['Loja'] == 'PORTAL VIDROS (MATRIZ INDÚSTRIA)')]
    clientes_temperado = df_temperado['Matriz_Cliente'].nunique()

    # Filtrar por grupo 'VIDRO' na filial (considerado 'Vidro Comum'), excluindo 'TÁBUA' do subgrupo
    df_vidro_comum = df_mes_atual[(df_mes_atual['Grupo'] == 'VIDRO') & (df_mes_atual['Subgrupo'] != 'TÁBUA') & (df_mes_atual['Loja'] == 'PORTAL VIDROS (FILIAL)')]
    clientes_vidro_comum = df_vidro_comum['Matriz_Cliente'].nunique()

    return clientes_agregados, clientes_vidro_comum, clientes_temperado

# #################### Card RECOMPRA
@app.callback(
    Output("recompra-ultimos-6-meses", "children"),
    [Input("vendedor-dropdown", "value")]
)
def update_recompra_ultimos_6_meses(vendedor_selecionado):
    global df  # Indica que df é uma variável global
    end_date = pd.to_datetime("today").normalize()
    start_date = (end_date - pd.DateOffset(months=6)).replace(day=1)

    # Usa uma nova variável para o DataFrame filtrado
    df_filtrado = df
    if vendedor_selecionado and vendedor_selecionado != "TODOS OS VENDEDORES":
        df_filtrado = df[df['Vendedor'] == vendedor_selecionado]

    results = []

    for month_offset in range(7):  # Inclui um mês extra para cálculo de positivação
        month_start = start_date + pd.DateOffset(months=month_offset)
        month_end = (month_start + pd.DateOffset(months=1)) - pd.Timedelta(days=1)
        
        # Filtra o DataFrame já ajustado com base no vendedor
        df_month = df_filtrado[(df_filtrado['Data_Pedido'] >= month_start) & (df_filtrado['Data_Pedido'] <= month_end)]
        unique_clients = df_month['Matriz_Cliente'].nunique()
        results.append((month_start, unique_clients))

    children = []
    for i in range(1, 7):
        current_month, current_count = results[i]
        prev_month, prev_count = results[i - 1]

        percent_change = ((current_count / prev_count) * 100) if prev_count > 0 else 0

        month_col = dbc.Col([
            html.Div(current_month.strftime('%b').upper()[:3], className="month-name text-center", style={"fontSize": "14px"}),
            html.P("QTD. CLIENTES", className="info-text text-center", style={"fontSize": "12px"}),
            html.H6(f"{current_count}", className="info-number text-center", style={"fontSize": "16px"}),
            html.P(f"{percent_change:.2f}% POSITIVAÇÃO", className="info-percent text-center", style={"fontSize": "13px"})
        ], width=2)

        children.append(month_col)

    month_row = dbc.Row(children, className="mb-4")

    return [month_row]

# #################### Card tabela Cliente Sintético
# Função para preparar os dados do cliente sintético
def preparar_dados_cliente_sintetico(vendedor_selecionado, df, ano_selecionado, visualizacao='total'):
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]
    # Garantir que 'Cliente_ID_Nome' esteja criado corretamente
    df['Cliente_ID_Nome'] = df['Matriz_Cliente'].astype(str) + ' - ' + df['Cliente']

    # Filtrar por ano
    df_filtrado = df[df['Data_Pedido'].dt.year == ano_selecionado]

    # Agrupar e somar os valores
    if visualizacao == 'total':
        df_agrupado = df_filtrado.groupby(['Cliente_ID_Nome', 'Cidade', df_filtrado['Data_Pedido'].dt.strftime('%m/%Y')])['TOTAL'].sum().unstack(fill_value=0)
    else:
        df_agrupado = df_filtrado.groupby(['Cliente_ID_Nome', 'Cidade', df_filtrado['Data_Pedido'].dt.strftime('%m/%Y')])['m2'].sum().unstack(fill_value=0)

    # Resetar índice para tornar as colunas 'Cliente_ID_Nome' e 'Cidade' parte do DataFrame
    df_agrupado.reset_index(inplace=True)

    # Inserir todos os meses como colunas, mesmo se não houver dados
    todos_os_meses = pd.date_range(start=f'{ano_selecionado}-01-01', end=f'{ano_selecionado}-12-31', freq='MS').strftime('%m/%Y').tolist()
    for mes in todos_os_meses:
        if mes not in df_agrupado.columns:
            df_agrupado[mes] = 0

    # Reordenar as colunas se necessário
    colunas = ['Cliente_ID_Nome', 'Cidade'] + todos_os_meses
    df_agrupado = df_agrupado[colunas]

    # Renomear colunas para o DataTable
    df_agrupado.columns = ['Cliente', 'Cidade'] + todos_os_meses

    return df_agrupado

# Callback para atualizar a tabela "Cliente Sintético"
@app.callback(
    Output("tabela-cliente-sintetico", "children"),
    [
        Input("interval-update", "n_intervals"),
        Input("filtro_ano", "value"),
        Input('id-busca-input', 'value'),
        Input('filtro_visualizacao', 'value'),
        Input('vendedor-dropdown', 'value')
    ] 
)
def update_tabela_cliente_sintetico(n_intervals, ano_selecionado, id_busca, visualizacao, vendedor_selecionado):
    df_cliente_sintetico = preparar_dados_cliente_sintetico(vendedor_selecionado, df, ano_selecionado, visualizacao)

    if id_busca:
        id_busca_str = str(id_busca)
        mask = df_cliente_sintetico['Cliente'].str.contains(id_busca_str, case=False, na=False)

        if mask.any():
            df_cliente_sintetico = pd.concat([df_cliente_sintetico[mask], df_cliente_sintetico[~mask]])
        else:
            return dbc.Alert(f'ID {id_busca} não encontrado.', color='danger')

    # Format values based on visualization type
    if visualizacao == 'metragem':
        for mes in df_cliente_sintetico.columns[2:]:
            df_cliente_sintetico[mes] = df_cliente_sintetico[mes].apply(lambda x: f"{x:.2f} m²" if x != 0 else x)
    else:  # Faturamento
        for mes in df_cliente_sintetico.columns[2:]:
            df_cliente_sintetico[mes] = df_cliente_sintetico[mes].apply(lambda x: f"R$ {x:,.2f}" if x != 0 else x)

    columns = [
        {'name': 'Cliente', 'id': 'Cliente'}, 
        {'name': 'Cidade', 'id': 'Cidade'}
    ] + [
        {'name': mes, 'id': mes} 
        for mes in df_cliente_sintetico.columns[2:]
    ]

    tabela = dash_table.DataTable(
        data=df_cliente_sintetico.to_dict('records'),
        columns=columns,
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'center',
            'minWidth': '120px', 'width': '160px', 'maxWidth': '180px',
            'overflow': 'hidden',
            'textOverflow': 'ellipsis',
        },
        style_cell_conditional=[
            {'if': {'column_id': 'Cliente'}, 'textAlign': 'left', 'minWidth': '160px', 'width': '280px', 'maxWidth': '300px'},
            {'if': {'column_id': 'Cidade'}, 'textAlign': 'center', 'minWidth': '80px', 'width': '110px', 'maxWidth': '120px'},
        ],
        style_header={
            'fontWeight': 'bold',
            'textAlign': 'center'
        },
        page_size=10,
    )

    return tabela

def create_cliente_sintetico_card():
    return dbc.Card(
        [
            dbc.CardHeader(html.H3("Cliente Sintético", className="text-left")),
            dbc.CardBody(
                [
                    dbc.Row(
                        [
                            dbc.Col(
                                dcc.Dropdown(
                                    id='filtro_visualizacao',
                                    options=[
                                        {'label': 'Faturamento', 'value': 'total'},
                                        {'label': 'Metragem', 'value': 'metragem'}
                                    ],
                                    value='total',
                                    clearable=False,
                                    style={'width': '150px'}
                                ),
                                width=2
                            ),
                            dbc.Col(
                                dbc.Input(
                                    id='id-busca-input',
                                    type='number',
                                    placeholder='Buscar por ID do Cliente',
                                    style={'width': '150px'}
                                ),
                                width=2
                            ),
                            dbc.Col(
                                dcc.Dropdown(
                                    id='filtro_ano',
                                    options=[{'label': ano, 'value': ano} for ano in range(datetime.now().year-1, datetime.now().year + 1)],
                                    value=datetime.now().year,
                                    clearable=False,
                                    style={'width': '120px'}
                                ),
                                width=2
                            ),
                            dbc.Col(
                                html.Button("Exportar Excel", id="btn_exportar", n_clicks=0, className="btn btn-primary"),
                                width=2
                            ),
                        ]
                    ),
                    html.Div(id="tabela-cliente-sintetico"),
                ]
            ),
        ],
        style={"width": "100%"},
    )

@app.callback(
    Output('download-excel', 'data'),
    Input('btn_exportar', 'n_clicks'),
    State('filtro_ano', 'value'),  # Captura o valor do ano selecionado no filtro
    prevent_initial_call=True
)
def exportar_para_excel(n_clicks, ano_selecionado):
    if n_clicks > 0:
        # Usando a função modificada que inclui todos os dados e passando o ano selecionado como argumento
        df_cliente_sintetico = preparar_dados_cliente_sintetico("TODOS OS VENDEDORES", df, ano_selecionado, 'total')

        # Criar um buffer de bytes em memória
        output = BytesIO()

        # Usar o Pandas para escrever o DataFrame no buffer como um Excel
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_cliente_sintetico.to_excel(writer, index=False)

        # Mover o cursor para o começo do arquivo
        output.seek(0)

        # Enviar o buffer para download
        return dcc.send_bytes(output.getvalue(), filename=f"clientes_sinteticos_{ano_selecionado}.xlsx")

    return None

# Pegar o os dados para os cards
valor_realizado = calc_realizado(df)
vendedor_names = get_vendedor_names(df)  
valor_projetado = calc_projecao_geral(valor_realizado)
cliente_sintetico_card = create_cliente_sintetico_card()
categorias_agregadas = ['ACESSÓRIOS', 'ALUMÍNIO', 'FERRAGEM', 'KIT PARA BOX PADRÃO', 'SILICONE']
categoria_vidro = ['VIDRO']

# Preparando os dados
df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y')
hoje = datetime.now()
meses = []

# Calculando os limites dos últimos três meses
for i in range(3):
    inicio_mes = (hoje - timedelta(days=hoje.day - 1) - pd.offsets.MonthBegin(n=i)).to_pydatetime()
    fim_mes = (inicio_mes + pd.offsets.MonthEnd(n=1)).to_pydatetime()
    meses.append((inicio_mes, fim_mes))

meses.reverse()
vendedor_padrao = "TODOS OS VENDEDORES"

# Calculando as somas para cada mês e categoria com o valor padrão para vendedor
somas_agregadas = [calcular_somas(df, categorias_agregadas, inicio, fim, vendedor_padrao) for inicio, fim in meses]
somas_vidro = [calcular_somas(df, categoria_vidro, inicio, fim, vendedor_padrao) for inicio, fim in meses]

#nomes_meses = [inicio.strftime('%B') for inicio, _ in meses]
nomes_meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
               'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

# Criando o gráfico de pilha empilhadas
fig_pilha = go.Figure(data=[
    go.Bar(name='Vidro', x=nomes_meses, y=somas_vidro, marker_color='blue'),
    go.Bar(name='Agregadas', x=nomes_meses, y=somas_agregadas, marker_color='green')
])

# Alterar a disposição do gráfico para empilhar
fig_pilha.update_layout(barmode='stack', title='Vendas por Categoria nos Últimos 3 Meses')

df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y')

##################### Card tabela faturamento vidro 3 meses
def calcular_somas_grupos(df, subcategoria):
    # Obtém o primeiro e o último dia do mês atual
    hoje = pd.to_datetime('today').normalize()
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = primeiro_dia_mes + pd.offsets.MonthEnd(1)

    # Filtra o DataFrame para incluir apenas as linhas do mês atual e da categoria especificada
    df_filtrado = df[(df['Data_Pedido'] >= primeiro_dia_mes) & 
                     (df['Data_Pedido'] <= ultimo_dia_mes) & 
                     (df['Tipo_Produto'] == subcategoria)]

    # Aplica o desconto antes de somar
    df_filtrado['total_produto_com_desconto'] = df_filtrado.apply(apply_discount, axis=1)

    # Retorna a soma dos totais com desconto
    return df_filtrado['total_produto_com_desconto'].sum()

@app.callback(
    Output('faturamento_vidro_card_container', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_faturamento_vidro_card(vendedor_selecionado):
    if not vendedor_selecionado:
        # Se por algum motivo o vendedor_selecionado for None, use o valor padrão
        vendedor_selecionado = "TODOS OS VENDEDORES"
    
    return create_faturamento_vidro_card(df, vendedor_selecionado)

def create_faturamento_vidro_table(df, vendedor_selecionado):
    subcategorias = {
        'TEMPERADO ENGENHARIA': ['ENGENHARIA TEMPERADO', 'BOX ENGENHARIA'],
        'TEMPERADO PRONTA ENTREGA': ['BOX PADRÃO', 'JANELA PADRÃO', 'PORTA PIVOTANTE'],
        'COMUM CORTADO': ['CORTADO ESPELHO', 'CORTADO FLOAT', 'CORTADO LAMINADO', 'CORTADO FANTASIA', 'CORTADO REFLETIVO BRONZE', 'CORTADO SERIGRAFADO'],
        'COMUM CHAPARIA': ['CHAPARIA ESPELHO', 'CHAPARIA FANTASIA', 'CHAPARIA FLOAT', 'CHAPARIA LAMINADO', 'CHAPARIA REFLETIVO BRONZE', 'CHAPARIA SERIGRAFADO'],
    }

    # Filtrar por vendedor, se não for "TODOS OS VENDEDORES"
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]

    hoje = pd.to_datetime('today').normalize()
    start_dates = [(hoje - pd.offsets.MonthBegin(n=i+1)).replace(day=1) for i in range(3, 0, -1)]

    data = []
    for grupo, tipos in subcategorias.items():
        grupo_row = {'Subcategoria': grupo}
        grupo_somas = {start_date.strftime('%b/%Y'): 0 for start_date in start_dates}

        for tipo in tipos:
            tipo_row = {'Subcategoria': f"· {tipo}"}
            faturamento_mes = {}
            for start_date in start_dates:
                fim_mes = start_date + pd.offsets.MonthEnd(1)

                df_filtrado = df[(df['Data_Pedido'] >= start_date) & (df['Data_Pedido'] <= fim_mes) & (df['Tipo_Produto'] == tipo)]
                
                df_filtrado['Faturamento_com_Desconto'] = df_filtrado.apply(apply_discount, axis=1)

                faturamento = df_filtrado['Faturamento_com_Desconto'].sum()
                
                tipo_row[start_date.strftime('%b/%Y')] = f"R$ {faturamento:,.2f}"
                faturamento_mes[start_date.strftime('%b/%Y')] = faturamento
                grupo_somas[start_date.strftime('%b/%Y')] += faturamento

            # Determina o TOP MÊS para a subcategoria
            top_mes_faturamento = max(faturamento_mes.values())
            top_mes_subcategoria = [mes for mes, faturamento in faturamento_mes.items() if faturamento == top_mes_faturamento][0]
            tipo_row['TOP MÊS'] = top_mes_subcategoria
            
            data.append(tipo_row)
        
        # Adiciona faturamento total e TOP MÊS ao grupo
        for mes, soma in grupo_somas.items():
            grupo_row[mes] = f"R$ {soma:,.2f}"
        top_mes_grupo_faturamento = max(grupo_somas.values())
        grupo_row['TOP MÊS'] = [mes for mes, faturamento in grupo_somas.items() if faturamento == top_mes_grupo_faturamento][0]
        data.insert(len(data) - len(tipos), grupo_row)

    # Adicionar a categoria "Outros Vidros" ao final
    row_outros_vidros = {'Subcategoria': 'OUTROS VIDROS'}
    soma_outros_vidros = {}
    for start_date in start_dates:
        fim_mes = start_date + pd.offsets.MonthEnd(1)
        df_filtrado = df[(df['Data_Pedido'] >= start_date) & (df['Data_Pedido'] <= fim_mes) & (~df['Tipo_Produto'].isin(sum(subcategorias.values(), []))) & (df['Grupo'] == 'VIDRO')]
        df_filtrado['Faturamento_com_Desconto_'] = df_filtrado.apply(apply_discount, axis=1)
        soma_grupo = df_filtrado['Faturamento_com_Desconto_'].sum()
        row_outros_vidros[start_date.strftime('%b/%Y')] = f"R$ {soma_grupo:,.2f}"
        soma_outros_vidros[start_date.strftime('%b/%Y')] = soma_grupo
    
    # Determinar o TOP MÊS para "OUTROS VIDROS"
    top_mes_outros_vidros = max(soma_outros_vidros, key=soma_outros_vidros.get)
    row_outros_vidros['TOP MÊS'] = top_mes_outros_vidros

    data.append(row_outros_vidros)

    columns = [
        {"name": "Subcategoria", "id": "Subcategoria", "type": "text"}
    ] + [
        {"name": mes.strftime('%b/%Y'), "id": mes.strftime('%b/%Y'), "type": "numeric", "format": Format(symbol=Symbol.yes, symbol_suffix='R$ ', scheme=Scheme.fixed)}
        for mes in start_dates
    ] + [
        {"name": "TOP MÊS", "id": "TOP MÊS", "type": "text"}
    ]

    return dash_table.DataTable(
        data=data,
        columns=columns,
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'left', 'padding': '5px', 'width': '210px'},
        style_header={'fontWeight': 'bold', 'textAlign': 'center', 'width': '210px'},
        style_cell_conditional=
        [{
            'if': {'column_id': 'Subcategoria'},
            'textAlign': 'left',
            'width': '210px',
            'whiteSpace': 'normal'
        }]
    )

# Função para criar o card da tabela
def create_faturamento_vidro_card(df, vendedor_selecionado):
    return dbc.Card(
        [
            dbc.CardHeader(html.H3("Faturamento dos Últimos 3 Meses de Vidro")),
            dbc.CardBody(create_faturamento_vidro_table(df, vendedor_selecionado)),
            dbc.CardBody(id="faturamento_vidro_3m")
        ]
    )

def format_to_currency(value):
    try:
        # Converte para float e formata como moeda
        numeric_value = float(value)
    except ValueError:
        return "R$ 0,00"  # Retorna um valor padrão para valores inválidos
    # Formata o número com separador de milhar como ponto e separador decimal como vírgula
    return f"R$ {numeric_value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def calcular_volume_por_categoria(df, categoria):
    # Obtém o primeiro e o último dia do mês atual
    hoje = pd.to_datetime('today').normalize()
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = primeiro_dia_mes + pd.offsets.MonthEnd(1)

    # Filtra o DataFrame para incluir apenas as linhas do mês atual e da categoria especificada
    df_filtrado = df[(df['Data_Pedido'] >= primeiro_dia_mes) & 
                     (df['Data_Pedido'] <= ultimo_dia_mes) & 
                     (df['Tipo_Produto'] == categoria)]

    # Calcula e retorna a soma da coluna 'm2' para as linhas filtradas
    return df_filtrado['m2'].sum()

def calc_projecao_categoria(realizado):
    hoje = pd.to_datetime('today')
    ontem = hoje - pd.offsets.BDay(1)  # Calcula o dia útil anterior
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = hoje + pd.offsets.MonthEnd(1)
    
    dias_uteis_ate_ontem = pd.bdate_range(start=primeiro_dia_mes, end=ontem).size
    total_dias_uteis_mes = pd.bdate_range(start=primeiro_dia_mes, end=ultimo_dia_mes).size

    if dias_uteis_ate_ontem > 0:
        projecao = (realizado / dias_uteis_ate_ontem) * total_dias_uteis_mes
        return projecao
    return 0

def calcular_media_faturamento_ultimos_3_meses(df, subcategorias):
    def calcular_somas_grupos_(df, subcategoria, inicio_mes, fim_mes):
        df_filtrado = df[(df['Data_Pedido'] >= inicio_mes) & (df['Data_Pedido'] <= fim_mes) & (df['Tipo_Produto'].isin([subcategoria]))]
        df_filtrado['total_produto_com_desconto'] = df_filtrado.apply(apply_discount, axis=1)
        return df_filtrado['total_produto_com_desconto'].sum()
    
    # Calcula a média do faturamento dos últimos 3 meses
    hoje = pd.to_datetime('today').normalize()
    start_dates = [(hoje - pd.offsets.MonthBegin(n=i+1)).replace(day=1) for i in range(3, 0, -1)]
    faturamento_total = sum(
        calcular_somas_grupos_(df, tipo, start_date, start_date + pd.offsets.MonthEnd(1))
        for start_date in start_dates
        for tipo in subcategorias
    ) 
    return (faturamento_total / 3)

@app.callback(
    Output('categoria_vidro_table_container', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_categoria_vidro_table(vendedor_selecionado):
    if not vendedor_selecionado:
        vendedor_selecionado = "TODOS OS VENDEDORES"
    
    # Agora você chama a função que cria a tabela, passando o DataFrame e o vendedor selecionado
    return create_categoria_vidro_card(df, vendedor_selecionado)

def create_categoria_vidro_card(df, vendedor_selecionado):
    return dbc.Card(
        [
            dbc.CardHeader(html.H3("Categoria Vidro")),
            dbc.CardBody(create_categoria_vidro_table(df, vendedor_selecionado))
        ]
    )

def create_categoria_vidro_table(df, vendedor_selecionado):
    subcategorias = {
        'TEMPERADO ENGENHARIA': ['ENGENHARIA TEMPERADO', 'BOX ENGENHARIA'],
        'TEMPERADO PRONTA ENTREGA': ['BOX PADRÃO', 'JANELA PADRÃO', 'PORTA PIVOTANTE'],
        'COMUM CORTADO': ['CORTADO ESPELHO', 'CORTADO FLOAT', 'CORTADO LAMINADO', 'CORTADO FANTASIA', 'CORTADO REFLETIVO BRONZE', 'CORTADO SERIGRAFADO'],
        'COMUM CHAPARIA': ['CHAPARIA ESPELHO', 'CHAPARIA FANTASIA', 'CHAPARIA FLOAT', 'CHAPARIA LAMINADO', 'CHAPARIA REFLETIVO BRONZE', 'CHAPARIA SERIGRAFADO'],
    }

    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]

    # Preparação da lista de dados e somatórios de grupo
    data = []
    for grupo, tipos in subcategorias.items():
        grupo_data = {'Subcategoria': grupo, 'Volume': 0, 'Realizado': 0, 'Projeção': 0, 'Meta': 0}
        grupo_total_realizado = 0
        grupo_total_meta = 0

        for tipo in tipos:
            volume = calcular_volume_por_categoria(df, tipo)
            realizado = calcular_somas_grupos(df, tipo)
            projecao = calc_projecao_categoria(realizado)
            meta = calcular_media_faturamento_ultimos_3_meses(df, [tipo])

            # Atualiza somatórios do grupo
            grupo_data['Volume'] += volume
            grupo_data['Realizado'] += realizado
            grupo_data['Projeção'] += projecao
            grupo_data['Meta'] += meta
            grupo_total_realizado += realizado
            grupo_total_meta += meta

            # Insere dados da subcategoria
            data.append({
                'Subcategoria': f"· {tipo}",
                'Volume': f"{volume:.2f} m²",
                'Realizado': "R$ {:,.2f}".format(realizado).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Projeção': "R$ {:,.2f}".format(projecao).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Meta': "R$ {:,.2f}".format(meta).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Projeção vs Meta': f"{(projecao / meta * 100) if meta else 0:.2f}%",
            })

        # Calcula e adiciona projeção vs meta para o grupo
        grupo_data['Projeção vs Meta'] = f"{(grupo_data['Projeção'] / grupo_data['Meta'] * 100) if grupo_data['Meta'] else 0:.2f}%"

        # Formata valores para o grupo
        grupo_data['Volume'] = f"{grupo_data['Volume']:.2f} m²"
        grupo_data['Realizado'] = "R$ {:,.2f}".format(grupo_data['Realizado']).replace(',', 'X').replace('.', ',').replace('X', '.')
        grupo_data['Projeção'] = "R$ {:,.2f}".format(grupo_data['Projeção']).replace(',', 'X').replace('.', ',').replace('X', '.')
        grupo_data['Meta'] = "R$ {:,.2f}".format(grupo_data['Meta']).replace(',', 'X').replace('.', ',').replace('X', '.')

        # Insere os dados agregados do grupo no início de sua seção
        data.insert(len(data) - len(tipos), grupo_data)
    df_outros_vidros = df[~df['Tipo_Produto'].isin(sum(subcategorias.values(), [])) & (df['Grupo'] == 'VIDRO')]
    # Adapte as funções de cálculo se necessário para aceitar um df como argumento

    volume_outros = calcular_volume_por_categoria(df_outros_vidros, "OUTROS VIDROS")
    realizado_outros = calcular_somas_grupos(df_outros_vidros, "OUTROS VIDROS")
    projecao_outros = calc_projecao_categoria(realizado_outros)
    meta_outros = calcular_media_faturamento_ultimos_3_meses(df_outros_vidros, ["OUTROS VIDROS"])
    projecao_vs_meta_outros = f"{(projecao_outros / meta_outros) * 100:.2f}%" if meta_outros > 0 else " Meta não definida"

    # Adicionando "OUTROS VIDROS" aos dados
    data.append({
        'Subcategoria': 'OUTROS VIDROS',
        'Volume': f"{volume_outros:,.2f} m²",
        'Realizado': "R$ {:,.2f}".format(realizado_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção': "R$ {:,.2f}".format(projecao_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Meta': "R$ {:,.2f}".format(meta_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção vs Meta': projecao_vs_meta_outros
    })


    return dash_table.DataTable(
        id='categoria-vidro-table',
        columns=[
            {'name': 'Subcategoria', 'id': 'Subcategoria'},
            {'name': 'Volume', 'id': 'Volume'},
            {'name': 'Realizado', 'id': 'Realizado'},
            {'name': 'Projeção', 'id': 'Projeção'},
            {'name': 'Meta', 'id': 'Meta', 'type': 'numeric', 'format': Format(symbol=Symbol.yes, symbol_suffix='R$ ', scheme=Scheme.fixed)},
            {'name': 'Projeção vs Meta', 'id': 'Projeção vs Meta'}
        ],
        data=data,
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'left', 'padding': '5px'},
        style_header={'fontWeight': 'bold', 'textAlign': 'center'},
        style_data_conditional=[
            {
                'if': {'filter_query': '{Subcategoria} contains "·"'},
                'paddingLeft': '10px',  
            },
            {
                'if': {'filter_query': '{Subcategoria} eq ""'},
                'display': 'none' 
            }
        ]
    )

################### TABELA FATURAMENTO DOS ÚLTIMOS 3 MESES DE AGREGADOS
@app.callback(
    Output('faturamento_agregados_3m_container', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_faturamento_agregados_table(vendedor_selecionado):
    if not vendedor_selecionado:
        vendedor_selecionado = "TODOS OS VENDEDORES"
    
    # Chama a função que cria a tabela, passando o DataFrame e o vendedor selecionado
    return create_categoria_vidro_agregados_card(df, vendedor_selecionado)

def create_categoria_vidro_agregados_card(df, vendedor_selecionado):
    return dbc.Card(
        [
            dbc.CardHeader(html.H3("Faturamento para os Últimos 3 Meses de Agregados")),
            dbc.CardBody(create_faturamento_agregados_table(df, vendedor_selecionado))
        ]
    )

def create_faturamento_agregados_table(df, vendedor_selecionado):
    subcategorias = {
        'KIT\'S PARA BOX': ['KIT BOX COMPLETO AL', 'KIT BOX COMPLETO IDEIA GLASS', 'KIT BOX COMPLETO IMPORTADO', 'KIT BOX COMPLETO PORTAL', 'KIT BOX COMPLETO PORTAL - AVARIA'],
        'KIT\'S PARA JANELA': ['KIT JANELA COMPLETA WD', 'KIT JANELA COMPLETO PORTAL'],
        'PERFIS PARA VIDRO TEMPERADO': ['PERFIS ENGENHARIA AL', 'PERFIS ENGENHARIA PERFILEVE', 'PERFIS ENGENHARIA PERFILEVE 3MTS'],
        'FERRAGENS': ['KIT FERRAGENS LGL', 'FERRAGENS LGL', 'MOLAS', 'ROLDANAS', 'PUXADORES'],
        'SERVIÇOS': ['MÃO DE OBRA', 'BENEFICIAMENTO', 'FRETE', 'CAIXA DE MADEIRA'],
        'OUTROS': ['SILICONE', 'FIXA ESPELHO', 'SUPORTES', 'BORRACHAS', 'ESCOVINHAS'],
    }
    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]

    hoje = pd.to_datetime('today').normalize()
    start_dates = [(hoje - pd.offsets.MonthBegin(n=i+1)).replace(day=1) for i in range(3, 0, -1)]

    data = []
    for grupo, tipos in subcategorias.items():
        grupo_row = {'Subcategoria': grupo}
        grupo_somas = {start_date.strftime('%b/%Y'): 0 for start_date in start_dates}

        for tipo in tipos:
            tipo_row = {'Subcategoria': f"· {tipo}"}
            faturamento_mes = {}
            for start_date in start_dates:
                fim_mes = start_date + pd.offsets.MonthEnd(1)
                df_filtrado = df[(df['Data_Pedido'] >= start_date) & (df['Data_Pedido'] <= fim_mes) & (df['Tipo_Produto'] == tipo)]
                
                df_filtrado['Faturamento_com_Desconto'] = df_filtrado.apply(apply_discount, axis=1)

                faturamento = df_filtrado['Faturamento_com_Desconto'].sum()
                
                tipo_row[start_date.strftime('%b/%Y')] = f"R$ {faturamento:,.2f}"
                faturamento_mes[start_date.strftime('%b/%Y')] = faturamento
                grupo_somas[start_date.strftime('%b/%Y')] += faturamento

            # Determina o TOP MÊS para a subcategoria
            top_mes_faturamento = max(faturamento_mes.values())
            top_mes_subcategoria = [mes for mes, faturamento in faturamento_mes.items() if faturamento == top_mes_faturamento][0]
            tipo_row['TOP MÊS'] = top_mes_subcategoria
            
            data.append(tipo_row)
        
        # Adiciona faturamento total e TOP MÊS ao grupo
        for mes, soma in grupo_somas.items():
            grupo_row[mes] = f"R$ {soma:,.2f}"
        top_mes_grupo_faturamento = max(grupo_somas.values())
        grupo_row['TOP MÊS'] = [mes for mes, faturamento in grupo_somas.items() if faturamento == top_mes_grupo_faturamento][0]
        data.insert(len(data) - len(tipos), grupo_row)

    columns = [
        {"name": "Subcategoria", "id": "Subcategoria", "type": "text"}
    ] + [
        {"name": mes.strftime('%b/%Y'), "id": mes.strftime('%b/%Y'), "type": "numeric", "format": Format(symbol=Symbol.yes, symbol_suffix='R$ ', scheme=Scheme.fixed)}
        for mes in start_dates
    ] + [
        {"name": "TOP MÊS", "id": "TOP MÊS", "type": "text"}
    ]

    return html.Div([
        dash_table.DataTable(
            data=data,
            columns=columns,
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'left', 'padding': '5px', 'width': '210px'},
            style_header={'fontWeight': 'bold', 'textAlign': 'left', 'width': '210px'},
            style_cell_conditional=
            [{
                'if': {'column_id': 'Subcategoria'},
                'textAlign': 'left',
                'width': '210px',
                'whiteSpace': 'normal'
            }]
        )
    ])

################### TABELA CATEGORIA AGREGADOS
@app.callback(
    Output('categoria_agregados_table_container', 'children'),
    [Input('vendedor-dropdown', 'value')]
)
def update_categoria_agregados_table(vendedor_selecionado):
    if not vendedor_selecionado:
        vendedor_selecionado = "TODOS OS VENDEDORES"
    
    # Chama a função que cria a tabela, passando o DataFrame e o vendedor selecionado
    return create_categoria_agregados_card(df, vendedor_selecionado)

def create_categoria_agregados_card(df, vendedor_selecionado):
    return dbc.Card(
        [
            dbc.CardHeader(html.H3("Categoria Agregados")),
            dbc.CardBody(create_categoria_agregadas_table(df, vendedor_selecionado))
        ]
    )

def create_categoria_agregadas_table(df, vendedor_selecionado):
    subcategorias = {
        'KIT\'S PARA BOX': ['KIT BOX COMPLETO AL', 'KIT BOX COMPLETO IDEIA GLASS', 'KIT BOX COMPLETO IMPORTADO', 'KIT BOX COMPLETO PORTAL', 'KIT BOX COMPLETO PORTAL - AVARIA'],
        'KIT\'S PARA JANELA': ['KIT JANELA COMPLETA WD', 'KIT JANELA COMPLETO PORTAL'],
        'PERFIS PARA VIDRO TEMPERADO': ['PERFIS ENGENHARIA AL', 'PERFIS ENGENHARIA PERFILEVE', 'PERFIS ENGENHARIA PERFILEVE 3MTS'],
        'FERRAGENS': ['KIT FERRAGENS LGL', 'FERRAGENS LGL', 'MOLAS', 'ROLDANAS', 'PUXADORES'],
        'SERVIÇOS': ['MÃO DE OBRA', 'BENEFICIAMENTO', 'FRETE', 'CAIXA DE MADEIRA'],
        'OUTROS': ['SILICONE', 'FIXA ESPELHO', 'SUPORTES', 'BORRACHAS', 'ESCOVINHAS'],
    }

    if vendedor_selecionado != "TODOS OS VENDEDORES":
        df = df[df['Vendedor'] == vendedor_selecionado]

    # Preparação da lista de dados e somatórios de grupo
    data = []
    for grupo, tipos in subcategorias.items():
        grupo_data = {'Subcategoria': grupo, 'Volume': 0, 'Realizado': 0, 'Projeção': 0, 'Meta': 0}
        grupo_total_realizado = 0
        grupo_total_meta = 0

        for tipo in tipos:
            realizado = calcular_somas_grupos(df, tipo)
            projecao = calc_projecao_categoria(realizado)
            meta = calcular_media_faturamento_ultimos_3_meses(df, [tipo])

            # Atualiza somatórios do grupo
            grupo_data['Realizado'] += realizado
            grupo_data['Projeção'] += projecao
            grupo_data['Meta'] += meta
            grupo_total_realizado += realizado
            grupo_total_meta += meta

            # Insere dados da subcategoria
            data.append({
                'Subcategoria': f"· {tipo}",
                'Realizado': "R$ {:,.2f}".format(realizado).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Projeção': "R$ {:,.2f}".format(projecao).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Meta': "R$ {:,.2f}".format(meta).replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Projeção vs Meta': f"{(projecao / meta * 100) if meta else 0:.2f}%",
            })

        # Calcula e adiciona projeção vs meta para o grupo
        grupo_data['Projeção vs Meta'] = f"{(grupo_data['Projeção'] / grupo_data['Meta'] * 100) if grupo_data['Meta'] else 0:.2f}%"

        # Formata valores para o grupo
        grupo_data['Realizado'] = "R$ {:,.2f}".format(grupo_data['Realizado']).replace(',', 'X').replace('.', ',').replace('X', '.')
        grupo_data['Projeção'] = "R$ {:,.2f}".format(grupo_data['Projeção']).replace(',', 'X').replace('.', ',').replace('X', '.')
        grupo_data['Meta'] = "R$ {:,.2f}".format(grupo_data['Meta']).replace(',', 'X').replace('.', ',').replace('X', '.')

        # Insere os dados agregados do grupo no início de sua seção
        data.insert(len(data) - len(tipos), grupo_data)
    df_outros_agregados = df[~df['Tipo_Produto'].isin(sum(subcategorias.values(), [])) & (df['Grupo'] == 'AGREGADOS')]
    # Adapte as funções de cálculo se necessário para aceitar um df como argumento

    realizado_outros = calcular_somas_grupos(df_outros_agregados, "OUTROS AGREGADOS")
    projecao_outros = calc_projecao_categoria(realizado_outros)
    meta_outros = calcular_media_faturamento_ultimos_3_meses(df_outros_agregados, ["OUTROS AGREGADOS"])
    projecao_vs_meta_outros = f"{(projecao_outros / meta_outros) * 100:.2f}%" if meta_outros > 0 else " Meta não definida"

    # Adicionando "OUTROS AGREGADOS" aos dados
    data.append({
        'Subcategoria': 'OUTROS AGREGADOS',
        'Realizado': "R$ {:,.2f}".format(realizado_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção': "R$ {:,.2f}".format(projecao_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Meta': "R$ {:,.2f}".format(meta_outros).replace(',', 'X').replace('.', ',').replace('X', '.'),
        'Projeção vs Meta': projecao_vs_meta_outros
    })

    return dash_table.DataTable(
        id='categoria-agregadas-table',
        columns=[
            {'name': 'Subcategoria', 'id': 'Subcategoria'},
            {'name': 'Realizado', 'id': 'Realizado'},
            {'name': 'Projeção', 'id': 'Projeção'},
            {'name': 'Meta', 'id': 'Meta', 'type': 'numeric', 'format': Format(symbol=Symbol.yes, symbol_suffix='R$ ', scheme=Scheme.fixed)},
            {'name': 'Projeção vs Meta', 'id': 'Projeção vs Meta'}
        ],
        data=data,
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'left', 'padding': '5px'},
        style_header={'fontWeight': 'bold', 'textAlign': 'center'},
        style_data_conditional=[
            {
                'if': {'filter_query': '{Subcategoria} contains "·"'},
                'paddingLeft': '10px',  
            },
            {
                'if': {'filter_query': '{Subcategoria} eq ""'},
                'display': 'none' 
            }
        ]
    )

########## CARD PONTUAÇÃO VENDEDOR DESTAQUE
def calcular_pontuacao(porcentagem):
    if porcentagem < 80:
        return 100
    elif porcentagem < 90:
        return 200
    elif porcentagem < 100:
        return 300
    else:
        return 400

# Função auxiliar para carregar imagens e converter para o formato adequado para uso no Dash
def encode_image(image_file):
    encoded = base64.b64encode(open(image_file, 'rb').read())
    return f'data:image/png;base64,{encoded.decode()}'

# Função para escolher a imagem com base na porcentagem
def image_for_percentage(percentage):
    if percentage >= 100:
        return encode_image('assets/happy.png')
    elif percentage >= 50:
        return encode_image('assets/moderate.png')
    else:
        return encode_image('assets/sad.png')

@app.callback(
    Output("pontuacao-vendedor-destaque", "children"), 
    [Input("vendedor-dropdown", "value")]
)

def atualizar_pontuacao_vendedor(vendedor_selecionado):
    if vendedor_selecionado == 'TODOS OS VENDEDORES':
        return "Selecione um vendedor"

    temp_path = 'META_VENDEDORES.xlsx'
    df_meta_vendedores = pd.read_excel(temp_path)

    df['Data_Pedido'] = pd.to_datetime(df['Data_Pedido'], format='%d/%m/%Y', errors='coerce')
    
    
    meta_vendedor = df_meta_vendedores[df_meta_vendedores['NOME VENDEDOR'] == vendedor_selecionado]
    if not meta_vendedor.empty:
        meta_vidro = meta_vendedor['META VIDRO'].values[0]
        meta_agregado = meta_vendedor['META AGREGADOS'].values[0]
        meta_geral = meta_vendedor['META VENDEDOR'].values[0]

        def projecao_vidro():
            df_ = df
            df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')

            # Filtrando o DataFrame pelo vendedor selecionado, pelo período, e pela categoria específica
            df_filtrado = df_[(df_['Vendedor'] == vendedor_selecionado) & 
                            (df_['Data_Pedido'] >= start_date_realizado) & 
                            (df_['Data_Pedido'] <= end_date_realizado) & 
                            (df_['Grupo'] == "VIDRO")]
            df_filtrado_unico = df_filtrado.drop_duplicates(subset='Id_Pedido', keep='first')

            # Calculando o valor realizado pelo vendedor selecionado para a categoria específica
            valor_realizado_vendedor = df_filtrado_unico['TOTAL'].sum()

            # Calculando os dias úteis até ontem e o total de dias úteis no mês
            dias_uteis_ate_ontem_ = dias_uteis_ate_ontem(start_date_realizado, end_date_realizado - pd.Timedelta(days=1))
            total_dias_uteis = total_dias_uteis_no_mes(start_date_realizado)

            # Calculando a projeção para a categoria específica
            if dias_uteis_ate_ontem_ > 0:  # Evitar divisão por zero
                projecao = (valor_realizado_vendedor / dias_uteis_ate_ontem_) * total_dias_uteis
            else: projecao = 0
            return projecao
        def projecao_agregado():
            agregados = ['ACESSÓRIOS', 'ALUMÍNIO', 'FERRAGEM', 'KIT PARA BOX PADRÃO', 'SILICONE']
            df_ = df
            df_['Data_Pedido'] = pd.to_datetime(df_['Data_Pedido'], format='%d/%m/%Y', errors='coerce')

            # Filtrando o DataFrame pelo vendedor selecionado, pelo período, e pela categoria específica
            df_filtrado = df_[(df_['Vendedor'] == vendedor_selecionado) & 
                            (df_['Data_Pedido'] >= start_date_realizado) & 
                            (df_['Data_Pedido'] <= end_date_realizado) & 
                            (df_['Grupo'].isin(agregados))]
            
            df_filtrado_unico = df_filtrado.drop_duplicates(subset='Id_Pedido', keep='first')

            # Calculando o valor realizado pelo vendedor selecionado para a categoria específica
            valor_realizado_vendedor = df_filtrado_unico['TOTAL'].sum()

            # Calculando os dias úteis até ontem e o total de dias úteis no mês
            dias_uteis_ate_ontem_ = dias_uteis_ate_ontem(start_date_realizado, end_date_realizado - pd.Timedelta(days=1))
            total_dias_uteis = total_dias_uteis_no_mes(start_date_realizado)

            # Calculando a projeção para a categoria específica
            if dias_uteis_ate_ontem_ > 0:  # Evitar divisão por zero
                projecao = (valor_realizado_vendedor / dias_uteis_ate_ontem_) * total_dias_uteis
            else: projecao = 0
            return projecao

        projecao_vidro_ = projecao_vidro()
        projecao_agregado_ = projecao_agregado()
        projecao_ = calc_projecao_vendedor(df, vendedor_selecionado)

        # Calcule a porcentagem para cada categoria
        porcentagem_vidro = (projecao_vidro_ / meta_vidro) * 100 if meta_vidro > 0 else 0
        porcentagem_agregado = (projecao_agregado_ / meta_agregado) * 100 if meta_agregado > 0 else 0
        porcentagem_vendedor = (projecao_ / meta_geral) * 100 if meta_geral > 0 else 0
        
        # Calcule a pontuação para cada categoria
        pontuacao_vidro = calcular_pontuacao(porcentagem_vidro)
        pontuacao_agregado = calcular_pontuacao(porcentagem_agregado)
        pontuacao_geral = calcular_pontuacao(porcentagem_vendedor)

        image_vidro = image_for_percentage(porcentagem_vidro)
        image_agregado = image_for_percentage(porcentagem_agregado)
        image_geral = image_for_percentage(porcentagem_vendedor)

        # Construa e retorne o conteúdo do card de pontuação
        return [
        html.Div([
            html.P(f"GERAL {porcentagem_vendedor:.2f}% {pontuacao_geral} PTS"),
            html.Img(src=image_geral, style={'height': '50px', 'width': '50px'}),
        ]),
        html.Div([
            html.P(f"VIDRO {porcentagem_vidro:.2f}% {pontuacao_vidro} PTS"),
            html.Img(src=image_vidro, style={'height': '50px', 'width': '50px'}),
        ]),
        html.Div([
            html.P(f"AGREGADO {porcentagem_agregado:.2f}% {pontuacao_agregado} PTS"),
            html.Img(src=image_agregado, style={'height': '50px', 'width': '50px'}),
        ]),
    ] 
    else:
        print("Nenhum vendedor correspondente encontrado.") 

########## LAYOUT DASH
app.layout = dbc.Container([
    dcc.Interval(
        id='interval-update', 
        interval=300*1000,  # 5 minutos
        n_intervals=0
    ),
    dcc.Store(id='meta-value-store'),
    dbc.Row([
        dbc.Col(html.H1("FAROL DE VENDAS", className="text-center"), width=12)
    ]),
    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardBody([html.H5("Filtro Vendedor", className="card-title"),
            dcc.Dropdown(
                id='vendedor-dropdown', 
                options=[{'label': 'TODOS OS VENDEDORES', 'value': 'TODOS OS VENDEDORES'}] + get_vendedor_names(df),
                value='TODOS OS VENDEDORES',
                clearable=False,
            )])]),
            width={"size": 2, "offset": 0}, className="mb-4")]),
    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardBody([
                html.H5("META GERAL", className="card-title"),
                    dbc.InputGroup([
                        dbc.InputGroupText("R$", style={'height': '38px'}),
                        dbc.Input(id="meta-geral-input", type="text", placeholder="Digite a Meta Geral", n_submit=0, style={'height': '38px'}),
            ]),
        ])]), width=2),
        dbc.Col(dbc.Card([dbc.CardBody([
                            html.H5("REALIZADO", className="card-title"),
                            html.P(f"R$ {valor_realizado:,.2f}", className="card-text")
        ])]), width=2),
        dbc.Col(dbc.Card([dbc.CardBody([
                            html.H5("PROJEÇÃO GERAL", className="card-title"),
                            html.P(f"R$ {valor_projetado:,.2f}", className="card-text")
        ])]), width=2),
        dbc.Col(dbc.Card([dbc.CardBody([
                            html.H5("ESCALONAMENTO", className="card-title"),
                            html.P("Aguardando valor da meta...", className="card-text"),
                        ])
                    ], id='escalation-card'), width=2),
        dbc.Col(dbc.Card(dbc.CardBody([
            html.H5("PONTUAÇÃO VENDEDOR DESTAQUE", className="card-title"),
            html.Div(id="pontuacao-vendedor-destaque"),
        ]), id='card-pontuacao-vendedor-destaque'), width=4)

    ], className="mb-4"),
        dbc.Row([
            dbc.Col(dbc.Card(dbc.CardBody([
            dcc.Graph(
                id='stack-chart',
                figure=fig_pilha,
                style={'height': '100%', 'width': '100%'},  
                config={'responsive': True})]),
                style={'height': '400px'}  
            ),width=4, className="mb-4"
        ),
                   
    dbc.Col([
        dbc.Row(dbc.Card([dbc.CardBody([
            html.H5("META POR VENDEDOR", className="card-title"),
            dbc.Input(id="meta-vendedor-input", type="text", placeholder="Digite a Meta do Vendedor e pressione Enter", n_submit=0, style={'height': '38px'}),
        ])]), className="mb-4"),  # Espaço para baixo entre os cards
        dbc.Row(dbc.Card([dbc.CardBody([
            html.H5("REALIZADO POR VENDEDOR", className="card-title"),
            html.P(id="realizado-vendedor", className="card-text"),
        ])]), className="mb-4"),
        dbc.Row(dbc.Card([dbc.CardBody([
            html.H5("PROJEÇÃO POR VENDEDOR", className="card-title"),
            html.P(id="projecao-vendedor", className="card-text"),
        ])]), className="mb-4")
    ], width=2),
    # Coluna de espaçamento à direita dos cards centrais
    dbc.Col(dbc.Card(dbc.CardBody([
            dcc.Graph(
                id='right-chart',
                style={'height': '100%', 'width': '100%'},  
                config={'responsive': True})]),
                style={'height': '400px'}  
            ),width=4, className="mb-4"
        ),], className="mb-4"),
    
    dbc.Row([
    dbc.Col(dbc.Card([
    dbc.CardHeader("QUANTIDADE DE CLIENTES ATENDIDOS", className="text-center"),
    dbc.CardBody([
        dbc.Row([
            dbc.Col(html.Div("AGREGADO", className="text-center"), width=4),
            dbc.Col(html.Div("VIDRO COMUM", className="text-center"), width=4),
            dbc.Col(html.Div("TEMPERADO", className="text-center"), width=4)
        ]),
        dbc.Row([
            dbc.Col(html.Div(id="clientes-atendidos-vidro", className="text-center"), width=4),
            dbc.Col(html.Div(id="clientes-atendidos-agregados", className="text-center"), width=4),
            dbc.Col(html.Div(id="clientes-atendidos-temperados", className="text-center"), width=4)
                ])
            ])
        ]), width=3),
    dbc.Col(dbc.Card([
        dbc.CardHeader("VENDA POR LOCALIDADE", className="text-center"),
        dbc.CardBody([
            dbc.Row([
                dbc.Col(html.Div("CAPITAL", className="text-center"), width=6),
                dbc.Col(html.Div("INTERIOR", className="text-center"), width=6)
            ]),
            dbc.Row([
                dbc.Col(html.Div(id="vendas-capital", className="text-center"), width=6),
                dbc.Col(html.Div(id="vendas-interior", className="text-center"), width=6)
                    ])
                ])
            ]),width=2),
        dbc.Col(dbc.Card([
                dbc.CardHeader("RECOMPRA DOS ÚLTIMOS 6 MESES", className="text-center"),
                dbc.CardBody(id="recompra-ultimos-6-meses")
            ]), width=5),
    ], className="mb-4"),
    
    
    # Tabela FATURAMENTO DOS ÚLTIMOS 3 MESES DE VIDRO
    dbc.Row(
        [
            html.Div(id='faturamento_vidro_card_container')
        ],
        className="mb-4"
    ),

    # Tabela CATEGORIA VIDRO
    dbc.Row(
        [
            html.Div(id='categoria_vidro_table_container')  
        ],
        className="mb-4"
    ),
    
    # Tabela FATURAMENTO DOS ÚLTIMOS 3 MESES DE AGREGADOS
    dbc.Row(
        [
            html.Div(id='faturamento_agregados_3m_container')
        ],
        className="mb-4"
    ),

    # Tabela Categoria Agregados
    dbc.Row(
        [
            html.Div(id='categoria_agregados_table_container')
        ],
        className="mb-4"
    ),

    # Tabela Cliente Sintético
    dbc.Row(
            [
                dbc.Col(cliente_sintetico_card, width=12)
            ],
        ),
    dcc.Download(id='download-excel')
    
], fluid=True)

if __name__ == "__main__":
    app.run_server(debug=True)
