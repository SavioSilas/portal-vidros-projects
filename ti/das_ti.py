"""
Projeto: Farol de Vendas - Dashboard Interativo
* @copyrigth    Sávio Silas <svosilas@gmail.com> - Desenvolvedor Portal Vidros
* @date         07 Março 2024
* @file         dash_vendas.py
* @brief        Dash de visialização de chamados
"""

import dash
from dash import html, dcc, Input, Output
import pandas as pd
import plotly.express as px
import mysql.connector
from flask import send_file
from dash import callback_context
import dash
from dash import html, dcc, Input, Output
import dash_table
import pandas as pd
import plotly.express as px
import mysql.connector

app = dash.Dash(__name__)

app.layout = html.Div([
    html.Div([
        dcc.DatePickerRange(
            id='date-picker-range',
            display_format='DD/MM/YYYY'
        )        
    ], style={'display': 'flex', 'margin': '10px 0'}),

    # Botões de download
    html.Div([
        html.Button(
            "Baixar Excel",
            id="download-button",
            n_clicks=0,
            style={'display': 'inline-block', 'margin': '10px 5px', 'padding': '10px', 'border': 'none',
                   'background-color': '#007bff', 'color': 'white', 'cursor': 'pointer', 'border-radius': '10px'}
        ),
        html.Button(
            "Baixar Excel (Dados Filtrados)",
            id="download-button-filtered",
            n_clicks=0,
            style={'display': 'inline-block', 'margin': '10px 10px', 'padding': '10px', 'border': 'none',
                   'background-color': '#28a745', 'color': 'white', 'cursor': 'pointer', 'border-radius': '10px'}
        ),
        dcc.Download(id="download-dataframe-xls"),
        dcc.Download(id="download-dataframe-xls-filtered")
    ], style={'text-align': 'right', 'padding-right': '10px'}),

    html.Div([
    html.Div(id='total-tickets', className='card', style={'width': '24%', 'height': '150px', 'display': 'inline-block', 'margin': '10px', 'background-color': '#D3E4CD', 'border-radius': '10px', 'text-align': 'center', 'vertical-align': 'middle', 'line-height': '150px', 'border': '2px solid #000'}),  
    html.Div(id='pending-tickets', className='card', style={'width': '24%', 'height': '150px', 'display': 'inline-block', 'margin': '10px', 'background-color': '#D3E4CD', 'border-radius': '10px', 'text-align': 'center', 'vertical-align': 'middle', 'line-height': '150px', 'border': '2px solid #000'}),  
    html.Div(id='closed-tickets', className='card', style={'width': '24%', 'height': '150px', 'display': 'inline-block', 'margin': '10px', 'background-color': '#D3E4CD', 'border-radius': '10px', 'text-align': 'center', 'vertical-align': 'middle', 'line-height': '150px', 'border': '2px solid #000'}),  
    html.Div(id='incident-tickets', className='card', style={'width': '24%', 'height': '150px', 'display': 'inline-block', 'margin': '10px', 'background-color': '#FAD2E1', 'border-radius': '10px', 'text-align': 'center', 'vertical-align': 'middle', 'line-height': '150px', 'border': '2px solid #000'}),  
], style={'display': 'flex', 'justify-content': 'space-around', 'flex-wrap': 'nowrap', 'margin': '10px'}),


    html.Div([
        dcc.Graph(id='chamados-mes', style={'flex': '1', 'minWidth': '300px'}),
        dcc.Graph(id='chamados-por-tecnico', style={'flex': '1', 'minWidth': '300px'}),
        dcc.Graph(id='chamados-categoria-pie', style={'flex': '1', 'minWidth': '300px'}),
        dcc.Download(id="download-dataframe-xls"),
    ], style={'display': 'flex', 'flex-wrap': 'wrap', 'justify-content': 'space-between'}), #'border': '2px solid #000

    # Atualização do Dash
    dcc.Interval(
        id='interval-component',
        interval=300*1000,  # milissegundos
        n_intervals=0
    )
], style={'position': 'relative', 'padding': '10px'})

# Callback para o botão "Baixar Excel"
@app.callback(
    Output("download-dataframe-xls", "data"),
    Input("download-button", "n_clicks"),
    prevent_initial_call=True  # Isso impede que o callback seja chamado na carga inicial
)
def download_as_excel(n_clicks):
    if n_clicks > 0:  # Verifica se o botão foi clicado
        df = fetch_data()  # Buscar os dados para download
        return dcc.send_data_frame(df.to_excel, "dados_completos.xlsx", index=False)
    return None  # Não retorna nada se o botão não foi clicado

# Callback para o botão "Baixar Excel (Filtrado)"
@app.callback(
    Output("download-dataframe-xls-filtered", "data"),
    [Input("download-button-filtered", "n_clicks"),
     Input('date-picker-range', 'start_date'),
     Input('date-picker-range', 'end_date')],
    prevent_initial_call = True
)
def download_filtered_data(n_clicks, start_date, end_date):
    # Usar callback_context para verificar qual entrada acionou o callback
    ctx = callback_context
    if not ctx.triggered:
        button_id = 'No clicks yet'
    else:
        button_id = ctx.triggered[0]['prop_id'].split('.')[0]

    # Verificar se o botão de download filtrado foi clicado
    if button_id == "download-button-filtered" and n_clicks > 0:
        df = fetch_data()  # Buscar os dados
        df['Data Abertura'] = pd.to_datetime(df['Data Abertura'], format="%d/%m/%Y")
        df['Data Fechamento'] = pd.to_datetime(df['Data Fechamento'], format="%d/%m/%Y")

        # Filtrar os dados com base nas datas selecionadas
        if start_date and end_date:
            filtered_df = df[(df['Data Abertura'] >= start_date) & (df['Data Fechamento'] <= end_date)]
        else:
            filtered_df = df

        return dcc.send_data_frame(filtered_df.to_excel, "dados_filtrados.xlsx", index=False)
    return None  

@app.callback(
    [
        Output('total-tickets', 'children'),
        Output('pending-tickets', 'children'),
        Output('closed-tickets', 'children'),
        Output('incident-tickets', 'children'),
        Output('chamados-mes', 'figure'),
        Output('chamados-por-tecnico', 'figure'),
        Output('chamados-categoria-pie', 'figure'),
        Output('date-picker-range', 'min_date_allowed'),
        Output('date-picker-range', 'max_date_allowed'),
        Output('date-picker-range', 'start_date'),
        Output('date-picker-range', 'end_date')
    ],
    [
        Input('date-picker-range', 'start_date'),
        Input('date-picker-range', 'end_date'),
        Input('interval-component', 'n_intervals')
    ]
)
def update_components(start_date, end_date, n_intervals):
    df = fetch_data()  # Buscar os dados sempre que o callback for disparado

    df['Data Abertura'] = pd.to_datetime(df['Data Abertura'])
    df['Data Fechamento'] = pd.to_datetime(df['Data Fechamento'])

    if start_date is None:
        start_date = df['Data Abertura'].min()

    if end_date is None:
        end_date = df['Data Abertura'].max()

    filtered_df = df[(df['Data Abertura'] >= start_date) & (df['Data Abertura'] <= end_date)]
    
    # Atualizar métricas de tickets
    total_tickets = len(filtered_df)
    pending_tickets = len(filtered_df[filtered_df['Status'] == 'Pendente'])
    closed_tickets = len(filtered_df[filtered_df['Status'] == 'Fechado'])
    incident_tickets = len(df[df['Categoria'].str.contains('INCIDENTE', case=False, na=False)])
    
    # Atualizar gráfico de chamados por mês
    filtered_df['Ano-Mês'] = filtered_df['Data Abertura'].dt.strftime('%Y-%m')
    chamados_mes = filtered_df['Ano-Mês'].value_counts().sort_index()
    fig_mes = px.bar(chamados_mes, x=chamados_mes.index, y=chamados_mes.values, labels={'x': 'Mês', 'y': 'Quantidade'})
    fig_mes.update_layout(title='Chamados por Mês')
    
    # Atualizar gráfico de chamados por técnico
    chamados_por_tecnico = filtered_df['Técnico'].value_counts().reset_index(name='Quantidade')
    chamados_por_tecnico.rename(columns={'index': 'Técnico'}, inplace=True)  # Garanta que a coluna é nomeada corretamente
    fig_tecnico = px.bar(chamados_por_tecnico, x='Quantidade', y='Técnico', orientation='h', labels={'Técnico': 'Técnico', 'Quantidade': 'Quantidade de Chamados'})
    fig_tecnico.update_layout(title='Chamados por Técnico')
    
    # Atualizar gráfico de chamados por categoria
    top_categorias = filtered_df['Categoria'].value_counts().nlargest(10)
    outras_quantidade = filtered_df['Categoria'].value_counts().iloc[10:].sum()
    #top_categorias = top_categorias._append(pd.Series(outras_quantidade, index=['Outros']))
    fig_categoria = px.pie(top_categorias, values=top_categorias.values, names=top_categorias.index, title='Top 10 Chamados por Categoria')
    #fig_categoria.update_traces(textinfo='percent+label') 
    fig_categoria.update_layout(showlegend=True)
    
    # Atualizar gráfico de chamados por mês
    fig_mes.update_layout(
        title={
            'text': "Chamados por Mês",
            'x': 0.5,
            'xanchor': 'center'
        }
    )

    # Atualizar gráfico de chamados por técnico
    fig_tecnico.update_layout(
        title={
            'text': "Chamados por Técnico",
            'x': 0.5,
            'xanchor': 'center'
        }
    )

    # Atualizar gráfico de chamados por categoria
    fig_categoria.update_layout(
        title={
            'text': "Top 10 Chamados por Categoria",
            'x': 0.5,
            'xanchor': 'center'
        }
    )

    return (
        f'Total de Tickets: {total_tickets}',
        f'Chamados Pendentes: {pending_tickets}',
        f'Chamados Fechados: {closed_tickets}',
        f'Incidentes: {incident_tickets}',
        fig_mes,
        fig_tecnico,
        fig_categoria,
        df['Data Abertura'].min(),
        df['Data Abertura'].max(),
        start_date,
        end_date
    )

if __name__ == '__main__':
    app.run_server(debug=True)
