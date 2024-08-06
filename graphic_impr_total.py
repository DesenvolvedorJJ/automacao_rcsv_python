import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

# Carregar o arquivo Excel e ler os dados
df = pd.read_excel('soma_total.xlsx')

# Verificar se as colunas necessárias estão presentes
if 'Mês' in df.columns and 'Quantidade total do mês' in df.columns:
    # Criar um novo Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'Dados'

    # Adicionar os dados do DataFrame na planilha
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Criar o gráfico de colunas
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "QUANTIDADE TOTAL DE IMPRESSÕES POR MÊS"
    chart.y_axis.title = ' '
    chart.x_axis.title = ' '

    # Adicionar os dados ao gráfico
    data = Reference(ws, min_col=2, min_row=1, max_row=7, max_col=2)
    categories = Reference(ws, min_col=1, min_row=2, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # Adicionar rótulos de dados no topo das colunas
    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.dLblPos = "t"  # Posicionar os rótulos no topo das colunas

    # Remover a legenda do gráfico
    chart.legend = None

    # Adicionar o gráfico à planilha
    ws.add_chart(chart, "G1")

    # Salvar o arquivo Excel com o gráfico
    wb.save('impr_total_com_grafico.xlsx')
else:
    print("Erro.")