from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# Nome do arquivo Excel
EXCEL_FILE = 'transf_estoque.xlsx'

# Função para inicializar o arquivo Excel se ele não existir
def init_excel(file_name):
    if not os.path.exists(file_name):
        df = pd.DataFrame(columns=['Nome', 'Cartão', 'Local de Origem', 'Local Destino', 'Data e Hora'])
        df.to_excel(file_name, index=False, sheet_name='Sheet1')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    nome = request.form['nome']
    cartao = request.form['cartao']
    local_origem = request.form['local_origem']
    local_destino = request.form['local_destino']
    datetime_submitted = request.form['datetime']

    # Carregar o arquivo Excel existente ou inicializar um novo
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name='Sheet1')
    else:
        df = pd.DataFrame(columns=['Nome', 'Cartão', 'Local de Origem', 'Local Destino', 'Data e Hora'])

    # Adicionar os novos dados
    new_data = pd.DataFrame([[nome, cartao, local_origem, local_destino, datetime_submitted]], 
                            columns=['Nome', 'Cartão', 'Local de Origem', 'Local Destino', 'Data e Hora'])
    df = pd.concat([df, new_data], ignore_index=True)

    # Salvar os dados atualizados de volta no arquivo Excel
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Atualizar a aba "LOCALIZAÇÃO" com a última data, hora e local destino de cada "Cartão"
    last_entries = df.sort_values('Data e Hora').groupby('Cartão').tail(1)[['Cartão', 'Local Destino', 'Data e Hora']]
    last_entries.columns = ['Cartão', 'Último Local Destino', 'Última Data e Hora']

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        last_entries.to_excel(writer, index=False, sheet_name='LOCALIZAÇÃO')

    return redirect(url_for('index'))

if __name__ == '__main__':
    init_excel(EXCEL_FILE)
    app.run(debug=True)
