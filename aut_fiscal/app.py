import pandas as pd
import mysql.connector
import numpy as np
from datetime import datetime
from datetime import datetime
from flask import Flask, request, render_template, redirect, url_for
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.sql import text
from flask import jsonify, session

app = Flask(__name__, template_folder='templates')

def fetch_data():
    config = {
        'user': '',
        'password': '',
        'host': '',
        'database': ''
    }

    conn = mysql.connector.connect(**config)
    query = '''
    select ***
    '''
    
    df = pd.read_sql(query, conn)
    conn.close()
    return df

app.config['SQLALCHEMY_DATABASE_URI'] = '*.db'
app.config['SQLALCHEMY_BINDS'] = {
    'verificadas': ''
}


app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

class NotaOriginal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    numero_nota = db.Column(db.String(120), unique=True)
    serie = db.Column(db.String(50))
    data_emissao = db.Column(db.DateTime)
    total_nota = db.Column(db.Float)
    cfop = db.Column(db.String(50))

    def __repr__(self):
        return f'<NotaOriginal {self.numero_nota}>'

class NotaVerificada(db.Model):
    __bind_key__ = 'verificadas'
    id = db.Column(db.Integer, primary_key=True)
    numero_nota = db.Column(db.String(120), unique=True)
    serie = db.Column(db.String(50))
    data_emissao = db.Column(db.DateTime)
    total_nota = db.Column(db.Float)
    cfop = db.Column(db.String(50))

    def __repr__(self):
        return f'<NotaVerificada {self.numero_nota}>'

with app.app_context():
    db.create_all()

@app.before_request
def create_tables():
    db.create_all()
    engine = db.get_engine(app, bind='verificadas')
    with engine.connect() as connection:
        connection.execute(text('CREATE TABLE IF NOT EXISTS nota_verificada (id INTEGER PRIMARY KEY, numero_nota VARCHAR(120), serie VARCHAR(50), data_emissao DATETIME, total_nota FLOAT, cfop VARCHAR(50))'))

@app.route('/clear_session')
def clear_session():
    session.clear()
    return 'Sessão limpa!'

@app.route('/')
def index_post():
    notas_originais = NotaOriginal.query.all()
    notas_verificadas = NotaVerificada.query.with_entities(NotaVerificada.numero_nota).all()
    notas_verificadas_ids = {nota.numero_nota for nota in notas_verificadas}
    notas_a_mostrar = [nota for nota in notas_originais if nota.numero_nota not in notas_verificadas_ids]
    cfop_5102 = df[df['cfop'] == '5102']
    cfop_5405 = df[df['cfop'] == '5405']
    cfop_5101 = df[df['cfop'] == '5101']
    max_rows = max(cfop_5102.shape[0], cfop_5405.shape[0], cfop_5101.shape[0]) if cfop_5102 is not None else 0

    return render_template('index.html', notas=notas_a_mostrar, max_rows=max_rows, cfop_5102=cfop_5102, cfop_5405=cfop_5405, cfop_5101=cfop_5101)

@app.route('/remover-nota', methods=['POST'])
def remover_nota():
    cfop = request.form['cfop']
    index = int(request.form['index'])
    if cfop == '5102':
        df.drop(df[df['cfop'] == '5102'].index[index], inplace=True)
    elif cfop == '5405':
        df.drop(df[df['cfop'] == '5405'].index[index], inplace=True)
    elif cfop == '5101':
        df.drop(df[df['cfop'] == '5101'].index[index], inplace=True)
    else:
        return jsonify({'success': False, 'message': 'CFOP inválido'})

    return jsonify({'success': True})

@app.route('/remover-nota-python', methods=['POST'])
def remover_nota_python():
    cfop = request.form['cfop']
    index = int(request.form['index'])

    if cfop == '5102':
        df.drop(df[df['cfop'] == '5102'].index[index], inplace=True)
    elif cfop == '5405':
        df.drop(df[df['cfop'] == '5405'].index[index], inplace=True)
    elif cfop == '5101':
        df.drop(df[df['cfop'] == '5101'].index[index], inplace=True)
    else:
        return jsonify({'success': False, 'message': 'CFOP inválido'})

    return jsonify({'success': True, 'df': df.to_dict('records')})

@app.route('/verificar-nota', methods=['POST'])
def verificar_nota():
    nota_id = request.form.get('nota_id')
    nota = NotaOriginal.query.get(nota_id)
    if nota:
        nova_nota = NotaVerificada(
            numero_nota=nota.numero_nota,
            serie=nota.serie,
            data_emissao=nota.data_emissao,
            total_nota=nota.total_nota,
            cfop=nota.cfop
        )
        db.session.add(nova_nota)
        db.session.commit()
        return redirect(url_for('index'))
    return 'Nota não encontrada', 404


df = fetch_data()
@app.route('/', methods=['GET', 'POST'])
def index():

    if 'df' not in session:
        df = fetch_data()
        session['df'] = df.to_json() 
    else:
        df = pd.read_json(session['df']) 

    
    message = "Digite um período para análise"  
    data_available = False 
    cfop_5404 = cfop_5405 = cfop_5101 = pd.DataFrame()
    max_rows = 0

    df['data_emissao'] = pd.to_datetime(df['data_emissao'], format='%d/%m/%Y')

    if request.method == 'POST' and 'start_date' in request.form and 'end_date' in request.form:
        start_date = datetime.strptime(request.form['start_date'], '%Y-%m-%d')
        end_date = datetime.strptime(request.form['end_date'], '%Y-%m-%d')
        
        df['data_emissao'] = pd.to_datetime(df['data_emissao'], format='%d/%m/%Y')
        df = df[(df['data_emissao'] >= start_date) & (df['data_emissao'] <= end_date)]
        df['data_emissao'] = df['data_emissao'].dt.strftime('%d/%m/%Y')

        data_available = True

        if 'remove_notes' in request.form:
            indices_to_remove = []
            for key, value in request.form.items():
                if 'selected_' in key and value == 'on':
                    _, cfop, index = key.split('_')
                    indices_to_remove.append(int(index))
            df = df.drop(indices_to_remove).reset_index(drop=True)
        

    else:
        start_date = end_date = datetime.now().date()
    
    start_date_str = start_date.strftime('%Y-%m-%d')
    end_date_str = end_date.strftime('%Y-%m-%d')
    df['data_emissao'] = pd.to_datetime(df['data_emissao'], format='%d/%m/%Y')
    df = df[(df['data_emissao'] >= start_date) & (df['data_emissao'] <= end_date)]
    df['data_emissao'] = df['data_emissao'].dt.strftime('%d/%m/%Y')

    duplicated_notes = df['numero_nota'].duplicated(keep=False)

    df['observacao'] = np.where(
        (df['valor_icms'] == 0) & ~duplicated_notes, 'Nota sem valor do ICMS', 
        np.where(
            (df['valor_aliq'] == 0) & ~duplicated_notes, 'Nota sem valor da Aliquota', ''
        )
    )
    
    df.loc[duplicated_notes & (df['valor_icms'] == 0), 'observacao'] = 'Nota com produto sem ICMS'
    df.loc[duplicated_notes & (df['valor_aliq'] == 0), 'observacao'] += ' Nota com produto sem Aliquota'

    duplicated_notes = df['numero_nota'].duplicated(keep=False)

    cfop_5404 = df[df['cfop'] == '5404']
    cfop_5404 = cfop_5404[(cfop_5404['valor_icms'] != 0) | (cfop_5404['valor_aliq'] != 0)]
    cfop_5404['observacao'] = np.where(
        (cfop_5404['valor_icms'] != 0) & ~duplicated_notes, 'Nota com valor de ICMS',
        np.where(
            (cfop_5404['valor_aliq'] != 0) & ~duplicated_notes, 'Nota com valor da Aliquota', ''
        )
    )

    cfop_5404.loc[duplicated_notes & (cfop_5404['valor_icms'] != 0), 'observacao'] = 'Nota com produto com ICMS'
    cfop_5404.loc[duplicated_notes & (cfop_5404['valor_aliq'] != 0), 'observacao'] += ' Nota com produto com Aliquota'

    cfop_5405 = df[(df['cfop'] == '5405') & ((df['valor_icms'] != 0) | (df['valor_aliq'] != 0))]
    cfop_5405['observacao'] = np.where((cfop_5405['valor_icms'] != 0), 'Nota com ICMS', '')
    cfop_5405['observacao'] += np.where((cfop_5405['valor_aliq'] != 0), ' Nota com Aliquota', '')
    
    cfop_5101 = df[df['cfop'] == '5101']
    cfop_5101 = df[(df['cfop'] == '5101') & (df['loja'] == 'PORTAL VIDROS (FILIAL)') ]
    cfop_5101 = df[(df['cfop'] == '5101') & (df['loja'] == 'PORTAL VIDROS (MATRIZ COMÉRCIO)')]
    
    cfop_5404 = df[df['cfop'] == '5102']
    cfop_5404 = cfop_5404[cfop_5404['observacao'].notna() & (cfop_5404['observacao'] != '')]

    if request.method == 'POST' and request.form.get('remove_notes'):
        for index in range(max(cfop_5404.shape[0], cfop_5405.shape[0], cfop_5101.shape[0])):
            if request.form.get(f'selected_5104_{index}', False):
                cfop_5404.drop(index, inplace=True)
            if request.form.get(f'selected_5405_{index}', False):
                cfop_5405.drop(index, inplace=True)
            if request.form.get(f'selected_5101_{index}', False):
                cfop_5101.drop(index, inplace=True)

    max_rows = max(cfop_5404.shape[0], cfop_5405.shape[0], cfop_5101.shape[0])
    cfop_5404['selecionado'] = False
    cfop_5405['selecionado'] = False
    cfop_5101['selecionado'] = False

    return render_template('index.html', cfop_5102=cfop_5404, cfop_5405=cfop_5405, cfop_5101=cfop_5101, 
                           max_rows=max_rows, start_date=start_date_str, end_date=end_date_str, message=message, data_available=data_available)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0')
