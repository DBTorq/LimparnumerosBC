from flask import Flask, request, render_template, send_file, flash
import os
import re
import pandas as pd
import phonenumbers

app = Flask(__name__)

UPLOAD_FOLDER = r'C:\Users\D-BTo\OneDrive\API_CRIS\Outros\LimparnumerosBC\uploads'
OUTPUT_FOLDER = r'C:\Users\D-BTo\OneDrive\API_CRIS\Outros\LimparnumerosBC\outputs'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Variável global para armazenar logs
logs = []

@app.route('/get_logs', methods=['GET'])
def get_logs():
    return {'logs': logs}

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        logs.append('Nenhum arquivo enviado!', 'error')
        return render_template('form.html')

    file = request.files['file']
    codigos_pais_input = request.form['codigos_pais']
    codigos_pais = [codigo.strip().upper() for codigo in codigos_pais_input.split(',')]

    if file.filename == '':
        logs.append('Nenhum arquivo selecionado!', 'error')
        return render_template('form.html')

    # Salvar o arquivo enviado
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    # Processar o arquivo
    output_path = os.path.join(OUTPUT_FOLDER, 'planilha_corrigida.xlsx')
    processar_planilha(file_path, output_path, codigos_pais)

    return send_file(output_path, as_attachment=True)

def processar_planilha(file_path, output_path, codigos_pais):
    global logs
    logs.append("Iniciando o processamento da planilha.")

    spreadsheet = pd.read_excel(file_path)
    colunas_corretas = {
        'telefone': ['telefone', 'Telefone', 'TELEFONE', 'PHONE', 'CELULAR', 'Numero', 'Número', 'numero', 'número'],
        'nome': ['nome', 'Nome', 'NOME', 'nome-completo', 'Nome-Completo', 'Clientes', 'Cliente', 'Leads', 'Lead'],
        'etiquetas': ['etiquetas', 'Etiqueta', 'etiqueta', 'Etiquetas', 'ETIQUETA']
    }


    logs.append("Corrigindo os nomes das colunas.")
    spreadsheet.columns = corrigir_nomes_colunas(spreadsheet.columns, colunas_corretas)
    colunas_finais = ['telefone', 'nome', 'etiquetas']
    for coluna in colunas_finais:
        if coluna not in spreadsheet.columns:
            spreadsheet[coluna] = ''

    spreadsheet = spreadsheet[colunas_finais]
    spreadsheet['telefone'] = spreadsheet['telefone'].apply(lambda x: formatar_telefone(x, codigos_pais))
    spreadsheet = spreadsheet.dropna(subset=['telefone'])
    spreadsheet.to_excel(output_path, index=False)

def corrigir_nomes_colunas(colunas, colunas_corretas):
    colunas_corrigidas = []
    for coluna in colunas:
        coluna_corrigida = None
        for nome_correto, variações in colunas_corretas.items():
            if coluna in variações:
                coluna_corrigida = nome_correto
                break
        if not coluna_corrigida:
            logs.append(f"Cabeçalho da coluna '{coluna}' não corresponde a telefone, nome ou etiquetas. Removendo.")
        colunas_corrigidas.append(coluna_corrigida or coluna)
    return colunas_corrigidas

def formatar_telefone(telefone, codigos_pais):
    if not telefone or telefone == 'nan':
        logs.append('Telefone vazio. Não foi possível ser importado.', 'warning')
        return None

    telefone_str = str(telefone).strip()

    if 'BR' in codigos_pais:
        telefone_formatado = formatar_telefone_br(telefone_str)
    else:
        for codigo_pais in codigos_pais:
            try:
                numero = phonenumbers.parse(telefone_str, codigo_pais)
                if phonenumbers.is_valid_number(numero):
                    return phonenumbers.format_number(numero, phonenumbers.PhoneNumberFormat.E164)
            except phonenumbers.NumberParseException:
                continue

    if telefone_formatado:
        return telefone_formatado
    else:
        logs.append(f"Telefone {telefone_str} inválido. Não foi possível ser importado.")
    return None

def formatar_telefone_br(telefone):
    telefone_limpo = re.sub(r'[()\s-]', '', telefone)
    if len(telefone_limpo) == 10:
        ddd = telefone_limpo[:2]
        numero = telefone_limpo[2:]
        return "55" + ddd + '9' + numero
    elif len(telefone_limpo) == 11:
        return "55" + telefone_limpo
    elif len(telefone_limpo) == 13 and telefone_limpo.startswith('55'):
        return telefone_limpo
    elif len(telefone_limpo) == 12 and telefone_limpo.startswith('55'):
        telefone_sem_55 = telefone_limpo[2:]
        ddd = telefone_sem_55[:2]
        numero_sem_ddd = telefone_sem_55[2:]
        if len(numero_sem_ddd) == 9:
            return "55" + ddd + numero_sem_ddd
        elif len(numero_sem_ddd) == 8:
            return "55" + ddd + '9' + numero_sem_ddd
    return None

if __name__ == '__main__':
    app.run(debug=True)
