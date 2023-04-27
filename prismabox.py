import pandas as pd
import re
from tkinter import Tk
import unidecode
from tkinter.filedialog import askopenfilename
import xlsxwriter

def format_cpf_cnpj(value):
    # Extrai apenas os dígitos
    digits = re.sub('\D', '', str(value))
    # Formata CPF (11 caracteres)
    if len(digits) == 11:
        return f'{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}'
    # Formata CNPJ (14 caracteres)
    elif len(digits) == 14:
        return f'{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}'
    # Retorna valor original se não for CPF nem CNPJ
    else:
        return value

def format_phone(value):
    # Extrai apenas os dígitos
    digits = re.sub('\D', '', str(value))
    # Formatação para 11 dígitos
    if len(digits) == 11:
        return f'({digits[:2]}){digits[2:7]}-{digits[7:]}'
    # Formatação para 10 dígitos
    elif len(digits) == 10:
        return f'({digits[:2]})9{digits[2:6]}-{digits[6:]}'
    # Retorna vazio se não tiver mais de 9 dígitos
    elif len(digits) <= 9:
        return ''
    # Retorna valor original se não for telefone
    else:
        return value

# Abre a janela de seleção de arquivo
root = Tk()
root.withdraw()
filename = askopenfilename()

# Lê o arquivo
df = pd.read_excel(filename)

# Formata CPF ou CNPJ
df['CPF ou CNPJ'] = df['CPF ou CNPJ'].apply(format_cpf_cnpj)

# Formata número de celular
df['Celular'] = df['Celular'].apply(format_phone)

# Função de estilo para pintar as células duplicadas da coluna de amarelo ou azul
def highlight_duplicates(s):
    if s.name == 'CPF ou CNPJ':
        return ['background-color: yellow' if v and df.duplicated(subset=[s.name], keep=False).iloc[i] else '' for i,v in enumerate(s)]
    elif s.name == 'Celular':
        return ['background-color: blue' if v and df.duplicated(subset=[s.name], keep=False).iloc[i] else '' for i,v in enumerate(s)]
    else:
        return ['' for i in enumerate(s)]

# Função de estilo para pintar as células inválidas da coluna de vermelho
def highlight_invalid(s):
    return ['background-color: red' if len(re.sub('\D', '', str(v))) <= 9 else '' for v in s]

# Função de estilo para pintar as células inválidas da coluna "Estado (Dois dígitos)" de vermelho
def highlight_invalid_state(s):
    estados = {
        'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS',
        'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC',
        'SP', 'SE', 'TO'
    }
    return ['background-color: red' if v.upper() not in estados else '' for v in s]

# Arrumar Siglas da maneira adequada:

# Dicionário com as siglas dos estados brasileiros
estados = {
    'acre': 'AC',
    'alagoas': 'AL',
    'amapa': 'AP',
    'amazonas': 'AM',
    'bahia': 'BA',
    'ceara': 'CE',
    'distrito federal': 'DF',
    'espirito santo': 'ES',
    'goias': 'GO',
    'maranhao': 'MA',
    'mato grosso': 'MT',
    'mato grosso do sul': 'MS',
    'minas gerais': 'MG',
    'para': 'PA',
    'paraiba': 'PB',
    'parana': 'PR',
    'pernambuco': 'PE',
    'piaui': 'PI',
    'rio de janeiro': 'RJ',
    'rio grande do norte': 'RN',
    'rio grande do sul': 'RS',
    'rondonia': 'RO',
    'roraima': 'RR',
    'santa catarina': 'SC',
    'sao paulo': 'SP',
    'sergipe': 'SE',
    'tocantins': 'TO'
}

# Substitui os nomes dos estados pelas siglas na coluna "Estado (Dois dígitos)"
df['Estado (Dois dígitos)'] = df['Estado (Dois dígitos)'].str.normalize('NFKD').apply(unidecode.unidecode).str.lower().replace(estados)

# Transforma a coluna "Estado (Dois dígitos)" em maiúsculas
df['Estado (Dois dígitos)'] = df['Estado (Dois dígitos)'].str.upper()

styled_df = df.style.apply(highlight_duplicates, subset=['CPF ou CNPJ', 'Celular']).apply(highlight_invalid, subset=['Celular']).apply(highlight_invalid_state, subset=['Estado (Dois dígitos)'])

styled_df.to_excel('dados_formatados.xlsx', index=False)

print('Correções e flags colocadas com sucesso!')
