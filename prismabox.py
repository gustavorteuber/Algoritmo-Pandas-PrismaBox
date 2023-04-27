import pandas as pd
import re
from tkinter import Tk
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

# Função de estilo para pintar as células duplicadas da coluna de amarelo
def highlight_duplicates(s):
    return ['background-color: yellow' if v and df.duplicated(subset=[s.name], keep=False).iloc[i] else '' for i,v in enumerate(s)]

# Função de estilo para pintar as células inválidas da coluna de vermelho
def highlight_invalid(s):
    return ['background-color: red' if len(re.sub('\D', '', str(v))) <= 9 else '' for v in s]

# Aplicar as funções de estilo nas colunas "CPF ou CNPJ" e "Celular"
styled_df = df.style.apply(highlight_duplicates, subset=['CPF ou CNPJ']).apply(highlight_invalid, subset=['Celular'])

# Salvar o arquivo Excel com as células pintadas
styled_df.to_excel('dados_formatados.xlsx', index=False)
