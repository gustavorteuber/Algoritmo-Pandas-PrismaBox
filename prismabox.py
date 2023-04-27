import pandas as pd
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename

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

# Abre a janela de seleção de arquivo
root = Tk()
root.withdraw()
filename = askopenfilename()

# Lê o arquivo
df = pd.read_excel(filename)

# Formata CPF ou CNPJ
df['CPF ou CNPJ'] = df['CPF ou CNPJ'].apply(format_cpf_cnpj)

# Cria um novo arquivo Excel com os dados formatados
df.to_excel('dados_formatados.xlsx', index=False)
