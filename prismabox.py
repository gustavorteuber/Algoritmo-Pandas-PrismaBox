import pandas as pd
import re

# Função para formatar o número de celular
def formatar_celular(celular):
    # Remove todos os caracteres que não sejam números
    numeros = re.sub(r'\D', '', celular)
    # Verifica se o número tem 11 dígitos (sem o DDD)
    if len(numeros) == 11:
        # Formata como (XX)XXXXX-XXXX
        celular_formatado = f'({numeros[:2]}){numeros[2:7]}-{numeros[7:]}'
    # Verifica se o número tem 10 dígitos (sem o DDD)
    elif len(numeros) == 10:
        # Formata como (XX)XXXX-XXXX
        celular_formatado = f'({numeros[:2]}){numeros[2:6]}-{numeros[6:]}'
    # Se não estiver nos formatos conhecidos, retorna o número original
    else:
        celular_formatado = celular
    return celular_formatado

# Carrega a planilha para o pandas
df = pd.read_excel('clientes.xlsx')

# Cria uma lista vazia para armazenar os CPFs e CNPJs inválidos
invalidos = []

# Loop pelos registros da planilha
for i, row in df.iterrows():
    # Formata o CPF ou CNPJ
    if 11 <= len(str(row['CPF ou CNPJ'])) <= 14:
        cpf = f'{row["CPF ou CNPJ"][:3]}.{row["CPF ou CNPJ"][3:6]}.{row["CPF ou CNPJ"][6:9]}-{row["CPF ou CNPJ"][9:]}'
        df.at[i, 'CPF ou CNPJ'] = cpf
    elif 14 <= len(str(row['CPF ou CNPJ'])) <= 18:
        cnpj = f'{row["CPF ou CNPJ"][:2]}.{row["CPF ou CNPJ"][2:5]}.{row["CPF ou CNPJ"][5:8]}/{row["CPF ou CNPJ"][8:12]}-{row["CPF ou CNPJ"][12:]}'
        df.at[i, 'CPF ou CNPJ'] = cnpj
    else:
        invalidos.append(row['CPF ou CNPJ'])
    # Formata o celular
    df.at[i, 'Celular'] = formatar_celular(str(row['Celular']))

# Cria uma planilha somente com os CPFs e CNPJs inválidos
df_invalidos = pd.DataFrame({'CPF ou CNPJ': invalidos})

# Verifica os CPFs duplicados
cpf_duplicados = df[df['CPF ou CNPJ'].duplicated(keep=False)]

# Cria uma planilha somente com os CPFs duplicados
df_duplicados = cpf_duplicados.sort_values('CPF ou CNPJ').drop_duplicates(subset='CPF ou CNPJ', keep='first')

# Salva as planilhas em arquivos .xlsx
with pd.ExcelWriter('corrigidakkkk.xlsx') as writer:
    df.to_excel(writer, sheet_name='Tabela Original', index=False)
with pd.ExcelWriter('cpjs_duplicados.xlsx') as writer:
    df_cnpj_duplicados.to_excel(writer, sheet_name='CNPJs ou CPFs Duplicados', index=False)
    df_invalidos.to_excel(writer, sheet_name='CPF ou CNPJ Inválidos', index=False)
    df_duplicados.to_excel(writer, sheet_name='CPF ou CNPJ Duplicados', index=False)

