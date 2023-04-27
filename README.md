# Algoritmo-Pandas-PrismaBox para a limpeza da planilha de clientes Prisma Box 

Este é um script Python que faz correções em um arquivo Excel com informações de clientes. Ele utiliza a biblioteca Pandas para ler o arquivo, realizar as operações de formatação e exportar o arquivo corrigido em Excel.

## Dependências

Para rodar o script, é necessário ter o Python 3 instalado na sua máquina, além das bibliotecas Pandas, xlsxwriter e unidecode. As dependências podem ser instaladas com os seguintes comandos em um terminal:

Certifique-se de ter o Python 3 instalado em seu computador. Você pode baixar o Python em [python.org](https://www.python.org/).

Instale as dependências usando o pip. Abra o terminal ou prompt de comando e digite o seguinte comando:

```shell
pip install pandas xlsxwriter unidecode
```

## Como Usar

1. Faça o download do script de limpeza de dados em formato .py e do arquivo Excel que deseja limpar.

2. Selecione na janela do gerenciador de arquivos o arquivo xlsx, caso de erro tente baixar o Tk:

Arch Linux:

```shell
sudo pacman -S tk 
```

Ubuntu:

```shell
sudo apt-get install python3-tk
```

CentOS ou Fedora:

```shell
sudo dnf install python3-tkinter
```

3. Execute o script com o seguinte comando:

```shell
python prismabox.py
```

4. O script é composto por algumas funções que realizam as seguintes operações:

- format_cpf_cnpj(value): recebe um valor e formata para o padrão de CPF ou CNPJ brasileiro, dependendo do tamanho do valor.

- format_phone(value): recebe um valor e formata para o padrão de telefone brasileiro com DDD, adicionando o "9" para números de celular com 10 dígitos.

- highlight_duplicates(s): recebe uma coluna de dados e retorna uma lista com o estilo das células duplicadas, destacando com cor amarela para CPF/CNPJ e azul para celular.

- highlight_invalid(s): recebe uma coluna de dados e retorna uma lista com o estilo das células inválidas, destacando com cor vermelha para valores que possuem menos de 10 dígitos.

Além disso, o script também faz uma correção na coluna "Estado (Dois dígitos)", substituindo o nome do estado por sua sigla de duas letras, e transforma todos os valores em maiúsculas.


Ao finalizar a execução, o script imprime a mensagem "Correções e flags colocadas com sucesso!".
