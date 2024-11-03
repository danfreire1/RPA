"""
Projeto: Correção de ortografia em textos usando a API da OpenAI
Descrição: Este projeto é responsável por corrigir textos disponibilizados em uma planilha Excel utilizando a API da OpenAI.
Autor: Daniel Freire da Costa
Data de Criação: 27 de Outubro de 2024
Versão: 1.0
Dependências: 
    - requests
    - json
    - os
    - pandas

Instruções de Uso:
1. Configure a chave da API da OpenAI em suas variáveis de ambiente.
2. Coloque os textos a serem corrigidos na planilha Excel especificada.
3. Execute o script para obter os textos corrigidos.

Referências:
Vídeo no Youtube de Mukesh Kala: https://www.youtube.com/watch?v=HIePjxzTuEk&list=WL&index=1
Documentação OpenAI: https://platform.openai.com/docs/api-reference/making-requests

"""

# Libs
import requests
import json
import os
import pandas as pd


# Funções
def lerArquivoExcel(filePath, sheet):
    """
    Lê uma planilha Excel e retorna um DataFrame.

    Parâmetros:
    filePath (str): O caminho do arquivo Excel a ser lido.
    sheet (str): O nome da planilha que será lida.

    Retorna:
    DataFrame: Um DataFrame contendo os dados da planilha especificada.
    """
    
    # ler planilha Excel 
    df = pd.read_excel(filePath, sheet_name=sheet)
    return df

def inserirTextoCorrigido(df, index, coluna, valor, filePath, sheet):
    """
    Insere um texto corrigido em uma coluna específica de um DataFrame e salva o DataFrame atualizado em um arquivo Excel.

    Parâmetros:
    df (DataFrame): O DataFrame a ser atualizado.
    index (int): O índice da linha onde o texto corrigido será inserido.
    coluna (str): O nome da coluna onde o texto corrigido será inserido.
    valor (str): O texto corrigido a ser inserido.
    filePath (str): O caminho do arquivo Excel onde o DataFrame atualizado será salvo.
    """

    # Substituir o valor do índice e da coluna do dataframe pelo texto corrigido
    df.at[index, coluna] = valor

    # Sobrescrever a planilha com os dados atualizados
    df.to_excel(filePath, sheet_name=sheet, index=False)

def chamadaApiOpenAI(apiKey, data):
    """
    Realiza uma chamada à API da OpenAI para corrigir um texto.

    Parâmetros:
    apiKey (str): A chave da API da OpenAI utilizada para autenticação.
    data (dict): O corpo da requisição em formato JSON, contendo os dados para a correção do texto.

    Retorna:
    str: O texto corrigido retornado pela API se a chamada for bem-sucedida, ou uma mensagem de erro em caso de falha.
    """

    # Headers da requisição
    headers = {
        'Authorization': f'Bearer {apiKey}',
        'Content-Type': 'application/json',
    }

    # Realizar chamada da API
    response = requests.post(url, headers=headers, data=json.dumps(data))

    # Verificando o código de retorno
    if response.status_code == 200:
        # Sucesso
        textoCorrigido = response.json()['choices'][0]['message']['content']
        return textoCorrigido
    else:
        # Erro
        print(f"Erro: {response.status_code}")
        return f'{response.status_code} Houve erro na chamada da API da OpenAI - {response.text}'



# Definir variáveis iniciais
apiKey = os.environ['OPENAI_API_KEY']
url = 'https://api.openai.com/v1/chat/completions'
excelPath = r'.\Exemplo1.xlsx'
sheet = 'Plan1'
colunaOutput = 'Output'

# Ler planilha Excel
excelDf = lerArquivoExcel(excelPath, sheet)

# Loop para cada linha do dataframe que contém os dados da planilha Excel
for i, data in excelDf.iterrows():
    # Obter dados da planilha
    textoInput = data["Input"]
    print(f'Texto input: {textoInput}')

    # Montar corpo da request JSON
    data = {
        "model": "gpt-3.5-turbo",
        "messages": [
            {
                "role": "user",
                "content": f"Por favor, corrija o seguinte texto: '{textoInput}' e me retorne somente o texto corrigido como resposta."
            }
        ]
    }

    # Usar função que irá fazer a chamada da API da OpenAI para corrigir o texto 
    textoCorrigido = chamadaApiOpenAI(apiKey, data)
    print(f'Texto corrigido (output): {textoCorrigido}')

    # Verificar se a função não retornou erro
    if 'Houve erro' in textoCorrigido:
        # Retornou erro, exibir mensagem erro e prosseguir para a próxima linha
        print(textoCorrigido)
        continue

    # Chamada da API teve sucesso, atualizar respectiva linha da planilha na coluna Output com o texto corrigido
    inserirTextoCorrigido(excelDf, i, colunaOutput, textoCorrigido, excelPath, sheet)