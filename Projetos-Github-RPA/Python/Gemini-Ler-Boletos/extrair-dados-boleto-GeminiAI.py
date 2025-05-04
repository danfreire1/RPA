"""
Projeto: Extração de Dados de Boletos com a API Gemini
Descrição: Este projeto realiza a extração automatizada de informações de boletos bancários a partir de arquivos PDF. Utilizando a API Gemini, os dados extraídos são organizados e inseridos em uma planilha Excel para posterior análise e utilização.
Autor: Daniel Freire da Costa
Data de Criação: 15 de Dezembro de 2024
Versão: 1.0
Dependências: 
    - google-generativeai
    - openpyxl
    - os
    - json
    - pathlib

Instruções de Uso:
1. Configure a chave da API Gemini como uma variável de ambiente chamada `GEMINI_API_KEY`.
2. Certifique-se de que os boletos estão localizados no diretório `./arquivos`.
3. Configure o arquivo Excel `output.xlsx` com o cabeçalho necessário (ou deixe que o script o crie automaticamente).
4. Execute o script para processar os boletos e extrair as informações.
5. Os dados serão adicionados ao arquivo Excel, incluindo os campos:
    - Beneficiário
    - Agência/Código do Beneficiário
    - Nosso Número
    - Valor do Documento
    - Data de Vencimento
    - Pagador
    - Endereço
    - CPF
    - Código de Barras
    - Nome do Arquivo PDF

Referências:
1. Boletos utilizados nos testes foram gerados automaticamente pelo ChatGPT.
2. Documentação do Gemini: https://ai.google.dev/gemini-api/docs/document-processing?hl=en&lang=python

"""

# Libs
import google.generativeai as genai
import os
import json
from pathlib import Path
from openpyxl import load_workbook


def inserir_dados_excel(dadosRetornoGemini, arquivo_excel, nome_pdf):
    try:
        workbook = load_workbook(arquivo_excel)
        sheet = workbook.active

        nova_linha = [
            dadosRetornoGemini.get('Beneficiário', ''),
            dadosRetornoGemini.get('Agência/Código do Beneficiário', ''),
            dadosRetornoGemini.get('Nosso Número', ''),
            dadosRetornoGemini.get('Valor do Documento', ''),
            dadosRetornoGemini.get('Data de Vencimento', ''),
            dadosRetornoGemini.get('Pagador', ''),
            dadosRetornoGemini.get('Endereço', ''),
            dadosRetornoGemini.get('CPF', ''),
            dadosRetornoGemini.get('Código de Barras', ''),
            nome_pdf
        ]

        sheet.append(nova_linha)
        workbook.save(arquivo_excel)
        print(f"Dados do arquivo {nome_pdf} inseridos com sucesso.")

    except Exception as e:
        print(f"Erro ao inserir dados no Excel: {e}")

# Variaveis iniciais
apiKey = os.getenv('GEMINI_API_KEY')
arquivoExcel = r'.\output.xlsx'
diretorioArquivos = './arquivos'

# Configurar API Gemini e inicializar
genai.configure(api_key=apiKey)
model = genai.GenerativeModel('gemini-1.5-flash')

# Loop para cada arquivo na pasta arquivos
for nomeArquivo in os.listdir(diretorioArquivos):
    # Obter caminho do arquivo
    caminhoBoleto = os.path.join(diretorioArquivos, nomeArquivo)

    # Fazer upload do arquivo para o Gemini
    arquivo = genai.upload_file(path=caminhoBoleto, display_name=Path(nomeArquivo).stem)

    # Confirmar o upload feito
    print(f'Arquivo enviado: {arquivo.display_name} - url: {arquivo.uri}')

    # Realizar extração dos dados do pdf enviado
    response = model.generate_content(['Me retorne em json as seguintes informações do arquivo pdf: Beneficiário, Agência/Código do Beneficiário, Nosso Número, Valor do Documento, Data de Vencimento, Pagador, Endereço, CPF e Código de Barras. Não precisa escrever json antes, só retorne no formato de um.', arquivo])

    print(response.text)

    # Tratar retorno do Gemini (O resultado está em response.text, porém mesmo explicando a ele para somente trazer a resposta em formato json, ele traz ela com o texto json antes. Portanto, aplico replace e depois uso a lib json.loads para transformar em um json e poder manipular o objeto criado)
    respostaFormatada = json.loads(response.text.replace('```json', '').replace('```', ''))
    print(respostaFormatada)

    # Chamar função para inserir dados no Excel
    inserir_dados_excel(respostaFormatada, arquivoExcel, nomeArquivo)