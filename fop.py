import re
import PyPDF2
import sys
import io
import pandas as pd
import os
import email
from email import policy
from email.parser import BytesParser
from PyPDF2 import PdfReader

def extracao_dados(pasta, nome):
    # Abre o arquivo PDF
    with open(pasta, 'rb') as pdf_file:
        read_pdf = PyPDF2.PdfReader(pdf_file)
        page = read_pdf.pages[0]
        page_content = page.extract_text()

        if page_content is None:
            print("Erro: Não foi possível extrair texto do PDF.")
            return {}

        if isinstance(page_content, bytes):
            page_content = page_content.decode('utf-8')

        # Usando expressões regulares para capturar os dados
        parsed = re.sub(r'\n+', ' ', page_content)  # Substitui quebras de linha por espaço
        parsed = re.sub(r'\s+', ' ', parsed)  # Remove espaços em excesso

        # Definindo os padrões para cada campo
        padrao_nome = re.compile(r'Nome:\s*\(Completo e sem abreviação\)\s*([A-Z\s]+)', re.IGNORECASE)
        padrao_matricula = re.compile(r'Matrícula:\s*\(\s*informar\s*no\s*padrão\s*x\.x{3}\.x{3}-x\s*\)\s*([A-Z\s]+)', re.IGNORECASE)
        padrao_cpf = re.compile(r'CPF:\s*\(informar no padrão.*?\)\s*([\d\s.-]{11,})', re.IGNORECASE)  # Aceita números e separadores
        padrao_lotacao = re.compile(r'Lotação\s*\(Unidade.*?\):\s*([^\n]+)', re.IGNORECASE)  # Captura tudo até uma nova linha
        padrao_data_nascimento = re.compile(r'Data de Nascimento:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE)
        padrao_data_admissao = re.compile(r'Data de Admissão:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE)

        # Buscando os dados
        dados = {
            'Nome': padrao_nome.search(parsed),
            'Matricula': '',
            'CPF': padrao_cpf.search(parsed),
            'Lotação': padrao_lotacao.search(parsed),
            'Data de Nascimento': padrao_data_nascimento.search(parsed),
            'Data de Admissão': padrao_data_admissao.search(parsed)
        }

        # Extraindo e retornando os dados
        resultados = {}
        for campo, correspondencia in dados.items():
            if correspondencia:
                resultados[campo] = correspondencia.group(1).strip()
            else:
                resultados[campo] = None

        # Corrigindo valores não capturados
        if resultados['Nome']:
            nome_parts = resultados['Nome'].split(' Matr')[0]  # Remove qualquer parte desnecessária
            resultados['Nome'] = nome_parts.strip()

        if resultados['Lotação']:
            lotacao_parts = resultados['Lotação'].split(' Data de Nascimento')[0]  # Remove o resto
            resultados['Lotação'] = lotacao_parts.strip()

        # Formatando o CPF no padrão xxx.xxx.xxx-xx
        if resultados['CPF']:
            cpf = resultados['CPF']
            resultados['CPF'] = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"  # Formata o CPF

        resultados['Matricula'] = nome
        # print(resultados)
        
        return resultados

def salvar_em_excel(dados, arquivo_saida):
    df = pd.DataFrame(dados)  # Cria um DataFrame com os dados
    df.to_excel(arquivo_saida, index=False)  # Salva o DataFrame em um arquivo Excel

def extrair_pdfs_do_eml(caminho_eml, pasta_destino):
    # Abre o arquivo .eml
    with open(caminho_eml, 'rb') as arquivo:
        msg = BytesParser(policy=policy.default).parse(arquivo)

        # Procura por anexos
        for anexo in msg.iter_attachments():
            if anexo.get_content_type() == 'application/pdf':
                # Salva o PDF
                nome_arquivo = anexo.get_filename()
                caminho_pdf = os.path.join(pasta_destino, nome_arquivo)
                with open(caminho_pdf, 'wb') as f:
                    f.write(anexo.get_payload(decode=True))
                print(f'PDF extraído: {caminho_pdf}')

def main(pasta_emails, pasta_destino):
    # Cria a pasta de destino, se não existir
    
    # os.makedirs(pasta_destino, exist_ok=True)

    # Percorre todos os arquivos na pasta
    for nome_arquivo in os.listdir(pasta_emails):
        if nome_arquivo.endswith('.eml'):
            caminho_eml = os.path.join(pasta_emails, nome_arquivo)
            extrair_pdfs_do_eml(caminho_eml, pasta_destino)

def processar_pdfs(pasta_pdf, arquivo_saida):
    dados_total = []
    # Percorre todos os arquivos na pasta
    for nome_arquivo in os.listdir(pasta_pdf):
        if nome_arquivo.endswith('.pdf'):
            caminho_pdf = os.path.join(pasta_pdf, nome_arquivo)
            print(f'Processando: {caminho_pdf}')

            # Executa a função de extração de dados
            dados_extraidos = extracao_dados(caminho_pdf, nome_arquivo)
            
            if dados_extraidos:  # Verifica se há dados extraídos
                dados_total.append(dados_extraidos)

    
    if dados_total:  # Verifica se há dados para salvar
        salvar_em_excel(dados_total, arquivo_saida)
        print(f'Dados salvos em: {arquivo_saida}')

def choice():

    opcao = input('Escolha uma opção: \n Opção 1: Pegar pdfs do zip \n Opção 2: So extrair dados dos pdfs \n Escolha: 1 ou 2: \n')

    if opcao == '1':

        pasta_origem = r'Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\01 - Eventual\JOVENS APRENDIZ\pdfs'
        pasta_destino = r'Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\01 - Eventual\JOVENS APRENDIZ\data'
        #main(pasta_origem, pasta_destino)

        arquivo_saida = r'Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\01 - Eventual\JOVENS APRENDIZ\dados_empregado.xlsx'

        # init(pasta_origem)
        processar_pdfs(pasta_origem, arquivo_saida)

    if opcao == '2':
        pasta_origem = r'Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\01 - Eventual\JOVENS APRENDIZ\pdfs'
        pasta_destino = r'Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\01 - Eventual\JOVENS APRENDIZ\data'
        main(pasta_origem, pasta_destino)

        arquivo_saida = r'Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\01 - Eventual\JOVENS APRENDIZ\dados_empregado.xlsx'

        processar_pdfs(pasta_origem, arquivo_saida)


choice()