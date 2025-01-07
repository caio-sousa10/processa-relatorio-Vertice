import pandas as pd
from tkinter import Tk, filedialog
import win32com.client as win32
from datetime import datetime
import os


def carregar_arquivo(path_arquivo, tipo="relatório de exames"):
    try:
        print(f'Carregando o arquivo {tipo}...')
        dados = pd.read_excel(path_arquivo, dtype=str, header=2)  # Lendo a partir da 3ª linha
        return dados
    except Exception as e:
        print(f"Erro ao carregar o arquivo {tipo}: {e}")
        exit()


def classificar_tipo_de_exame(nome):
    nome = str(nome).lower()
    if 'audiometria' in nome:
        return 'Audiometria'
    elif 'anamnese' in nome:
        return 'Anamnese'
    elif 'acuidade' in nome or 'acuidade visual' in nome:
        return 'Acuidade visual'
    elif 'exame físico' in nome:
        return 'Exame físico'
    else:
        return 'Exame(s) não realizado(s)'


def gerar_relatorio(dados, empresa=None):
    print('Filtrando exames pendentes...')
    if empresa:
        dados = dados[dados['Nome Empresa'] == empresa]  # Filtro para a empresa específica
    exames = dados[dados['Refazer em'].notna()]
    if exames.empty:
        print("Nenhum exame pendente encontrado.")
        exit()

    print('Gerando o relatório organizado...')
    relatorio = exames.groupby('Tipo do Exame').size().reset_index(name='Quantidade')
    return relatorio


def salvar_relatorio(relatorio, empresa):
    data_atual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    nome_saida = f'C:/SOC_Relatorios/relatorio_{empresa}_{data_atual}.xlsx'

    # Verificar se o diretório existe
    if not os.path.exists(os.path.dirname(nome_saida)):
        os.makedirs(os.path.dirname(nome_saida))

    relatorio.to_excel(nome_saida, index=False)
    print(f'Relatório salvo em: {nome_saida}')
    return nome_saida


def enviar_email_com_anexo(arquivo, destinatarios, assunto, mensagem):
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)  # 0 corresponde a um e-mail

        mail.Subject = assunto
        mail.Body = mensagem
        mail.To = ";".join(destinatarios)

        mail.Attachments.Add(arquivo)

        # Salvar como rascunho
        mail.Save()
        print("E-mail salvo como rascunho!")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")


def obter_emails(path_arquivo_emails):
    try:
        print("Carregando os e-mails...")        
        emails_df = pd.read_excel(path_arquivo_emails, header=2)  # Lendo a partir da 3ª linha

        # Verifica as colunas no DataFrame
        print("Colunas encontradas no arquivo de e-mails:", emails_df.columns)

        # Certifique-se de que as colunas que você deseja usar existem
        if 'Nome Empresa' not in emails_df.columns or 'E-mail 1' not in emails_df.columns or 'E-mail 2' not in emails_df.columns:
            print("As colunas 'Nome Empresa', 'E-mail 1' ou 'E-mail 2' não estão presentes no arquivo.")
            exit()

        # Remover espaços extras nos nomes das colunas
        emails_df.columns = emails_df.columns.str.strip()

        # Extraindo os e-mails de ambas as colunas
        emails_1 = emails_df['E-mail 1'].dropna().tolist()  # E-mails da coluna 'E-mail 1'
        emails_2 = emails_df['E-mail 2'].dropna().tolist()  # E-mails da coluna 'E-mail 2'

        # Unindo os e-mails e removendo duplicatas
        emails = list(set(emails_1 + emails_2))  # Garantir que não haja e-mails duplicados

        if not emails:
            print("Nenhum e-mail encontrado na planilha.")
            exit()

        return emails
    except Exception as e:
        print(f"Erro ao carregar os e-mails: {e}")
        exit()


def main():
    # Passo 1: Carregar a planilha de e-mails e empresas
    print("Selecione a planilha com os e-mails das empresas.")
    Tk().withdraw()
    path_arquivo_emails = filedialog.askopenfilename(title="Selecione o arquivo de e-mails", filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
    if not path_arquivo_emails:
        print("Nenhum arquivo de e-mails foi selecionado. O programa será encerrado.")
        exit()

    # Carregar e-mails
    emails_df = carregar_arquivo(path_arquivo_emails, tipo="planilha de empresas e e-mails")
    
    # Passo 2: Carregar o arquivo de exames (relatório do SOC)
    print("Selecione o arquivo de exames.")
    path_arquivo = filedialog.askopenfilename(title="Selecione o relatório de exames", filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
    if not path_arquivo:
        print("Nenhum arquivo foi selecionado. O programa será encerrado.")
        exit()

    dados = carregar_arquivo(path_arquivo, tipo="relatório de exames")
    dados.columns = ['Empresa', 'Unidade', 'Nome', 'Cargo', 'Setor', 'Exame', 'Refazer em']
    dados = dados.iloc[1:].reset_index(drop=True)  # Ajusta os dados, removendo a primeira linha com informações inúteis

    # Renomeia a coluna 'Empresa' para 'Nome Empresa' para manter consistência com a planilha de e-mails
    dados.rename(columns={'Empresa': 'Nome Empresa'}, inplace=True)

    dados['Tipo do Exame'] = dados['Exame'].apply(classificar_tipo_de_exame)

    # Passo 3: Solicitar a empresa ao usuário e filtrar os dados
    empresas = emails_df['Nome Empresa'].unique().tolist()
    empresa = input(f"Escolha uma das empresas: {empresas}\nDigite o nome da empresa: ")

    if empresa not in empresas:
        print("Empresa não encontrada! Programa encerrado.")
        exit()

    # Filtra os exames para a empresa
    dados_filtrados = gerar_relatorio(dados, empresa)
    nome_arquivo = salvar_relatorio(dados_filtrados, empresa)

    # Passo 4: Obter os e-mails para a empresa selecionada
    emails_empresa = emails_df[emails_df['Nome Empresa'] == empresa]
    emails = list(set(emails_empresa['E-mail 1'].dropna().tolist() + emails_empresa['E-mail 2'].dropna().tolist()))

    if not emails:
        print(f"Nenhum e-mail encontrado para a empresa {empresa}.")
        exit()

    # Passo 5: Enviar e-mail com o anexo
    assunto = f"Relatório de Exames Pendentes - {empresa}"
    mensagem = f"Olá, segue o relatório de exames pendentes da empresa {empresa}. Qualquer dúvida, estou à disposição."

    enviar_email_com_anexo(nome_arquivo, emails, assunto, mensagem)


if __name__ == "__main__":
    main()

# SUGESTAO PARA FAZER : PROCURAR EMPRESA COM BASE NA SIMILARIDADE DO TEXTO DIGITADO 
# AO INVES DE TER QUE DIGITAR IGUAL