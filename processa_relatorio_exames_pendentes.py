import pandas as pd
from tkinter import Tk, filedialog
import win32com.client as win32
from datetime import datetime
import os
import warnings

# Ignorando erro de estilo no terminal
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def carregar_arquivo(path_arquivo, tipo="relatório de exames"):
    """
    Lê um arquivo Excel e retorna seu conteúdo como um DataFrame do Pandas.
    
    Parâmetros:
    - path_arquivo: Caminho do arquivo a ser carregado.
    - tipo: Tipo do arquivo (para mensagens de erro).
    
    Retorno:
    - DataFrame com os dados do arquivo.
    """
    try:
        print(f'Carregando o arquivo {tipo}...')
        dados = pd.read_excel(path_arquivo, dtype=str, header=2)  # Lendo a partir da 3ª linha
        return dados
    except Exception as e:
        print(f"Erro ao carregar o arquivo {tipo}: {e}")
        exit()

def classificar_tipo_de_exame(nome):
    """
    Classifica o tipo de exame com base no nome fornecido.
    
    Parâmetro:
    - nome: Nome do exame.
    
    Retorno:
    - Nome padronizado do exame.
    """
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
    """
    Filtra exames pendentes e gera um relatório para uma empresa específica (se informada).
    
    Parâmetros:
    - dados: DataFrame contendo os exames.
    - empresa: Nome da empresa para filtrar os dados.
    
    Retorno:
    - DataFrame com o relatório de exames pendentes.
    """
    print('Filtrando exames pendentes...')
    if empresa:
        dados = dados[dados['Nome Empresa'] == empresa]  # Filtra apenas a empresa desejada
    exames = dados[dados['Refazer em'].notna()]  # Filtra apenas exames que precisam ser refeitos
    if exames.empty:
        print("Nenhum exame pendente encontrado.")
        exit()

    print('Gerando o relatório organizado...')
    relatorio = exames.groupby('Tipo do Exame').size().reset_index(name='Quantidade')
    return relatorio

def salvar_relatorio(relatorio, empresa):
    """
    Salva o relatório gerado em um arquivo Excel dentro da pasta 'C:/SOC_Relatorios/'.
    
    Parâmetros:
    - relatorio: DataFrame contendo os dados do relatório.
    - empresa: Nome da empresa para nomear o arquivo.
    
    Retorno:
    - Caminho do arquivo salvo.
    """
    data_atual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    nome_saida = f'C:/SOC_Relatorios/relatorio_{empresa}_{data_atual}.xlsx'

    # Criar pasta caso não exista
    os.makedirs(os.path.dirname(nome_saida), exist_ok=True)

    relatorio.to_excel(nome_saida, index=False)
    print(f'Relatório salvo em: {nome_saida}')
    return nome_saida

def enviar_email_com_anexo(arquivo, destinatarios, assunto, mensagem):
    """
    Envia um e-mail via Outlook com um anexo e uma assinatura de imagem.
    
    Parâmetros:
    - arquivo: Caminho do arquivo anexo.
    - destinatarios: Lista de e-mails para envio.
    - assunto: Assunto do e-mail.
    - mensagem: Texto principal do e-mail.
    """
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)

        mail.Subject = assunto
        mail.To = ";".join(destinatarios)

        assinatura_path = r"C:\Users\Vértice TI - Marília\Pictures\Assinatura de email caio.jpg" # r string para aceitar barra invertida

        mail.HTMLBody = f"""
        <p>{mensagem}</p>
        <img src="cid:assinatura">
        """

        assinatura_anexo = mail.Attachments.Add(assinatura_path)
        assinatura_anexo.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "assinatura")

        mail.Attachments.Add(arquivo)
        mail.Save()
        print("E-mail salvo como rascunho!")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")

def main():
    """Fluxo principal do programa."""
    
    # Seleção do arquivo de e-mails
    print("Selecione a planilha com os e-mails das empresas.")
    Tk().withdraw()
    path_arquivo_emails = filedialog.askopenfilename(title="Selecione o arquivo de e-mails", filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
    if not path_arquivo_emails:
        print("Nenhum arquivo de e-mails foi selecionado. O programa será encerrado.")
        exit()

    emails_df = carregar_arquivo(path_arquivo_emails, tipo="planilha de empresas e e-mails")

    # Seleção do arquivo de exames
    print("Selecione o arquivo de exames.")
    path_arquivo = filedialog.askopenfilename(title="Selecione o relatório de exames", filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
    if not path_arquivo:
        print("Nenhum arquivo foi selecionado. O programa será encerrado.")
        exit()

    dados = carregar_arquivo(path_arquivo, tipo="relatório de exames")
    dados.columns = ['Empresa', 'Unidade', 'Nome', 'Cargo', 'Setor', 'Exame', 'Refazer em']
    dados = dados.iloc[1:].reset_index(drop=True)

    dados.rename(columns={'Empresa': 'Nome Empresa'}, inplace=True)
    dados['Tipo do Exame'] = dados['Exame'].apply(classificar_tipo_de_exame)

    # Detecção automática da empresa
    if 'Nome Empresa' in dados.columns:
        empresas_unicas = dados['Nome Empresa'].dropna().unique()

        if len(empresas_unicas) == 1:
            empresa = empresas_unicas[0]
            print(f"Empresa detectada automaticamente: {empresa}")
        else:
            empresa = dados['Nome Empresa'].value_counts().idxmax()
            print(f"Empresa mais frequente detectada: {empresa}")
    else:
        print("Erro: A coluna 'Nome Empresa' não foi encontrada no relatório.")
        exit()

    # Gerar e salvar relatório
    dados_filtrados = gerar_relatorio(dados, empresa)
    nome_arquivo = salvar_relatorio(dados_filtrados, empresa)

    # Obter e-mails da empresa
    emails_empresa = emails_df[emails_df['Nome Empresa'] == empresa]
    emails = list(set(emails_empresa['E-mail 1'].dropna().tolist() + emails_empresa['E-mail 2'].dropna().tolist()))

    if not emails:
        print(f"Nenhum e-mail encontrado para a empresa {empresa}.")
        exit()

    # Enviar e-mail com o relatório
    assunto = f"Relatório de Exames Pendentes - {empresa}"
    mensagem = f"Olá, segue o relatório de exames pendentes da empresa {empresa}. Qualquer dúvida, estou à disposição. <br> <br> At.te,"

    enviar_email_com_anexo(nome_arquivo, emails, assunto, mensagem)

if __name__ == "__main__":
    main()
