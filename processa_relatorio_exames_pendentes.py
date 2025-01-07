# Protótipo de processamento de relatório de exames pendentes
# exportado do SOC de forma organizada para excel
# Aumenta produtividade (menos tempo que manualmente) e reduz possíveis erros humanos


# Instalar pandas via terminal antes comando: pip install pandas openpyxl

import pandas as pd
from datetime import datetime

# Verificar com Bruno o caminho do arquivo, tem que estar em formato planilha excel (extensao .xlsx)
path_arquivo = r'C:\SOC_Relatorios\exames_pendentes.xlsx' #Exemplo 

# Tratamento de exceção para lidar com eventuais erros
try:
    print('Carregando o arquivo de exames...')
    dados = pd.read_excel(path_arquivo)
except FileNotFoundError:
    print('Arquivo não encontrado! Verifique se o arquivo foi exportado corretamente para o caminho...')
    exit()

# Filtramento de exames pendentes // modificar caso necessário
print('Filtrando...')
exames = dados[dados['status'] == 'Pendente']

# Verifica se há exames pendentes
if exames.empty:
    print("Nenhum exame pendente encontrado.")
    exit()

# Organização por tipo e quantidade
print('Gerando o relatório organizado...')
relatorio = exames.groupby('Tipo do Exame').size().reset_index(name = 'Quantidade')

# Sáida de arquivo com nome e data
data_atual = datetime.now().strftime('%Y-%m-%d')
nome_saida = rf'C:\SOC_Relatorios\relatorio_formatado_{data_atual}.xlsx' #Exemplo

# Salvando o relatório
relatorio.to_excel(nome_saida, index = False) # Excluindo indice do canto esquerdo excel
print(f'Relatório salvo em: {nome_saida}') # Utilizando f string para colocar valor da variável na própria string

###################### SOLICITAR FEEDBACK DO BRUNO PARA MELHORIAS E SUGESTOES ####################################