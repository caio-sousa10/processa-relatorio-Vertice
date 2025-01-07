Funcionalidades

	Leitura e processamento de arquivos Excel exportados do SOC.
	Filtragem de exames pendentes por empresa.
	Agrupamento e contagem por tipo de exame.
	Geração automática de relatório com a data e hora no nome do arquivo.
	Criação de diretórios de saída caso necessário.
	Tratamento de erros para garantir que o processo seja robusto e sem falhas.

Tecnologias Utilizadas

	Python 3.x
	Pandas: para manipulação e análise dos dados.
	Tkinter: para abrir caixas de diálogo e selecionar arquivos.
 	Win32com (pywin32): Para automação do Outlook e envio de e-mails.
  	Datetime: Para manipular datas e gerar nomes de arquivos.
   	OS: Para lidar com diretórios e salvar arquivos no sistema.
 
Como Usar

	Certifique-se de ter as bibliotecas necessárias instaladas:
 
 	1️⃣ pip install pandas pywin32
	
	2️⃣ Execute o script	
 	
 	3️⃣ Selecione os arquivos quando solicitado
  
	4️⃣ Digite o nome da empresa desejada para gerar o relatório.

	5️⃣ O script gera um relatório formatado e salva um e-mail no Outlook com o anexo.

	📩 O e-mail será salvo como rascunho para conferência antes do envio.
