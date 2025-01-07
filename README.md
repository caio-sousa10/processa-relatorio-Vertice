Funcionalidades

	Leitura e processamento de arquivos Excel exportados do SOC.
	Filtragem de exames com o status "Pendente".
	Agrupamento e contagem por tipo de exame.
	Geração automática de relatório com a data no nome do arquivo.
	Criação de diretórios de saída caso necessário.
	Tratamento de erros para garantir que o processo seja robusto e sem falhas.

Tecnologias Utilizadas

	Python 3.x
	Pandas: para manipulação e análise dos dados.
	OpenPyXL: para exportação dos dados para o formato Excel.
 
Como Usar

	Certifique-se de ter as bibliotecas necessárias instaladas:
 
 	pip install pandas openpyxl
	
	Configure o caminho do arquivo Excel com os dados de exames pendentes.
	Execute o script e o relatório será gerado automaticamente no diretório configurado.
