Funcionalidades principais
Interface Web:

Página inicial com formulário de upload para arquivos PDF.
Página de resultado com link para download do arquivo processado.
Processamento de PDFs:

Conversão de PDFs para Markdown usando a biblioteca docling.
Extração de tabelas formatadas e informações específicas do conteúdo do PDF.
Tratamento de Dados:

Ajustes nas tabelas extraídas, como renomeação de colunas, remoção de linhas desnecessárias e preenchimento de valores ausentes.
Organização dos dados em um DataFrame pandas.
Exportação para Excel:

Os dados extraídos são consolidados em uma planilha Excel.
Aplicação de formatação, como cabeçalhos estilizados e formatação de tabela com estilos pré-definidos.
Sistema de Arquivos:

Diretórios separados para uploads de arquivos e saída dos arquivos processados.
Criação automática de pastas uploads e outputs, caso não existam.
Download de Arquivos:

Permite que o usuário baixe o arquivo Excel gerado após o processamento.
Dependências
Bibliotecas Python:
Flask
pandas
numpy
openpyxl
re
docling
Estrutura de Arquivos e Diretórios
uploads/: Diretório para armazenar os PDFs enviados pelos usuários.
outputs/: Diretório onde os arquivos Excel processados são salvos.
templates/: Diretório contendo os arquivos HTML usados pela aplicação (não listado no código, mas necessário).
