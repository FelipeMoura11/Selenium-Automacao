# Automação de Login, Download e Processamento de Arquivo com Selenium e SQL

Este projeto implementa um script em Python que utiliza o Selenium para automatizar o login em um sistema web, aplicar filtros, realizar uma pesquisa, baixar um arquivo `.xls`, e processá-lo para inserção dos dados extraídos em um banco de dados SQL Server.

## Funcionalidades

- **Automação de Login**: O script realiza login automaticamente em um sistema web utilizando Selenium.
- **Fechamento de Pop-ups**: Após o login, qualquer pop-up que surgir é automaticamente fechado.
- **Navegação e Filtros**: Navega até a página de pesquisa, aplica filtros e executa a pesquisa de forma automática.
- **Download de Arquivo**: Após a pesquisa, o arquivo resultante (em formato `.xls`) é baixado.
- **Processamento de Arquivo**: O arquivo baixado é processado usando BeautifulSoup para extrair dados de uma tabela.
- **Integração com Banco de Dados**: Os dados extraídos são inseridos em um banco de dados SQL Server usando `pyodbc`.

## Tecnologias Utilizadas

- **Python**
- **Selenium**: Automação de tarefas no navegador.
- **BeautifulSoup**: Processamento de arquivos HTML.
- **PyODBC**: Integração com o banco de dados SQL Server.
- **WebDriver Manager**: Gerenciamento automático do WebDriver do Chrome.

## Requisitos
- Selenium
- WebDriver Manager
- BeautifulSoup
- PyODBC
- Driver ODBC para SQL Server

