# SharepointerSelenium

## Descrição

`SharepointerSelenium` é uma biblioteca Python que facilita o download e upload de arquivos no SharePoint, especialmente útil em situações onde a API não está disponível. Ele automatiza o login, lida com verificações de credenciais e realiza operações de download e upload, com suporte para gerenciar conflitos de nome de arquivo (como sobrescrever ou manter ambos).

## Funcionalidades

- **Login Automatizado:** A classe gerencia o login no SharePoint, solicitando credenciais (caso não fornecidas), validando e corrigindo e-mails ou senhas incorretas.
- **Download de Arquivos:** Faz o download de arquivos de uma pasta no SharePoint, com controle de tempo de espera e verificações de sucesso.
- **Upload de Arquivos:** Suporta o upload de um ou mais arquivos, com a opção de sobrescrever arquivos de mesmo nome ou manter ambos.
- **Gestão de Conflitos de Arquivos:** Quando um arquivo já existe no SharePoint, permite que o usuário escolha se deseja substituir o arquivo existente ou manter ambos.

## Pré-requisitos
<table>
  <thead align="center">
    <tr border: none;>
      <td><b>VERSÃO PYTHON</b></td>
      <td><b>BIBLIOTECAS EXTERNAS</b></td>
      <td><b>WEBDRIVER</b></td>
    </tr>
  </thead>
  <tbody align="center">
    <tr>
      <td>
        <img alt="" src="https://img.shields.io/badge/python-V3.7%2B-black?style=flat-square&logo=python&labelColor=yellow"/>
      </td>
      <td>
        <img alt="" src="https://img.shields.io/badge/selenium-V4.24.0-black?style=flat-square&logo=selenium&logoColor=white&labelColor=%2359b943"/>
      </td>
      <td>
        <img alt="" src="https://img.shields.io/badge/edge%20webdriver-128.0.2739.79-black?style=flat-square&labelColor=blue"/>
      </td>
    </tr>
  </tbody>
</table>

### Instalação do WebDriver

Você pode baixar o WebDriver do Edge [aqui](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/?form=MA13LH).

## Instalação

Clone o repositório e instale as dependências necessárias:

```bash
git clone https://github.com/RonanBenitis/sharepointer_selenium.git
cd sharepointer_selenium
pip install -r requirements.txt
```

Certifique-se de ter o **WebDriver do Edge** instalado.

## Exemplo de Uso

```python
from sharepointer_selenium import SharepointerSelenium

# Configurações iniciais
sharepoint_url = 'https://nomedosharepoint.sharepoint.com'
username = 'seu-email@dominio.com'
password = 'sua-senha'

# Instanciando o objeto Sharepointer
sharepoint = SharepointerSelenium(sharepoint_url, username, password)

# Realizar download de um arquivo
sharepoint.download_file('https://nomedosharepoint.sharepoint.com/pasta', 'documento.xlsx')

# Realizar upload de arquivos
sharepoint.upload_file(['arquivo1.txt', 'arquivo2.txt'], 'https://nomedosharepoint.sharepoint.com/pasta')
```

## Parâmetros da Classe

### `__init__(self, sharepoint_url: str, username: str = None, password: str = None, download_dir: str = None)`

- `sharepoint_url`: URL base do SharePoint.
- `username` (opcional): Seu e-mail para login no SharePoint.
  - Caso username não for passado via código, o programa solicitará via terminal
- `password` (opcional): Sua senha para login no SharePoint.
  - Caso password não for passado via código, o programa solicitará via terminal
  - A senha **não será replicada no terminal**, ou seja, aparentará não estar sendo preenchida, mas, o código estará registrando o preenchimento até o apertar do Enter 
- `download_dir` (opcional): Caminho para a pasta de download. Padrão é uma pasta `downloads` no diretório do código.

### Métodos principais

#### `download_file(sharepoint_folder_url: str, file_name_with_ext: str, wait_download_time: int = 60)`

Realiza o download de um arquivo específico do SharePoint.

- `sharepoint_folder_url`: URL da pasta onde o arquivo está.
- `file_name_with_ext`: Nome do arquivo (incluindo a extensão) que deseja baixar.
- `wait_download_time` (opcional): Tempo máximo de espera para o download, em segundos (padrão: 60).

#### `upload_file(local_files_path: str | WindowsPath (pathlib) | list, sharepoint_folder_url: str, replace: bool = None, wait_upload_time: int = 60)`

Faz o upload de um ou mais arquivos para o SharePoint.

- `local_files_path`: Caminho do(s) arquivo(s) local(is).
- `sharepoint_folder_url`: URL da pasta onde o upload será feito.
- `replace` (opcional): Define se deve sobrescrever ou manter ambos os arquivos ao encontrar um nome duplicado.

## Observações

- Caso opte por utilizar as credenciais via código, é recomendável utilizar variáveis de ambiente para definir a senha (.env), por questões de segurança.
- Caso as credenciais não sejam válidas, o programa solicitará novos dados para o login.
- Para mais informações, as principais funcionalidades do código estão documentadas, também, via docstring.
- Esta aplicação aceita parametros utilizando pathlib para upload, seja em lista ou seja unitariamente.
