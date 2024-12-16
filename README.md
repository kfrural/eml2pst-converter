# Conversor de Arquivos EML para PST

Este projeto fornece um script Python para converter grandes volumes de arquivos EML (Mensagens Eletrônicas) em formato PST (Personal Storage Table), compatível com o Microsoft Outlook.

## Funcionalidades Principais

- Converte múltiplos arquivos EML em um único arquivo PST
- Suporta conversão de até 1000 arquivos por execução
- Preserva cabeçalhos importantes como Assunto, Remetente e Destinatário
- Inclui tratamento de erros e logs para monitorar o progresso

## Requisitos

- Python 3.6+
- Biblioteca `win32com.client` (para interação com o Outlook)
- Biblioteca `email` (parte padrão do Python)

## Instalação

1. Clone o repositório:
   ```
   git clone https://github.com/seu-usuario/converter-eml-para-pst.git
   ```

2. Instale as dependências:
   ```
   pip install pywin32
   ```

## Como Usar

1. Certifique-se de ter instalado o Microsoft Outlook na sua máquina
2. Execute o script principal:
   ```
   python main.py
   ```

## Configuração

O script lida automaticamente com a maioria das configurações, mas você pode ajustar alguns parâmetros:

- Crie um arquivo `.env` na raiz do projeto com as seguintes variáveis:
  ```
  EMFOLDER_PATH=/caminho/para/seus/arquivos/eml
  ```

