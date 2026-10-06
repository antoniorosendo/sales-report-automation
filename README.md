# Email Automation Python

Script em Python que lê uma lista de contatos em uma planilha Excel e envia um e-mail personalizado, com anexo, para cada pessoa da lista.

O projeto nasceu de um caso prático: enviar o painel de vendas de uma papelaria (dados fictícios) para os funcionários, pedindo uma análise dos resultados.

## Funcionalidades

- Lê nome, e-mail e setor de uma planilha `.xlsx` com `openpyxl`.
- Envia uma mensagem personalizada para cada contato (nome e setor aparecem no texto).
- Anexa um arquivo (o painel de vendas em Excel).
- Usa uma única conexão com o Gmail para todos os envios.
- Se um envio falhar, continua para o próximo e mostra no final quem falhou.
- Guarda a senha fora do código, em um arquivo `.env`.

## Tecnologias

- Python 3
- [openpyxl](https://openpyxl.readthedocs.io/): leitura e criação de planilhas
- `smtplib` e `email` (biblioteca padrão): envio de e-mails
- [python-dotenv](https://pypi.org/project/python-dotenv/): leitura de variáveis do arquivo `.env`

## Estrutura do projeto

```
.
├── enviar_emails.py        # script principal de envio
├── criar_planilha.py       # gera a planilha de contatos fictícios
├── funcionarios.xlsx       # lista de funcionarios (dados fictícios)
├── painel_vendas.xlsx      # arquivo enviado como anexo
├── .env                    # modelo do arquivo .env
├── .gitignore
└── README.md
```

## Como usar

### 1. Clonar o repositório e instalar as dependências

```bash
git clone https://github.com/antoniorosendo/sales-report-automation.git
cd sales-report-automation
pip install openpyxl python-dotenv
```

### 2. Gerar uma senha de app do Gmail

O Gmail não aceita a senha normal da conta em scripts. É preciso:

1. Ativar a verificação em duas etapas na conta Google.
2. Acessar **Conta Google → Segurança → Senhas de app** e gerar uma senha.

### 3. Criar o arquivo `.env`

Copie o modelo e preencha com a sua senha de app:

```bash
cp .env.example .env
```

Conteúdo do `.env` (sem aspas e sem espaços ao redor do `=`):

```
GMAIL_APP_PASSWORD=sua_senha_de_app_aqui
```

> O `.env` está no `.gitignore` e **não** vai para o repositório.

### 4. Configurar o script

No começo do `enviar_emails.py`, ajuste:

```python
REMETENTE = "seu.email@gmail.com"
ANEXO = PASTA / "painel_vendas.xlsx"   # ou None para enviar sem anexo
ASSUNTO = "Painel de vendas da papelaria"
```

### 5. Preparar a lista de contatos

O arquivo `funcionarios.xlsx` precisa ter este formato, com cabeçalho na primeira linha:

| Nome | Email | Setor |
|---|---|---|
| Ana Souza | ana.souza@example.com | Vendas |
| Bruno Lima | bruno.lima@example.com | Marketing |

Os e-mails do repositório são fictícios (`@example.com`) e **não recebem mensagens**. Para testar, troque por e-mails seus.

### 6. Executar

```bash
python enviar_emails.py
```

Exemplo de saída:

```
8 contatos encontrados.
Enviado para Ana Souza <ana.souza@example.com>
...
Concluído: 8 enviados, 0 falhas.
```

## Observações

- O Gmail limita o envio a cerca de 500 e-mails por dia em contas comuns. Este projeto não é indicado para disparo em massa.
- O script espera uma pausa de 2 segundos entre os envios (`PAUSA_ENTRE_ENVIOS`) para não parecer spam.
- Todos os dados do projeto (vendas e contatos) são fictícios.

## Segurança

- Nunca coloque a senha direto no código.
- Nunca faça commit do `.env`.
- Se a senha de app vazar, revogue em **Conta Google → Segurança → Senhas de app** e gere outra.

## Autor

Antonio Carlos Rosendo da Silva