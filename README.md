# Automação de Leitura de Planilhas (GECEN)

Este projeto contém scripts para automatizar a leitura da planilha de "Indicadores Controle Mensal" e gerar um novo arquivo Excel formatado com dados de Dimensão e Fato.

## Arquivos do Projeto

* **dimensao.py**: Lê os dados da planilha original e gera a aba "DIMENSÃO" com informações de ligações totais e ativas.
* **fato.py**: Lê os dados da planilha original e gera a aba "FATO" com o histórico mensal, volume macromedido, micromedido, perdas e IDF.
* **requirements.txt**: Lista de ferramentas necessárias para o projeto funcionar.

## Pré-requisitos e Instalação

Siga os passos abaixo cuidadosamente se for a primeira vez que executa scripts em Python.

### Passo 1: Se não tiver o Python instalado na máquina 

1.  Acesse o site oficial do Python (python.org) e baixe a versão mais recente.
2.  Ao iniciar a instalação, é fundamental marcar a opção que diz **"Add Python to PATH"** (Adicionar Python ao PATH) antes de clicar em instalar. Se não marcar essa opção, os comandos não funcionarão no terminal.
3.  Conclua a instalação.

### Passo 2: Preparar a Pasta

1.  Crie uma pasta no seu computador.
2.  Coloque os arquivos `dimensao.py`, `fato.py` e `requirements.txt` dentro dessa pasta.
3.  Coloque também a planilha original do Excel dentro desta mesma pasta.
    * **Importante:** O código foi programado para ler um arquivo com o nome padrão: `CONTROLE MENSAL 2024.xlsx`.
    * **Como usar um nome diferente:** Se o seu arquivo tiver outro nome, nos arquivos `dimensao.py` e `fato.py` procure onde está escrito o nome do arquivo antigo (entre aspas) e substitua pelo nome exato do seu arquivo novo. Salve as alterações.

### Passo 3: Instalar as Bibliotecas

As bibliotecas são ferramentas adicionais que o código usa para mexer no Excel.

1.  Abra a pasta onde você salvou os arquivos.
2.  Abra o terminal
3.  No terminal, digite o seguinte comando e aperte Enter:

    pip install -r requirements.txt

4.  Aguarde o fim da instalação. Se aparecerem mensagens de sucesso, está pronto.

## Como Executar

Sempre certifique-se de que a planilha Excel original está fechada antes de rodar os scripts.

1.  Abra o terminal na pasta do projeto (conforme explicado no Passo 3, item 1 e 2).
2.  Para gerar a tabela de dimensão, digite o comando abaixo e aperte Enter:

    python dimensao.py

3.  Para gerar a tabela de fatos, digite o comando abaixo e aperte Enter:

    python fato.py

## Resultado

Após executar os comandos, um novo arquivo chamado `TRATAMENTO DE DADOS.xlsx` será criado (ou atualizado) na mesma pasta. Ele conterá as abas com os dados processados prontos para uso.

## Resolução de Problemas Comuns

* **Erro "PermissionError"**: Acontece se o arquivo Excel estiver aberto. Feche o Excel e tente novamente.
* **Erro "FileNotFoundError"**: Ocorre se o script não encontrar a planilha original. Verifique se o nome do arquivo Excel está exatamente igual ao que está escrito dentro do código (Passo 2) e se ele está na mesma pasta dos scripts.
* **Comando "python" não reconhecido**: Provavelmente o Python não foi adicionado ao PATH durante a instalação. Reinstale o Python e lembre-se de marcar a caixa "Add Python to PATH".
