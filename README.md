# Relatório de vendas por vendedor

Projeto da seção 13 do curso de Python + RPA.

O script lê uma planilha com todas as vendas, separa em um arquivo Excel por vendedor e depois monta um e-mail no Outlook pra cada um, com a tabela das vendas no corpo e o arquivo dele em anexo.

## Como funciona

1. Lê `dados/vendas.xlsx` e agrupa pela coluna `Vendedor`. Cada grupo vira um arquivo em `relatorios/` com o nome do vendedor.
2. Lê `dados/emails.xlsx`. A coluna A é o e-mail e a coluna B é o nome do vendedor.
3. Pra cada linha, procura o arquivo com o nome da coluna B, cria o e-mail ("Boa noite, fulano, segue o relatório com suas vendas"), coloca a tabela no corpo e anexa o arquivo.

Por padrão os e-mails só são abertos na tela pra conferir. Pra enviar direto use `--enviar`.

## Rodando

```bash
pip install -r requirements.txt
python gerar_exemplo.py   # só se não tiver as planilhas do curso
python main.py
python main.py --enviar
```

A parte do e-mail usa o `pywin32` pra controlar o Outlook, então precisa de Windows com o Outlook instalado e configurado. Em outro sistema o script só gera os arquivos.

Se a coluna do vendedor na sua planilha tiver outro nome, é só mudar `COLUNA_VENDEDOR` no começo do `main.py`.
