"""
Quebra a planilha de vendas em um arquivo por vendedor e manda
um e-mail pelo Outlook pra cada um com o relatório dele.

Uso:
    python main.py              -> só gera os arquivos e abre os e-mails pra conferir
    python main.py --enviar     -> gera os arquivos e já envia
"""
import sys
from pathlib import Path

import pandas as pd

PASTA = Path(__file__).parent
ARQ_VENDAS = PASTA / "dados" / "vendas.xlsx"
ARQ_EMAILS = PASTA / "dados" / "emails.xlsx"
PASTA_SAIDA = PASTA / "relatorios"

COLUNA_VENDEDOR = "Vendedor"


def quebrar_por_vendedor(arq_vendas, pasta_saida):
    vendas = pd.read_excel(arq_vendas)
    pasta_saida.mkdir(exist_ok=True)

    arquivos = {}
    for vendedor, tabela in vendas.groupby(COLUNA_VENDEDOR):
        caminho = pasta_saida / f"{vendedor}.xlsx"
        tabela.to_excel(caminho, index=False)
        arquivos[vendedor] = caminho
        print(f"{vendedor}: {len(tabela)} vendas -> {caminho.name}")

    return vendas, arquivos


def montar_corpo(nome, tabela):
    tabela_html = tabela.to_html(index=False, border=1, justify="left")
    return f"""
    <p>Boa noite, {nome}, segue o relatório com suas vendas</p>
    {tabela_html}
    <p>Qualquer dúvida é só responder este e-mail.</p>
    """


def enviar_emails(vendas, arq_emails, pasta_saida, enviar=False):
    # import aqui dentro porque o win32com só existe no Windows
    import win32com.client as win32

    outlook = win32.Dispatch("outlook.application")
    contatos = pd.read_excel(arq_emails)

    for _, linha in contatos.iterrows():
        email = linha.iloc[0]   # coluna A
        nome = linha.iloc[1]    # coluna B

        anexo = pasta_saida / f"{nome}.xlsx"
        if not anexo.exists():
            print(f"Não achei o arquivo de {nome}, pulando")
            continue

        tabela = vendas[vendas[COLUNA_VENDEDOR] == nome]

        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = f"Relatório de vendas - {nome}"
        mail.HTMLBody = montar_corpo(nome, tabela)
        mail.Attachments.Add(str(anexo.resolve()))

        if enviar:
            mail.Send()
            print(f"Enviado para {nome} ({email})")
        else:
            mail.Display()
            print(f"E-mail de {nome} aberto pra conferir")


if __name__ == "__main__":
    enviar = "--enviar" in sys.argv

    vendas, _ = quebrar_por_vendedor(ARQ_VENDAS, PASTA_SAIDA)

    if sys.platform != "win32":
        print("\nO envio usa o Outlook, então só funciona no Windows. Os arquivos já foram gerados.")
    else:
        enviar_emails(vendas, ARQ_EMAILS, PASTA_SAIDA, enviar)
