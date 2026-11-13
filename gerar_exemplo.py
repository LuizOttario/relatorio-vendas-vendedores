# Gera planilhas de exemplo em dados/ pra quem não tem os arquivos do curso.
# Se você tem os arquivos originais, é só colocar eles na pasta dados/ com os mesmos nomes.
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

pasta = Path(__file__).parent / "dados"
pasta.mkdir(exist_ok=True)

vendedores = {
    "Ana": "ana@exemplo.com",
    "Bruno": "bruno@exemplo.com",
    "Carla": "carla@exemplo.com",
    "Diego": "diego@exemplo.com",
}
produtos = {"Notebook": 3500, "Monitor": 900, "Teclado": 150, "Mouse": 80, "Headset": 250}

random.seed(1)
linhas = []
inicio = date(2024, 1, 1)
for _ in range(60):
    produto = random.choice(list(produtos))
    qtd = random.randint(1, 5)
    linhas.append({
        "Data": inicio + timedelta(days=random.randint(0, 89)),
        "Vendedor": random.choice(list(vendedores)),
        "Produto": produto,
        "Quantidade": qtd,
        "Valor": produtos[produto] * qtd,
    })

vendas = pd.DataFrame(linhas).sort_values("Data")
vendas.to_excel(pasta / "vendas.xlsx", index=False)

emails = pd.DataFrame({"E-mail": list(vendedores.values()), "Vendedor": list(vendedores)})
emails.to_excel(pasta / "emails.xlsx", index=False)

print("Planilhas de exemplo criadas em", pasta)
