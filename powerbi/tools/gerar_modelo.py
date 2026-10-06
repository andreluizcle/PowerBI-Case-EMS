"""Gera a tabela Vendas do modelo (TMDL) a partir de "Dados - Negócios.xlsx".

Os 75 registros ficam embutidos numa expressão M (#table), então o modelo
não depende de nenhum caminho de arquivo do computador.

Uso:  python tools/gerar_modelo.py
"""
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]  # pasta powerbi/
XLSX = RAIZ.parent / "Dados - Negócios.xlsx"
SAIDA = RAIZ / "Oportunidades.SemanticModel" / "definition" / "tables" / "Vendas.tmdl"

# Nome da coluna no xlsx -> nome da coluna no modelo
COLUNAS = {
    "Região": "Região",
    "Estado": "Estado",
    "Canal": "Canal",
    "Unidades vendidas (mil)": "Unidades (mil)",
    "Preço médio por unidade (R$)": "Preço médio por unidade (R$)",
}


def ler_base() -> pd.DataFrame:
    df = pd.read_excel(XLSX, sheet_name=0)
    df = df.rename(columns=COLUNAS)[list(COLUNAS.values())]
    df["Preço médio por unidade (R$)"] = df["Preço médio por unidade (R$)"].round(2)

    # Conferências básicas antes de gerar qualquer coisa
    assert len(df) == 75, f"esperava 75 linhas, li {len(df)}"
    assert df.notna().all().all(), "há células vazias na base"
    assert df["Estado"].nunique() == 25, "esperava 25 UFs"
    assert sorted(df["Canal"].unique()) == ["Hospitalar", "Institucional", "Varejo"]
    assert not df.duplicated(["Estado", "Canal"]).any(), "há Estado×Canal repetido"
    return df


def texto_m(valor: str) -> str:
    return '"' + str(valor).replace('"', '""') + '"'


def linhas_m(df: pd.DataFrame) -> list[str]:
    linhas = []
    for i, r in enumerate(df.itertuples(index=False)):
        virgula = "," if i < len(df) - 1 else ""
        linhas.append(
            f"{{{texto_m(r[0])}, {texto_m(r[1])}, {texto_m(r[2])}, {int(r[3])}, {r[4]:.2f}}}{virgula}"
        )
    return linhas


def gerar_tmdl(df: pd.DataFrame) -> str:
    t1, t2, t4 = "\t", "\t\t", "\t\t\t\t"
    m = "    "  # indentação interna do código M
    dados = "\n".join(f"{t4}{m}{m}{m}{linha}" for linha in linhas_m(df))

    return f"""table Vendas

{t1}/// Receita total em R$ milhões (a coluna Receita está em R$ mil).
{t1}measure 'Receita (R$ mi)' = SUM(Vendas[Receita (R$ mil)]) / 1000
{t2}formatString: #,0.0

{t1}/// Unidades vendidas em milhões (a coluna Unidades está em mil).
{t1}measure 'Volume (mi un.)' = SUM(Vendas[Unidades (mil)]) / 1000
{t2}formatString: #,0.0

{t1}/// Preço médio ponderado = receita ÷ volume. Nunca a média simples da coluna de preço.
{t1}measure 'Preço Médio (R$)' = DIVIDE([Receita (R$ mi)], [Volume (mi un.)])
{t2}formatString: #,0.00

{t1}column Região
{t2}dataType: string
{t2}summarizeBy: none
{t2}sourceColumn: Região

{t2}annotation SummarizationSetBy = User

{t1}column Estado
{t2}dataType: string
{t2}summarizeBy: none
{t2}sourceColumn: Estado

{t2}annotation SummarizationSetBy = User

{t1}column Canal
{t2}dataType: string
{t2}summarizeBy: none
{t2}sourceColumn: Canal

{t2}annotation SummarizationSetBy = User

{t1}column 'Unidades (mil)'
{t2}dataType: int64
{t2}formatString: #,0
{t2}summarizeBy: sum
{t2}sourceColumn: Unidades (mil)

{t2}annotation SummarizationSetBy = User

{t1}/// Preço médio por unidade de cada linha (UF × canal). Não somar nem tirar média simples: use a medida Preço Médio (R$).
{t1}column 'Preço médio por unidade (R$)'
{t2}dataType: double
{t2}formatString: #,0.00
{t2}summarizeBy: none
{t2}sourceColumn: Preço médio por unidade (R$)

{t2}annotation SummarizationSetBy = User

{t1}/// Receita da linha em R$ mil = Unidades (mil) × Preço médio por unidade (R$).
{t1}column 'Receita (R$ mil)'
{t2}dataType: double
{t2}formatString: #,0.00
{t2}summarizeBy: sum
{t2}sourceColumn: Receita (R$ mil)

{t2}annotation SummarizationSetBy = User

{t1}/// Rótulo da combinação, ex.: "SC - Varejo".
{t1}column 'Estado - Canal'
{t2}dataType: string
{t2}summarizeBy: none
{t2}sourceColumn: Estado - Canal

{t2}annotation SummarizationSetBy = User

{t1}partition Vendas = m
{t2}mode: import
{t2}source =
{t4}let
{t4}{m}// {len(df)} registros copiados de "Dados - Negócios.xlsx" (preço arredondado a 2 casas)
{t4}{m}Fonte = #table(
{t4}{m}{m}type table [#"Região" = text, Estado = text, Canal = text, #"Unidades (mil)" = Int64.Type, #"Preço médio por unidade (R$)" = number],
{t4}{m}{m}{{
{dados}
{t4}{m}{m}}}
{t4}{m}),
{t4}{m}// Receita (R$ mil) = Unidades (mil) × Preço (R$); arredondada a centavos só para tirar ruído de ponto flutuante
{t4}{m}Receita = Table.AddColumn(Fonte, "Receita (R$ mil)", each Number.Round([#"Unidades (mil)"] * [#"Preço médio por unidade (R$)"], 2), type number),
{t4}{m}EstadoCanal = Table.AddColumn(Receita, "Estado - Canal", each [Estado] & " - " & [Canal], type text)
{t4}in
{t4}{m}EstadoCanal

"""


def main() -> None:
    df = ler_base()
    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    # UTF-8 sem BOM e fim de linha CRLF, como o Power BI Desktop grava
    SAIDA.write_text(gerar_tmdl(df), encoding="utf-8", newline="\r\n")
    print(f"OK: {SAIDA.relative_to(RAIZ)} gerado com {len(df)} registros")


if __name__ == "__main__":
    main()
