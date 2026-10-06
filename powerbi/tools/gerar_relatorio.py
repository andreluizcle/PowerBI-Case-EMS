"""Gera a página "Oportunidades" do relatório (page.json + os visual.json do PBIR).

Cada visual é descrito uma vez na lista VISUAIS abaixo (tipo, posição, campos e
formatação). Uso:  python tools/gerar_relatorio.py

Atenção: se você alterar a página no Power BI Desktop e salvar, rodar este
script de novo sobrescreve a página e os visuais com a definição daqui.
"""
import hashlib
import json
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
PASTA_PAGINA = RAIZ / "Oportunidades.Report" / "definition" / "pages" / "oportunidades"
PASTA_VISUAIS = PASTA_PAGINA / "visuals"
SCHEMA_PAGINA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json"
TABELA = "Vendas"

# Cores da especificação
FUNDO_PAGINA = "#E6EDF0"
BRANCO = "#FFFFFF"
BORDA = "#E6E6E6"  # mesma cor de borda do tema base do Power BI
AZUL_ESCURO, CINZA_AZULADO, LARANJA = "#0E4A5C", "#8AA3B2", "#D9701A"
CORES_REGIAO = {"Nordeste": AZUL_ESCURO, "Norte": AZUL_ESCURO,
                "Centro-Oeste": CINZA_AZULADO, "Sudeste": CINZA_AZULADO, "Sul": CINZA_AZULADO}
CORES_CANAL = {"Varejo": LARANJA, "Hospitalar": AZUL_ESCURO, "Institucional": CINZA_AZULADO}


# ------------------------------------------------------------------ ajudantes
def coluna(nome: str) -> dict:
    return {"Column": {"Expression": {"SourceRef": {"Entity": TABELA}}, "Property": nome}}


def medida(nome: str) -> dict:
    return {"Measure": {"Expression": {"SourceRef": {"Entity": TABELA}}, "Property": nome}}


def projecao(campo: dict) -> dict:
    nome = next(iter(campo.values()))["Property"]
    return {"field": campo, "queryRef": f"{TABELA}.{nome}", "nativeQueryRef": nome}


def literal(valor: str) -> dict:
    """Valor de formatação no formato do PBIR: 'texto', 12D (número), 2L (inteiro), true/false."""
    return {"expr": {"Literal": {"Value": valor}}}


def texto(valor: str) -> dict:
    return literal("'" + valor.replace("'", "''") + "'")


def cor(hexa: str) -> dict:
    return {"solid": {"color": texto(hexa)}}


def cores_por_valor(nome_coluna: str, mapa: dict) -> list[dict]:
    """Uma cor fixa para cada valor da coluna (ex.: Varejo = laranja), via seletor de dado."""
    return [{
        "properties": {"fill": cor(hexa)},
        "selector": {"data": [{"scopeId": {"Comparison": {
            "ComparisonKind": 0,
            "Left": coluna(nome_coluna),
            "Right": {"Literal": {"Value": "'" + valor + "'"}},
        }}}]},
    } for valor, hexa in mapa.items()]


def conteiner(titulo: str) -> dict:
    """Caixa branca com borda fina e cantos arredondados (valores do tema base,
    declarados explicitamente) e o título-mensagem do visual. Fundo, borda,
    margens e cabeçalho vão sempre juntos; o subtítulo automático fica desligado."""
    return {
        "title": [{"properties": {
            "show": literal("true"), "text": texto(titulo), "fontSize": literal("12D"), "bold": literal("true"),
        }}],
        "subTitle": [{"properties": {"show": literal("false")}}],
        "background": [{"properties": {"show": literal("true"), "color": cor(BRANCO), "transparency": literal("0D")}}],
        "border": [{"properties": {"show": literal("true"), "color": cor(BORDA), "radius": literal("8D"), "width": literal("1D")}}],
        "padding": [{"properties": {"top": literal("14D"), "bottom": literal("16D"), "left": literal("16D"), "right": literal("16D")}}],
        "visualHeader": [{"properties": {"show": literal("true")}}],
    }


def ordenar(campo: dict, direcao: str) -> dict:
    """Ordenação definida explicitamente (não a padrão do Power BI)."""
    return {"sort": [{"field": campo, "direction": direcao}], "isDefaultSort": False}


def rotulos(casas: int | None = None) -> list[dict]:
    propriedades = {"show": literal("true")}
    if casas is not None:
        propriedades["labelPrecision"] = literal(f"{casas}L")
    return [{"properties": propriedades}]


def sem_eixo_de_valores() -> list[dict]:
    """Esconde o eixo de valores e as linhas de grade: os rótulos já mostram os números."""
    return [{"properties": {"show": literal("false"), "gridlineShow": literal("false")}}]


def gradiente_fundo(nome_medida: str, cor_menor: str, cor_maior: str) -> dict:
    """Formatação condicional da matriz: fundo da célula em gradiente de 2 cores,
    do menor ao maior valor (mínimo e máximo automáticos). Vale só para as células
    (matchingOption 1), não para os totais. As cores de parada NÃO levam "expr"."""
    ref = f"{TABELA}.{nome_medida}"  # mesmo queryRef usado na projeção da matriz
    return {
        "properties": {"backColor": {"solid": {"color": {"expr": {"FillRule": {
            "Input": {"SelectRef": {"ExpressionName": ref}},
            "FillRule": {"linearGradient2": {
                "min": {"color": {"Literal": {"Value": f"'{cor_menor}'"}}},
                "max": {"color": {"Literal": {"Value": f"'{cor_maior}'"}}},
                "nullColoringStrategy": {"strategy": {"Literal": {"Value": "'asZero'"}}},
            }},
        }}}}}},
        "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": ref},
    }


def filtro_top_n(nome_coluna: str, n: int, coluna_ordem: str) -> dict:
    """Filtro de nível de visual "N superior": mantém os n valores de nome_coluna com
    maior soma de coluna_ordem. A subconsulta ordena por agregação (Soma) da coluna,
    não pela medida: a referência oficial indica erro no Desktop com medida ali."""
    def col(prop):
        return {"Column": {"Expression": {"SourceRef": {"Source": "v"}}, "Property": prop}}

    return {
        "name": "Filter" + hashlib.md5(f"top{n}-{nome_coluna}".encode()).hexdigest()[:24],
        "field": coluna(nome_coluna),
        "type": "TopN",
        "filter": {
            "Version": 2,
            "From": [
                {"Name": "subquery", "Type": 2, "Expression": {"Subquery": {"Query": {
                    "Version": 2,
                    "From": [{"Name": "v", "Entity": TABELA, "Type": 0}],
                    "Select": [{**col(nome_coluna), "Name": "field"}],
                    "OrderBy": [{"Direction": 2, "Expression": {"Aggregation": {"Expression": col(coluna_ordem), "Function": 0}}}],
                    "Top": n,
                }}}},
                {"Name": "v", "Entity": TABELA, "Type": 0},
            ],
            "Where": [{"Condition": {"In": {
                "Expressions": [col(nome_coluna)],
                "Table": {"SourceRef": {"Source": "subquery"}},
            }}}],
        },
        "howCreated": "User",
    }


def visual(nome, ordem, x, y, largura, altura, tipo, papeis=None, objects=None, vco=None, ordem_por=None,
           filtros=None) -> dict:
    corpo = {"visualType": tipo}
    if papeis:
        corpo["query"] = {"queryState": {
            papel: {"projections": [projecao(c) for c in campos]} for papel, campos in papeis.items()
        }}
        if ordem_por:
            corpo["query"]["sortDefinition"] = ordem_por
    if objects:
        corpo["objects"] = objects
    if vco:
        corpo["visualContainerObjects"] = vco
    resultado = {
        "$schema": SCHEMA,
        "name": nome,
        "position": {"x": x, "y": y, "z": ordem * 1000, "height": altura, "width": largura, "tabOrder": ordem * 1000},
        "visual": corpo,
    }
    if filtros:
        resultado["filterConfig"] = {"filters": filtros}
    return resultado


RECEITA, VOLUME, PRECO = medida("Receita (R$ mi)"), medida("Volume (mi un.)"), medida("Preço Médio (R$)")
REGIAO, CANAL, ESTADO_CANAL = coluna("Região"), coluna("Canal"), coluna("Estado - Canal")

# ------------------------------------------------------------------ página
PAGINA = {
    "$schema": SCHEMA_PAGINA,
    "name": "oportunidades",
    "displayName": "Oportunidades",
    "displayOption": "FitToPage",
    "height": 720,
    "width": 1280,
    "objects": {
        "background": [{"properties": {"color": cor(FUNDO_PAGINA), "transparency": literal("0D")}}],
    },
}

# ------------------------------------------------------------------ os 10 elementos
VISUAIS = [
    # 1. Título da página (fonte 20: o usuário reduziu de 22 para o texto não cortar)
    visual("01_titulo", 1, 20, 14, 720, 52, "textbox", objects={
        "general": [{"properties": {"paragraphs": [{
            "textRuns": [{
                "value": "Lançamento do genérico: onde está a oportunidade?",
                "textStyle": {"fontWeight": "bold", "fontSize": "20pt"},
            }],
        }]}}],
    }),
    # 2. Segmentação Canal: estilo lado a lado (Tile), horizontal, sem cabeçalho, sem seleção salva
    visual("02_seg_canal", 2, 760, 18, 340, 44, "slicer", {"Values": [CANAL]}, objects={
        "data": [{"properties": {"mode": literal("'HorizontalList'")}}],
        "general": [{"properties": {"orientation": literal("1D")}}],
        "header": [{"properties": {"show": literal("false")}}],
    }),
    # 3. Segmentação Região: suspensa, sem cabeçalho
    visual("03_seg_regiao", 3, 1112, 18, 148, 44, "slicer", {"Values": [REGIAO]}, objects={
        "data": [{"properties": {"mode": literal("'Dropdown'")}}],
        "header": [{"properties": {"show": literal("false")}}],
    }),
    # 4-6. Cartões (fundo branco e cantos arredondados vêm do próprio cartão no tema base)
    visual("04_card_receita", 4, 20, 80, 190, 84, "cardVisual", {"Data": [RECEITA]}),
    visual("05_card_volume", 5, 220, 80, 190, 84, "cardVisual", {"Data": [VOLUME]}),
    visual("06_card_preco", 6, 420, 80, 190, 84, "cardVisual", {"Data": [PRECO]}),
    # 7. Matriz Região × Canal (linhas pela receita total, decrescente; totais ligados = padrão)
    visual("07_matriz", 7, 20, 176, 590, 320, "pivotTable",
           {"Rows": [REGIAO], "Columns": [CANAL], "Values": [RECEITA]},
           objects={"values": [gradiente_fundo("Receita (R$ mi)", "#F2F8F9", "#4DB6AC")]},
           vco=conteiner("5 dos 15 blocos valem metade do mercado, todos no Norte e Nordeste"),
           ordem_por=ordenar(RECEITA, "Descending")),
    # 8. Barras: receita por região
    visual("08_barras_regiao", 8, 626, 80, 634, 190, "clusteredBarChart",
           {"Category": [REGIAO], "Y": [RECEITA]},
           objects={"dataPoint": cores_por_valor("Região", CORES_REGIAO),
                    "labels": rotulos(), "valueAxis": sem_eixo_de_valores()},
           vco=conteiner("Norte e Nordeste concentram 58% da receita"),
           ordem_por=ordenar(RECEITA, "Descending")),
    # 9. Barras: estado × canal (cor pela legenda Canal)
    visual("09_barras_top10", 9, 626, 282, 634, 422, "clusteredBarChart",
           {"Category": [ESTADO_CANAL], "Series": [CANAL], "Y": [RECEITA]},
           objects={"dataPoint": cores_por_valor("Canal", CORES_CANAL),
                    "labels": rotulos(), "valueAxis": sem_eixo_de_valores(),
                    "legend": [{"properties": {"show": literal("true"), "position": literal("'Top'")}}]},
           vco=conteiner("As 5 maiores combinações estado × canal são todas de varejo"),
           ordem_por=ordenar(RECEITA, "Descending"),
           filtros=[filtro_top_n("Estado - Canal", 10, "Receita (R$ mil)")]),
    # 10. Colunas: preço médio por canal (mesmas cores de canal; ordem alfabética, como na imagem)
    visual("10_colunas_preco", 10, 20, 508, 590, 196, "clusteredColumnChart",
           {"Category": [CANAL], "Y": [PRECO]},
           objects={"dataPoint": cores_por_valor("Canal", CORES_CANAL),
                    "labels": rotulos(casas=2),
                    "valueAxis": [{"properties": {"start": literal("0D")}}]},
           vco=conteiner("O varejo tem o maior preço por unidade, mas só 7% acima do institucional"),
           ordem_por=ordenar(CANAL, "Ascending")),
]


def gravar(destino: Path, conteudo: dict) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text(json.dumps(conteudo, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\r\n")


def main() -> None:
    gravar(PASTA_PAGINA / "page.json", PAGINA)
    nomes = {v["name"] for v in VISUAIS}
    if PASTA_VISUAIS.exists():  # não apaga nada: só avisa sobre visuais que não estão na lista
        for pasta in PASTA_VISUAIS.iterdir():
            if pasta.is_dir() and pasta.name not in nomes:
                print(f"AVISO: visual '{pasta.name}' não está na lista deste script (mantido como está)")
    for v in VISUAIS:
        gravar(PASTA_VISUAIS / v["name"] / "visual.json", v)
    print(f"OK: page.json + {len(VISUAIS)} visuais gerados em {PASTA_PAGINA.relative_to(RAIZ)}")


if __name__ == "__main__":
    main()
