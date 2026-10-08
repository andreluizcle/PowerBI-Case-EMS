"""Gera a página 2 ("Impacto") e o que ela precisa no modelo, SEM tocar na página 1.

Modelo
  - tables/Cenarios.tmdl e tables/Descontos.tmdl: tabelas novas, embutidas (#table),
    desconectadas (sem relacionamentos).
  - Vendas.tmdl: as medidas novas são INSERIDAS no fim do bloco de medidas; nada do
    que já existe é alterado (só linhas novas). Se já estiverem lá, não insere de novo.
Relatório
  - pages/impacto/: page.json (fundo, filtro de página Região = Nordeste) + visuais.
  - pages.json: acrescenta "impacto" no fim da ordem das páginas.

Uso:  python tools/gerar_impacto.py
"""
import hashlib
import json
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gerar_relatorio import (  # noqa: E402  (mesmo estilo e cores da página 1)
    AZUL_ESCURO, CORES_CANAL, FUNDO_PAGINA, LARANJA, SCHEMA, SCHEMA_PAGINA, conteiner, cor, cores_por_valor, gravar,
    literal, ordenar, rotulos, sem_eixo_de_valores, texto)

RAIZ = Path(__file__).resolve().parents[1]
TABELAS = RAIZ / "Oportunidades.SemanticModel" / "definition" / "tables"
PAGINAS = RAIZ / "Oportunidades.Report" / "definition" / "pages"
PASTA_PAGINA = PAGINAS / "impacto"
PASTA_VISUAIS = PASTA_PAGINA / "visuals"
T1, T2, T4, M = "\t", "\t\t", "\t\t\t\t", "    "

# ------------------------------------------------------------------ modelo: tabelas novas
CENARIOS = [  # Cenário, Ordem, Penetração dos genéricos, Participação da EMS entre os genéricos
    ("Conservador", 1, 0.40, 0.23),
    ("Base", 2, 0.50, 0.23),
    ("Com foco no Nordeste", 3, 0.50, 0.30),
]
DESCONTOS = [(0.35, "35% (teto)"), (0.40, "40%"), (0.45, "45%")]  # Desconto, Rótulo


def coluna_tmdl(nome, tipo, formato=None, ordenar_por=None, expressao=None, descricao=None) -> str:
    """Bloco TMDL de uma coluna. Com 'expressao' vira coluna calculada (DAX)."""
    nome_tmdl = f"'{nome}'" if re.search(r"[\s.=:']", nome) else nome
    linhas = [f"{T1}/// {descricao}"] if descricao else []
    linhas.append(f"{T1}column {nome_tmdl}" + (f" = {expressao}" if expressao else ""))
    linhas.append(f"{T2}dataType: {tipo}")
    if formato:
        linhas.append(f"{T2}formatString: {formato}")
    linhas.append(f"{T2}summarizeBy: none")
    if not expressao:
        linhas.append(f"{T2}sourceColumn: {nome}")
    if ordenar_por:
        linhas.append(f"{T2}sortByColumn: {ordenar_por}")
    linhas += ["", f"{T2}annotation SummarizationSetBy = User", ""]
    return "\n".join(linhas)


def particao_m(tabela: str, tipo_tabela: str, linhas_m: list[str]) -> str:
    dados = ",\n".join(f"{T4}{M}{M}{M}{l}" for l in linhas_m)
    return "\n".join([
        f"{T1}partition {tabela} = m",
        f"{T2}mode: import",
        f"{T2}source =",
        f"{T4}let",
        f"{T4}{M}Fonte = #table(",
        f"{T4}{M}{M}{tipo_tabela},",
        f"{T4}{M}{M}{{",
        dados,
        f"{T4}{M}{M}}}",
        f"{T4}{M})",
        f"{T4}in",
        f"{T4}{M}Fonte",
        "",
    ])


def tmdl_cenarios() -> str:
    return "\n".join([
        "/// Cenários de share (hipóteses do case). Tabela desconectada: não filtra Vendas.",
        "table Cenarios",
        "",
        coluna_tmdl("Cenário", "string", ordenar_por="Ordem"),
        coluna_tmdl("Ordem", "int64", formato="0"),
        coluna_tmdl("Penetração dos genéricos", "double", formato="0.0%",
                    descricao="Quanto os genéricos levam do volume da molécula (hipótese)."),
        coluna_tmdl("Participação da EMS entre os genéricos", "double", formato="0.0%",
                    descricao="Participação da EMS entre os genéricos (23% hoje, segundo a imprensa; 30% é hipótese)."),
        coluna_tmdl("Share total", "double", formato="0.0%",
                    expressao="Cenarios[Penetração dos genéricos] * Cenarios[Participação da EMS entre os genéricos]",
                    descricao="Share da EMS no volume total = penetração × participação."),
        particao_m("Cenarios", 'type table [#"Cenário" = text, Ordem = Int64.Type, '
                   '#"Penetração dos genéricos" = number, #"Participação da EMS entre os genéricos" = number]',
                   [f'{{"{c}", {o}, {p:.2f}, {s:.2f}}}' for c, o, p, s in CENARIOS]),
    ])


def tmdl_descontos() -> str:
    return "\n".join([
        "/// Níveis de desconto sobre o preço atual (35% é o mínimo exigido para genérico). Tabela desconectada.",
        "table Descontos",
        "",
        coluna_tmdl("Desconto", "double", formato="0%"),
        coluna_tmdl("Rótulo", "string", ordenar_por="Desconto"),
        particao_m("Descontos", 'type table [Desconto = number, #"Rótulo" = text]',
                   [f'{{{d:.2f}, "{r}"}}' for d, r in DESCONTOS]),
    ])


# ------------------------------------------------------------------ modelo: medidas novas (tabela Vendas)
# (nome, expressão DAX exata da especificação, formato, descrição)
MEDIDAS = [
    ("Fator Preço Genérico", "0.65", "0%",
     "Premissa: preço do genérico = 65% do preço atual (teto Anvisa/CMED: ao menos 35% abaixo do referência)."),
    ("Margem Genérico", "0.30", "0%", "Premissa: margem bruta de 30% no preço-teto (dado do case)."),
    ("Mercado atual (R$ mi)", "[Receita (R$ mi)]", "#,0.0", None),
    ("Mercado a preço de genérico (R$ mi)", "[Receita (R$ mi)] * [Fator Preço Genérico]", "#,0.0", None),
    ("Share Base", 'CALCULATE(MAX(Cenarios[Share total]), Cenarios[Cenário] = "Base")', "0.0%", None),
    ("Receita EMS (R$ mi)", "[Mercado a preço de genérico (R$ mi)] * [Share Base]", "#,0.0", None),
    ("Lucro bruto EMS (R$ mi)", "[Receita EMS (R$ mi)] * [Margem Genérico]", "#,0.0", None),
    ("Share do cenário", "SELECTEDVALUE(Cenarios[Share total])", "0.0%", None),
    ("Unidades EMS no cenário (mi)", "[Volume (mi un.)] * [Share do cenário]", "#,0.0", None),
    ("Receita EMS no cenário (R$ mi)", "[Mercado a preço de genérico (R$ mi)] * [Share do cenário]", "#,0.0", None),
    ("Lucro bruto no cenário (R$ mi)", "[Receita EMS no cenário (R$ mi)] * [Margem Genérico]", "#,0.0", None),
    ("Custo unitário (R$)", "[Preço Médio (R$)] * [Fator Preço Genérico] * (1 - [Margem Genérico])", "#,0.00", None),
    ("Preço do genérico no desconto (R$)", "[Preço Médio (R$)] * (1 - SELECTEDVALUE(Descontos[Desconto]))", "#,0.00", None),
    ("Lucro por unidade no desconto (R$)", "[Preço do genérico no desconto (R$)] - [Custo unitário (R$)]", "#,0.00", None),
    # Cenário conservador (cartões 3 e 4 e barras por canal)
    ("Share Conservador", 'CALCULATE(MAX(Cenarios[Share total]), Cenarios[Cenário] = "Conservador")', "0.0%", None),
    ("Receita EMS conservador (R$ mi)", "[Mercado a preço de genérico (R$ mi)] * [Share Conservador]", "#,0.0", None),
    ("Lucro bruto EMS conservador (R$ mi)", "[Receita EMS conservador (R$ mi)] * [Margem Genérico]", "#,0.0", None),
]


def inserir_medidas() -> str:
    """Acrescenta ao fim do bloco de medidas de Vendas.tmdl as medidas da lista que ainda não existem.
    Nada do que já está no arquivo é alterado (só linhas novas)."""
    arquivo = TABELAS / "Vendas.tmdl"
    with open(arquivo, encoding="utf-8", newline="") as f:
        texto = f.read()
    fim = "\r\n" if "\r\n" in texto else "\n"
    faltam = [m for m in MEDIDAS if f"measure '{m[0]}' =" not in texto]
    if not faltam:
        return "todas as medidas novas já estavam em Vendas.tmdl (nada inserido)"
    linhas = texto.split(fim)
    # ponto de inserção: logo antes da primeira coluna (e da descrição dela, se houver)
    i = next(n for n, l in enumerate(linhas) if l.startswith(f"{T1}column "))
    while linhas[i - 1].startswith(f"{T1}///"):
        i -= 1
    bloco = []
    for nome, expr, fmt, desc in faltam:
        if desc:
            bloco.append(f"{T1}/// {desc}")
        bloco += [f"{T1}measure '{nome}' = {expr}", f"{T2}formatString: {fmt}", ""]
    with open(arquivo, "w", encoding="utf-8", newline="") as f:
        f.write(fim.join(linhas[:i] + bloco + linhas[i:]))
    return f"{len(faltam)} medida(s) inserida(s) em Vendas.tmdl: " + ", ".join(m[0] for m in faltam)


def gravar_tmdl(nome: str, conteudo: str) -> None:
    (TABELAS / nome).write_text(conteudo, encoding="utf-8", newline="\r\n")


# ------------------------------------------------------------------ relatório: ajudantes
def campo(tipo: str, tabela: str, nome: str) -> dict:
    return {tipo: {"Expression": {"SourceRef": {"Entity": tabela}}, "Property": nome}}


def col(tabela: str, nome: str) -> dict:
    return campo("Column", tabela, nome)


def med(nome: str) -> dict:
    return campo("Measure", "Vendas", nome)


def projecao(c) -> dict:
    """Campo do visual. Aceita (campo, rótulo) para trocar só o texto exibido neste visual
    (displayName); o campo do modelo continua o mesmo."""
    c, rotulo = c if isinstance(c, tuple) else (c, None)
    alvo = next(iter(c.values()))
    p = {"field": c, "queryRef": f"{alvo['Expression']['SourceRef']['Entity']}.{alvo['Property']}",
         "nativeQueryRef": alvo["Property"]}
    if rotulo:
        p["displayName"] = rotulo
    return p


def visual(nome, ordem, x, y, largura, altura, tipo, papeis=None, objects=None, ordem_por=None, vco=None) -> dict:
    corpo = {"visualType": tipo}
    if papeis:
        corpo["query"] = {"queryState": {p: {"projections": [projecao(c) for c in cs]} for p, cs in papeis.items()}}
        if ordem_por:
            corpo["query"]["sortDefinition"] = ordem_por
    if objects:
        corpo["objects"] = objects
    if vco:
        corpo["visualContainerObjects"] = vco
    return {
        "$schema": SCHEMA,
        "name": nome,
        "position": {"x": x, "y": y, "z": ordem * 1000, "height": altura, "width": largura, "tabOrder": ordem * 1000},
        "visual": corpo,
    }


def caixa_texto(nome, ordem, x, y, largura, altura, conteudo: str, tamanho: str, negrito: bool = False,
                vco=None) -> dict:
    """Caixa de texto com um parágrafo (lista nativa de parágrafos, como o Desktop grava)."""
    estilo = {"fontWeight": "bold"} if negrito else {}
    estilo["fontSize"] = tamanho
    return visual(nome, ordem, x, y, largura, altura, "textbox", objects={
        "general": [{"properties": {"paragraphs": [{"textRuns": [{"value": conteudo, "textStyle": estilo}]}]}}],
    }, vco=vco)


def filtro_pagina_regiao(regiao: str) -> dict:
    """Filtro de página fixo (travado no modo de leitura): Vendas[Região] = regiao."""
    def c(src):
        return {"Column": {"Expression": {"SourceRef": src}, "Property": "Região"}}

    return {
        "name": "Filter" + hashlib.md5(f"pagina-impacto-regiao-{regiao}".encode()).hexdigest()[:24],
        "field": c({"Entity": "Vendas"}),
        "type": "Categorical",
        "filter": {
            "Version": 2,
            "From": [{"Name": "v", "Entity": "Vendas", "Type": 0}],
            "Where": [{"Condition": {"In": {
                "Expressions": [c({"Source": "v"})],
                "Values": [[{"Literal": {"Value": f"'{regiao}'"}}]],
            }}}],
        },
        "howCreated": "User",
        "isLockedInViewMode": True,
    }


# ------------------------------------------------------------------ relatório: página e visuais
PAGINA = {
    "$schema": SCHEMA_PAGINA,
    "name": "impacto",
    "displayName": "Impacto",
    "displayOption": "FitToPage",
    "height": 720,
    "width": 1280,
    "filterConfig": {"filters": [filtro_pagina_regiao("Nordeste")]},
    "objects": {"background": [{"properties": {"color": cor(FUNDO_PAGINA), "transparency": literal("0D")}}]},
}

CENARIO, CANAL, ROTULO = col("Cenarios", "Cenário"), col("Vendas", "Canal"), col("Descontos", "Rótulo")
# Visuais retirados da página a pedido do usuário (a pasta é apagada). O gráfico de descontos saiu, mas a
# tabela Descontos e as medidas de desconto continuam no modelo, sem uso (é reversível).
REMOVIDOS = ["10_premissas", "09_colunas_desconto"]

# Tabela de cenários: (campo, nome de exibição no cabeçalho, largura fixa em px). Soma <= 1208 px
# (1240 menos as margens de 16 + 16); larguras folgadas para cabeçalhos e "Com foco no Nordeste" em 1 linha.
COLUNAS_TABELA = [
    (CENARIO, None, 190),
    (col("Cenarios", "Penetração dos genéricos"), "Genéricos levam (% do volume)", 200),
    (col("Cenarios", "Participação da EMS entre os genéricos"), "EMS entre os genéricos", 170),
    (med("Share do cenário"), "Share da EMS no total", 160),
    (med("Unidades EMS no cenário (mi)"), "Unidades (mi)", 140),
    (med("Receita EMS no cenário (R$ mi)"), "Receita (R$ mi)", 150),
    (med("Lucro bruto no cenário (R$ mi)"), "Lucro bruto (R$ mi)", 160),
]


def larguras_fixas(colunas) -> dict:
    """Largura fixa por coluna (columnWidth + seletor metadata = queryRef da coluna). Só vale com
    a largura automática desligada (autoSizeColumnWidth false + columnAdjustment 'fixedWidth')."""
    return {
        "columnHeaders": [{"properties": {
            "autoSizeColumnWidth": literal("false"), "columnAdjustment": literal("'fixedWidth'"),
            "wordWrap": literal("true"), "fontSize": literal("9D")}}],
        "columnWidth": [{"properties": {"value": literal(f"{largura}D")},
                         "selector": {"metadata": projecao(c)["queryRef"]}} for c, _, largura in colunas],
    }

def cor_unica(hexa: str) -> list[dict]:
    """Gráfico de uma série só: uma cor para todas as colunas/barras (dataPoint.defaultColor)."""
    return [{"properties": {"defaultColor": cor(hexa)}}]


CINZA_LEGENDA = "#616161"  # cinza secundário do tema base


def legenda(texto_legenda: str, tamanho: int) -> list[dict]:
    """Subtítulo do visual (Formatar visual > Geral > Título > Subtítulo): cinza, fonte menor que o título."""
    return [{"properties": {"show": literal("true"), "text": texto(texto_legenda), "fontSize": literal(f"{tamanho}D"),
                            "fontColor": cor(CINZA_LEGENDA), "titleWrap": literal("true")}}]


def conteiner_com_legenda(titulo: str, texto_legenda: str, tamanho: int = 10) -> dict:
    """Mesmo contêiner da página 1 (caixa branca, borda, título 12 negrito) + legenda cinza (10 por padrão),
    com espaço menor entre a legenda e o conteúdo para não espremer o visual."""
    vco = conteiner(titulo)
    vco["subTitle"] = legenda(texto_legenda, tamanho)
    vco["spacing"] = [{"properties": {"customizeSpacing": literal("true"), "spaceBelowSubTitle": literal("6D"),
                                      "spaceBelowTitleArea": literal("6D")}}]
    return vco


URL_FONTE = "https://exame.com/insight/sem-remedios-grupo-ems-espera-aval-do-cade-em-setembro-para-aquisicao-da-medley/p"


def botao_link(nome, ordem, x, y, largura, altura, rotulo: str, url: str) -> dict:
    """Botão com ação 'URL da Web' (Formatar botão > Ação > Tipo: URL da Web). Tabelas não aceitam link, botões sim.
    Segue o contrato do modelo oficial: retângulo arredondado, sem caixa/borda do contêiner, ícone à esquerda e
    desligado. No Desktop, em edição, o link abre com Ctrl+clique; no modo de leitura e no Power BI online, com clique."""
    return visual(nome, ordem, x, y, largura, altura, "actionButton", objects={
        "text": [{"properties": {"show": literal("true")}},
                 {"properties": {"text": texto(rotulo), "fontSize": literal("11D"), "bold": literal("true"),
                                 "fontColor": cor("#FFFFFF"), "horizontalAlignment": literal("'center'")},
                  "selector": {"id": "default"}}],
        "icon": [{"properties": {"show": literal("false")}},
                 {"properties": {"placement": literal("'left'")}, "selector": {"id": "default"}}],
        "fill": [{"properties": {"show": literal("true")}},
                 {"properties": {"fillColor": cor(AZUL_ESCURO)}, "selector": {"id": "default"}}],
        "outline": [{"properties": {"show": literal("false")}}],
        "shape": [{"properties": {"tileShape": literal("'rectangleRounded'"), "roundEdge": literal("8L")}},
                  {"properties": {"tileShape": literal("'rectangleRounded'")}, "selector": {"id": "default"}}],
    }, vco={
        "background": [{"properties": {"show": literal("false")}}],
        "border": [{"properties": {"show": literal("false")}}],
        "visualLink": [{"properties": {"show": literal("true"), "type": literal("'WebUrl'"), "webUrl": texto(url)}}],
    })


def compacto(vco: dict) -> dict:
    """Margens de cima/baixo 8 px e espaço mínimo entre título, legenda e conteúdo (laterais continuam 16 px).
    Usado na tabela de cenários: com a legenda em 2 linhas, 200 px de altura ficavam apertados (barra de rolagem)."""
    vco["padding"] = [{"properties": {"top": literal("8D"), "bottom": literal("8D"), "left": literal("16D"),
                                      "right": literal("16D")}}]
    vco["spacing"] = [{"properties": {"customizeSpacing": literal("true"), "spaceBelowTitle": literal("2D"),
                                      "spaceBelowSubTitle": literal("2D"), "spaceBelowTitleArea": literal("2D")}}]
    return vco


def cartao(nome, ordem, x, medida, titulo, texto_legenda) -> dict:
    """Cartão com título + legenda do contêiner (a legenda fica logo abaixo do título) e o valor.
    Para caber em 298x108: título 10 negrito, legenda 9, margens e espaçamentos menores. O rótulo e a
    borda internos do cartão ficam desligados; a caixa branca arredondada vem do contêiner."""
    return visual(nome, ordem, x, 76, 298, 108, "cardVisual", {"Data": [med(medida)]},
                  objects={
                      "label": [{"properties": {"show": literal("false")}, "selector": {"id": "default"}}],
                      "outline": [{"properties": {"show": literal("false")}, "selector": {"id": "default"}}],
                      "padding": [{"properties": {"paddingIndividual": literal("false"), "paddingUniform": literal("0D")},
                                   "selector": {"id": "default"}}],
                  },
                  vco={
                      "title": [{"properties": {"show": literal("true"), "text": texto(titulo), "fontSize": literal("10D"),
                                                "bold": literal("true"), "titleWrap": literal("false")}}],
                      "subTitle": legenda(texto_legenda, 9),
                      "spacing": [{"properties": {"customizeSpacing": literal("true"), "spaceBelowTitle": literal("2D"),
                                                  "spaceBelowSubTitle": literal("2D"), "spaceBelowTitleArea": literal("4D")}}],
                      "background": [{"properties": {"show": literal("true"), "color": cor("#FFFFFF"),
                                                     "transparency": literal("0D")}}],
                      "border": [{"properties": {"show": literal("true"), "color": cor("#E6E6E6"), "radius": literal("8D"),
                                                 "width": literal("1D")}}],
                      "padding": [{"properties": {"top": literal("10D"), "bottom": literal("8D"), "left": literal("12D"),
                                                  "right": literal("12D")}}],
                      "visualHeader": [{"properties": {"show": literal("true")}}],
                  })


VISUAIS = [
    # 1. Título da página
    #    (largura 1000: abre espaço à direita para o botão do link; o texto ocupa ~600 px)
    caixa_texto("01_titulo", 1, 20, 14, 1000, 52, "Impacto esperado: Nordeste, todos os canais", "22pt", negrito=True),
    # 2-5. Cartões em cadeia: mercado atual -> a preço de genérico -> receita EMS -> lucro bruto EMS
    cartao("02_card_mercado_atual", 2, 20, "Mercado atual (R$ mi)", "Mercado atual do Nordeste (R$ mi)",
           "Soma dos 3 canais (dados do Case)"),
    cartao("03_card_mercado_generico", 3, 334, "Mercado a preço de genérico (R$ mi)",
           "Mercado a preço de genérico (R$ mi)", "Limite Anvisa (65%)"),
    cartao("04_card_receita_ems", 4, 648, "Receita EMS conservador (R$ mi)",
           "Receita da EMS, cenário conservador (R$ mi)", "9,2% de share"),
    cartao("05_card_lucro_ems", 5, 962, "Lucro bruto EMS conservador (R$ mi)",
           "Lucro bruto da EMS, cenário conservador (R$ mi)", "30% de margem (Case)"),
    # 6. Tabela dos 3 cenários (sem linha de total: nela as medidas de cenário ficariam vazias).
    #    Nomes de exibição só no cabeçalho; cabeçalhos com quebra de linha e larguras fixas (cabe em 760 px).
    visual("06_tabela_cenarios", 6, 20, 196, 1240, 200, "tableEx",
           {"Values": [(c, rotulo) if rotulo else c for c, rotulo, _ in COLUNAS_TABELA]},
           objects={
               "total": [{"properties": {"totals": literal("false")}}],
               **larguras_fixas(COLUNAS_TABELA),
               "values": [{"properties": {"fontSize": literal("9D")}}],
           },
           ordem_por=ordenar(CENARIO, "Ascending"),  # Cenário é ordenado pela coluna Ordem
           # Legenda = fonte dos números da tabela (matéria verificada) + hipóteses; fonte 9 para caber em 2 linhas.
           # O link da matéria NÃO vai no visual: tabela não tem ação de link documentada (fica no relatório final).
           vco=compacto(conteiner_com_legenda(
               "Três cenários de share, cada um com a sua lógica",
               "Fonte: Exame INSIGHT, 12/08/2026, “Sem remédios: Grupo EMS espera aval do Cade em setembro para "
               "aquisição da Medley” (genéricos = mais de 40% das unidades comercializadas; EMS = cerca de 23% do "
               "mercado de genéricos, sem a Medley). Hipóteses: 50% de volume de genéricos (cenários base e com foco) "
               "e 30% de participação da EMS no cenário com foco.", tamanho=9))),
    # 7. Colunas: lucro bruto por cenário (Conservador -> Base -> Com foco)
    visual("07_colunas_lucro_cenario", 7, 20, 408, 610, 296, "clusteredColumnChart",
           {"Category": [CENARIO], "Y": [med("Lucro bruto no cenário (R$ mi)")]},
           objects={"dataPoint": cor_unica(AZUL_ESCURO), "labels": rotulos()},
           ordem_por=ordenar(CENARIO, "Ascending"),
           vco=conteiner_com_legenda("Do conservador ao com foco, o lucro bruto vai de R$ 4,1 mi a R$ 6,7 mi",
                                     "Lucro bruto = receita × 30% de margem (Case)")),
    # 8. Barras: lucro bruto do cenário-base por canal
    visual("08_barras_lucro_canal", 8, 646, 408, 614, 296, "clusteredBarChart",
           {"Category": [CANAL], "Y": [med("Lucro bruto EMS conservador (R$ mi)")]},
           objects={"dataPoint": cores_por_valor("Canal", CORES_CANAL), "labels": rotulos(),
                    "valueAxis": sem_eixo_de_valores()},  # como as barras da página 1
           ordem_por=ordenar(med("Lucro bruto EMS conservador (R$ mi)"), "Descending"),
           vco=conteiner_com_legenda("O Hospitalar gera a maior fatia do lucro bruto no Nordeste (cenário conservador)",
                                     "Cenário conservador: 9,2% de share e 30% de margem (Case)")),
    # 11. Link clicável para a matéria da Exame citada na legenda da tabela (canto direito da faixa do título)
    botao_link("11_link_fonte", 11, 1040, 20, 220, 40, "Abrir fonte: Exame INSIGHT", URL_FONTE),
]


def eh_caixa_fontes(arquivo: Path) -> bool:
    """Caixa de texto cujo título ou texto começa com 'Fontes' (criada por uma versão anterior do pedido)."""
    if not arquivo.exists():
        return False
    v = json.loads(arquivo.read_text(encoding="utf-8")).get("visual", {})
    if v.get("visualType") != "textbox":
        return False
    titulo = json.dumps(v.get("visualContainerObjects", {}).get("title", []), ensure_ascii=False)
    paragrafos = v.get("objects", {}).get("general", [{}])[0].get("properties", {}).get("paragraphs", [])
    texto_caixa = "".join(r.get("value", "") for p in paragrafos for r in p.get("textRuns", []) if isinstance(r.get("value"), str))
    return "'Fontes" in titulo or texto_caixa.strip().startswith("Fontes")


def atualizar_ordem_paginas() -> str:
    arquivo = PAGINAS / "pages.json"
    meta = json.loads(arquivo.read_text(encoding="utf-8"))
    if PAGINA["name"] in meta["pageOrder"]:
        return "pages.json já tinha a página 'impacto'"
    meta["pageOrder"].append(PAGINA["name"])  # página inicial continua a mesma
    gravar(arquivo, meta)
    return "pages.json: 'impacto' acrescentada no fim da ordem das páginas"


def main() -> None:
    for arquivo, gerar in (("Cenarios.tmdl", tmdl_cenarios), ("Descontos.tmdl", tmdl_descontos)):
        if (TABELAS / arquivo).exists():  # modelo pronto: não regrava (o Desktop pode ter acrescentado dados nele)
            print(f"OK: {arquivo} já existe (não alterado)")
        else:
            gravar_tmdl(arquivo, gerar())
            print(f"OK: {arquivo} criado")
    print("OK: " + inserir_medidas())
    gravar(PASTA_PAGINA / "page.json", PAGINA)
    for nome in REMOVIDOS:
        if (PASTA_VISUAIS / nome).exists():
            shutil.rmtree(PASTA_VISUAIS / nome)
            print(f"OK: visual '{nome}' removido da página")
    nomes = {v["name"] for v in VISUAIS}
    for pasta in (PASTA_VISUAIS.iterdir() if PASTA_VISUAIS.exists() else []):
        if pasta.is_dir() and pasta.name not in nomes and eh_caixa_fontes(pasta / "visual.json"):
            shutil.rmtree(pasta)  # a referência agora fica no subtítulo da tabela: não deve existir caixa "Fontes"
            print(f"OK: caixa de texto 'Fontes' ({pasta.name}) removida")
    if PASTA_VISUAIS.exists():  # não apaga nada: só avisa
        for pasta in PASTA_VISUAIS.iterdir():
            if pasta.is_dir() and pasta.name not in nomes:
                print(f"AVISO: visual '{pasta.name}' não está na lista deste script (mantido como está)")
    for v in VISUAIS:
        gravar(PASTA_VISUAIS / v["name"] / "visual.json", v)
    print(f"OK: página 'impacto' + {len(VISUAIS)} visuais gravados")
    print("OK: " + atualizar_ordem_paginas())


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
