"""Valida o projeto PBIP inteiro. Uso:  python tools/validar.py

1. Codificação: todo arquivo de texto do projeto em UTF-8 sem BOM.
2. JSON bem formado e validado contra o JSON Schema oficial da Microsoft
   (cópia local em tools/schemas/, baixada de github.com/microsoft/json-schemas).
3. TMDL: indentação com tabs, colunas/medidas com nomes e tipos exatos,
   nomes sem conflito (o Power BI não diferencia maiúsculas de minúsculas).
4. Dados: a tabela embutida tem 75 linhas, bate linha a linha com o xlsx
   (lido com pandas) e reproduz os números de referência do case.
5. Relatório: página, pastas e nomes coerentes; todo campo usado num visual
   existe no modelo; visuais dentro da página.
6. Acabamento da página 1.
7-9. Página 2 "Impacto": tabelas Cenarios/Descontos e medidas novas; números
   recalculados com pandas (Região = Nordeste) e números escritos nos títulos;
   página, filtro de página e visuais contra a especificação.
10. Página 1 intacta: diff do git contra o branch main.
"""
import json
import re
import subprocess
import sys
import urllib.error
import urllib.request
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pandas as pd
from jsonschema import Draft7Validator
from referencing import Registry, Resource
from referencing.jsonschema import DRAFT7

RAIZ = Path(__file__).resolve().parents[1]
XLSX = RAIZ.parent / "Dados - Negócios.xlsx"
MODELO = RAIZ / "Oportunidades.SemanticModel"
RELATORIO = RAIZ / "Oportunidades.Report"
TABELA = MODELO / "definition" / "tables" / "Vendas.tmdl"
SCHEMAS = RAIZ / "tools" / "schemas"
URL_SCHEMAS = "https://developer.microsoft.com/json-schemas/"
URL_GITHUB = "https://raw.githubusercontent.com/microsoft/json-schemas/main/"

PAGINA = {"name": "oportunidades", "displayName": "Oportunidades", "width": 1280, "height": 720}

COLUNAS_ESPERADAS = {
    "Região": "string",
    "Estado": "string",
    "Canal": "string",
    "Unidades (mil)": "int64",
    "Preço médio por unidade (R$)": "double",
    "Receita (R$ mil)": "double",
    "Estado - Canal": "string",
}
MEDIDAS_ESPERADAS = {
    "Receita (R$ mi)": ("SUM(Vendas[Receita (R$ mil)]) / 1000", "#,0.0"),
    "Volume (mi un.)": ("SUM(Vendas[Unidades (mil)]) / 1000", "#,0.0"),
    "Preço Médio (R$)": ("DIVIDE([Receita (R$ mi)], [Volume (mi un.)])", "#,0.00"),
}

# Os 10 elementos da especificação (prompt.txt): tipo, x, y, largura, altura, campos por papel
VISUAIS_ESPERADOS = {
    "01_titulo": ("textbox", (20, 14, 720, 52), {}),
    "02_seg_canal": ("slicer", (760, 18, 340, 44), {"Values": ["Canal"]}),
    "03_seg_regiao": ("slicer", (1112, 18, 148, 44), {"Values": ["Região"]}),
    "04_card_receita": ("cardVisual", (20, 80, 190, 84), {"Data": ["Receita (R$ mi)"]}),
    "05_card_volume": ("cardVisual", (220, 80, 190, 84), {"Data": ["Volume (mi un.)"]}),
    "06_card_preco": ("cardVisual", (420, 80, 190, 84), {"Data": ["Preço Médio (R$)"]}),
    "07_matriz": ("pivotTable", (20, 176, 590, 320),
                  {"Rows": ["Região"], "Columns": ["Canal"], "Values": ["Receita (R$ mi)"]}),
    "08_barras_regiao": ("clusteredBarChart", (626, 80, 634, 190), {"Category": ["Região"], "Y": ["Receita (R$ mi)"]}),
    "09_barras_top10": ("clusteredBarChart", (626, 282, 634, 422),
                        {"Category": ["Estado - Canal"], "Series": ["Canal"], "Y": ["Receita (R$ mi)"]}),
    "10_colunas_preco": ("clusteredColumnChart", (20, 508, 590, 196), {"Category": ["Canal"], "Y": ["Preço Médio (R$)"]}),
}
TITULO = "Lançamento do genérico: onde está a oportunidade?"
FONTE_TITULO = "20pt"  # especificação pede 22; o usuário reduziu para 20 para o texto não cortar (decisão registrada)
FUNDO_PAGINA = "#E6EDF0"
COM_CAIXA_BRANCA = ["07_matriz", "08_barras_regiao", "09_barras_top10", "10_colunas_preco"]
CORES_CANAL = {"Varejo": "#D9701A", "Hospitalar": "#0E4A5C", "Institucional": "#8AA3B2"}
CORES = {  # cor esperada de cada valor, por visual (especificação)
    "08_barras_regiao": ("Região", {"Nordeste": "#0E4A5C", "Norte": "#0E4A5C",
                                    "Centro-Oeste": "#8AA3B2", "Sudeste": "#8AA3B2", "Sul": "#8AA3B2"}),
    "09_barras_top10": ("Canal", CORES_CANAL),
    "10_colunas_preco": ("Canal", CORES_CANAL),
}
ORDENACAO = {  # campo e direção esperados (especificação; 10 segue a imagem)
    "07_matriz": ("Receita (R$ mi)", "Descending"),
    "08_barras_regiao": ("Receita (R$ mi)", "Descending"),
    "09_barras_top10": ("Receita (R$ mi)", "Descending"),
    "10_colunas_preco": ("Canal", "Ascending"),
}
ROTULOS = {"08_barras_regiao": None, "09_barras_top10": None, "10_colunas_preco": "2L"}  # casas decimais
TITULOS = {  # textos exatos da especificação
    "07_matriz": "5 dos 15 blocos valem metade do mercado, todos no Norte e Nordeste",
    "08_barras_regiao": "Norte e Nordeste concentram 58% da receita",
    "09_barras_top10": "As 5 maiores combinações estado × canal são todas de varejo",
    "10_colunas_preco": "O varejo tem o maior preço por unidade, mas só 7% acima do institucional",
}
ESTILO_SEGMENTACAO = {"02_seg_canal": "'HorizontalList'", "03_seg_regiao": "'Dropdown'"}

# ================================================================ Página 2 "Impacto" (prompt da página 2)
FASE_PAGINA2 = 3  # 1: só os 4 cartões; 2 e 3: os 10 elementos
ORDEM_PAGINAS = ["oportunidades", "impacto"]
PAGINA2 = {"name": "impacto", "displayName": "Impacto", "width": 1280, "height": 720}
MEDIDAS_IMPACTO = {  # nome: (DAX exato da especificação, formato)
    "Fator Preço Genérico": ("0.65", "0%"),
    "Margem Genérico": ("0.30", "0%"),
    "Mercado atual (R$ mi)": ("[Receita (R$ mi)]", "#,0.0"),
    "Mercado a preço de genérico (R$ mi)": ("[Receita (R$ mi)] * [Fator Preço Genérico]", "#,0.0"),
    "Share Base": ('CALCULATE(MAX(Cenarios[Share total]), Cenarios[Cenário] = "Base")', "0.0%"),
    "Receita EMS (R$ mi)": ("[Mercado a preço de genérico (R$ mi)] * [Share Base]", "#,0.0"),
    "Lucro bruto EMS (R$ mi)": ("[Receita EMS (R$ mi)] * [Margem Genérico]", "#,0.0"),
    "Share do cenário": ("SELECTEDVALUE(Cenarios[Share total])", "0.0%"),
    "Unidades EMS no cenário (mi)": ("[Volume (mi un.)] * [Share do cenário]", "#,0.0"),
    "Receita EMS no cenário (R$ mi)": ("[Mercado a preço de genérico (R$ mi)] * [Share do cenário]", "#,0.0"),
    "Lucro bruto no cenário (R$ mi)": ("[Receita EMS no cenário (R$ mi)] * [Margem Genérico]", "#,0.0"),
    "Custo unitário (R$)": ("[Preço Médio (R$)] * [Fator Preço Genérico] * (1 - [Margem Genérico])", "#,0.00"),
    "Preço do genérico no desconto (R$)": ("[Preço Médio (R$)] * (1 - SELECTEDVALUE(Descontos[Desconto]))", "#,0.00"),
    "Lucro por unidade no desconto (R$)": ("[Preço do genérico no desconto (R$)] - [Custo unitário (R$)]", "#,0.00"),
    "Share Conservador": ('CALCULATE(MAX(Cenarios[Share total]), Cenarios[Cenário] = "Conservador")', "0.0%"),
    "Receita EMS conservador (R$ mi)": ("[Mercado a preço de genérico (R$ mi)] * [Share Conservador]", "#,0.0"),
    "Lucro bruto EMS conservador (R$ mi)": ("[Receita EMS conservador (R$ mi)] * [Margem Genérico]", "#,0.0"),
}
CENARIO_CARTOES_P2 = "Conservador"  # cenário dos cartões 3-4 e das barras por canal
COLUNAS_P2 = {  # tabela: {coluna: (tipo, formato, ordenar por)}
    "Cenarios": {"Cenário": ("string", None, "Ordem"), "Ordem": ("int64", "0", None),
                 "Penetração dos genéricos": ("double", "0.0%", None),
                 "Participação da EMS entre os genéricos": ("double", "0.0%", None),
                 "Share total": ("double", "0.0%", None)},
    "Descontos": {"Desconto": ("double", "0%", None), "Rótulo": ("string", None, "Desconto")},
}
EXPR_SHARE_TOTAL = "Cenarios[Penetração dos genéricos] * Cenarios[Participação da EMS entre os genéricos]"
LINHAS_CENARIOS = [("Conservador", 1, "0.40", "0.23"), ("Base", 2, "0.50", "0.23"), ("Com foco no Nordeste", 3, "0.50", "0.30")]
LINHAS_DESCONTOS = [("0.35", "35% (teto)"), ("0.40", "40%"), ("0.45", "45%")]
ESPERADO_P2 = {  # <verificacao> do prompt, com Região = Nordeste
    "cartoes": ("229,8", "149,4", "13,7", "4,1"),
    "cenarios": {"Conservador": ("9,2%", "4,3", "13,7", "4,1"), "Base": ("11,5%", "5,4", "17,2", "5,2"),
                 "Com foco no Nordeste": ("15,0%", "7,0", "22,4", "6,7")},
    "canais": {"Hospitalar": "1,7", "Institucional": "1,2", "Varejo": "1,2"},  # cenário conservador
    "descontos": {"35% (teto)": "0,96", "40%": "0,71", "45%": "0,47"},
}
VISUAIS_P2 = {  # tipo, (x, y, largura, altura), campos por papel (na ordem) — geometria dos ajustes da página 2
    "01_titulo": ("textbox", (20, 14, 1000, 52), {}),  # estreitada para caber o botão do link
    "11_link_fonte": ("actionButton", (1040, 20, 220, 40), {}),  # botão com o link da matéria da Exame
    "02_card_mercado_atual": ("cardVisual", (20, 76, 298, 108), {"Data": ["Mercado atual (R$ mi)"]}),
    "03_card_mercado_generico": ("cardVisual", (334, 76, 298, 108), {"Data": ["Mercado a preço de genérico (R$ mi)"]}),
    "04_card_receita_ems": ("cardVisual", (648, 76, 298, 108), {"Data": ["Receita EMS conservador (R$ mi)"]}),
    "05_card_lucro_ems": ("cardVisual", (962, 76, 298, 108), {"Data": ["Lucro bruto EMS conservador (R$ mi)"]}),
    "06_tabela_cenarios": ("tableEx", (20, 196, 1240, 200), {"Values": [
        "Cenário", "Penetração dos genéricos", "Participação da EMS entre os genéricos", "Share do cenário",
        "Unidades EMS no cenário (mi)", "Receita EMS no cenário (R$ mi)", "Lucro bruto no cenário (R$ mi)"]}),
    "07_colunas_lucro_cenario": ("clusteredColumnChart", (20, 408, 610, 296),
                                 {"Category": ["Cenário"], "Y": ["Lucro bruto no cenário (R$ mi)"]}),
    "08_barras_lucro_canal": ("clusteredBarChart", (646, 408, 614, 296),
                              {"Category": ["Canal"], "Y": ["Lucro bruto EMS conservador (R$ mi)"]}),
}
REMOVIDOS_P2 = ["10_premissas", "09_colunas_desconto"]  # caixa "Premissas" e gráfico de descontos retirados
BASE_AJUSTES = "pagina-2-ajustes"  # branch de partida desta rodada: tabelas iguais a ele; Vendas só com acréscimos
LARGURA_UTIL_TABELA = 1240 - 16 - 16  # largura da tabela menos as margens do contêiner
CARTOES_P2 = ["02_card_mercado_atual", "03_card_mercado_generico", "04_card_receita_ems", "05_card_lucro_ems"]
ELEMENTOS_POR_FASE = {1: CARTOES_P2, 2: list(VISUAIS_P2), 3: list(VISUAIS_P2)}
TEXTOS_P2 = {  # caixas de texto: (texto exato da especificação, tamanho, negrito)
    "01_titulo": ("Impacto esperado: Nordeste, todos os canais", "22pt", True),
}
ORDENACAO_P2 = {  # cenários na ordem da coluna Ordem, descontos na ordem de Desconto, canais por lucro
    "06_tabela_cenarios": ("Cenário", "Ascending"),
    "07_colunas_lucro_cenario": ("Cenário", "Ascending"),
    "08_barras_lucro_canal": ("Lucro bruto EMS conservador (R$ mi)", "Descending"),
}
CABECALHOS_P2 = [None, "Genéricos levam (% do volume)", "EMS entre os genéricos", "Share da EMS no total",
                 "Unidades (mi)", "Receita (R$ mi)", "Lucro bruto (R$ mi)"]  # nomes de exibição (None = nome do campo)
ROTULOS_P2 = {"07_colunas_lucro_cenario": None, "08_barras_lucro_canal": None}
CARTOES_TITULOS_P2 = {  # título de cada cartão (título do contêiner), texto exato dos ajustes da página 2
    "02_card_mercado_atual": "Mercado atual do Nordeste (R$ mi)",
    "03_card_mercado_generico": "Mercado a preço de genérico (R$ mi)",
    "04_card_receita_ems": "Receita da EMS, cenário conservador (R$ mi)",
    "05_card_lucro_ems": "Lucro bruto da EMS, cenário conservador (R$ mi)",
}
LEGENDAS_P2 = {  # legenda = subtítulo do visual, texto exato dos ajustes da página 2
    "02_card_mercado_atual": "Soma dos 3 canais (dados do Case)",
    "03_card_mercado_generico": "Limite Anvisa (65%)",
    "04_card_receita_ems": "9,2% de share",
    "05_card_lucro_ems": "30% de margem (Case)",
    # Texto da 4a com o trecho das hipóteses corrigido (aprovado: nenhum cenário usa 60%) e, a pedido do usuário,
    # "Hipóteses nossas" -> "Hipóteses" e "penetração" -> "volume de genéricos"
    "06_tabela_cenarios": "Fonte: Exame INSIGHT, 12/08/2026, “Sem remédios: Grupo EMS espera aval do Cade em setembro para "
                          "aquisição da Medley” (genéricos = mais de 40% das unidades comercializadas; EMS = cerca de 23% "
                          "do mercado de genéricos, sem a Medley). Hipóteses: 50% de volume de genéricos (cenários base e "
                          "com foco) e 30% de participação da EMS no cenário com foco.",
    "07_colunas_lucro_cenario": "Lucro bruto = receita × 30% de margem (Case)",
    "08_barras_lucro_canal": "Cenário conservador: 9,2% de share e 30% de margem (Case)",
}
LEGENDAS_APLICADAS = set(LEGENDAS_P2)  # cartões, tabela e gráficos: todas as legendas aplicadas
LEGENDA_ANTIGA_TABELA = "Share = genéricos levam × EMS entre os genéricos"  # não pode existir em lugar nenhum
URL_FONTE = "https://exame.com/insight/sem-remedios-grupo-ems-espera-aval-do-cade-em-setembro-para-aquisicao-da-medley/p"
ROTULO_BOTAO_FONTE = "Abrir fonte: Exame INSIGHT"
CINZA_LEGENDA = "#616161"
TITULOS_P2 = {  # textos da especificação; os números neles são conferidos contra o cálculo
    "06_tabela_cenarios": "Três cenários de share, cada um com a sua lógica",
    "07_colunas_lucro_cenario": "Do conservador ao com foco, o lucro bruto vai de R$ 4,1 mi a R$ 6,7 mi",
    "08_barras_lucro_canal": "O Hospitalar gera a maior fatia do lucro bruto no Nordeste (cenário conservador)",
}
# Arquivos da página 1 (e da casca do projeto) que não podem mudar em relação ao branch main
PROTEGIDOS = ["Oportunidades.Report/definition/pages/oportunidades", "Oportunidades.Report/definition/report.json",
              "Oportunidades.Report/definition/version.json", "Oportunidades.Report/StaticResources",
              "Oportunidades.Report/definition.pbir", "Oportunidades.pbip", "Oportunidades.SemanticModel/definition.pbism",
              "Oportunidades.SemanticModel/definition/database.tmdl", "Oportunidades.SemanticModel/definition/cultures"]
SO_ACRESCIMOS = ["Oportunidades.SemanticModel/definition/tables/Vendas.tmdl",  # medidas novas, nada removido
                 "Oportunidades.SemanticModel/definition/model.tmdl"]  # o Desktop pode acrescentar "ref table"

falhas: list[str] = []


def ok(cond: bool, msg: str) -> bool:
    print(("  OK     " if cond else "  FALHA  ") + msg)
    if not cond:
        falhas.append(msg)
    return cond


def secao(titulo: str) -> None:
    print(f"\n== {titulo} ==")


def r1(x) -> str:  # arredonda como em finanças (meio para cima), vírgula decimal
    return str(Decimal(str(x)).quantize(Decimal("0.1"), ROUND_HALF_UP)).replace(".", ",")


def r2(x) -> str:
    return str(Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP)).replace(".", ",")


def rel(p: Path) -> str:
    return str(p.relative_to(RAIZ))


# ---------------------------------------------------------------- 1. codificação
def arquivos_do_projeto() -> list[Path]:
    exts = {".json", ".pbir", ".pbism", ".pbip", ".tmdl"}
    return sorted(
        p for p in RAIZ.rglob("*")
        if p.is_file() and (p.suffix in exts or p.name == ".platform")
        and "tools" not in p.relative_to(RAIZ).parts and ".pbi" not in p.parts and ".git" not in p.parts
    )


def checar_codificacao(arquivos: list[Path]) -> None:
    secao("1. Codificação (UTF-8 sem BOM)")
    ruins = []
    for p in arquivos:
        b = p.read_bytes()
        try:
            b.decode("utf-8")
        except UnicodeDecodeError:
            ruins.append(f"{rel(p)} não é UTF-8")
        if b.startswith(b"\xef\xbb\xbf"):
            ruins.append(f"{rel(p)} tem BOM")
    ok(not ruins, f"{len(arquivos)} arquivos verificados" + ("" if not ruins else ": " + "; ".join(ruins)))


# ---------------------------------------------------------------- 2. JSON + schema
def carregar_schema(uri: str) -> dict:
    uri = uri.split("#")[0]
    if not uri.startswith(URL_SCHEMAS):
        raise ValueError(f"schema fora do repositório oficial: {uri}")
    caminho = SCHEMAS / uri[len(URL_SCHEMAS):]
    if not caminho.exists():  # baixa uma vez e guarda para validar offline depois
        caminho.parent.mkdir(parents=True, exist_ok=True)
        with urllib.request.urlopen(URL_GITHUB + uri[len(URL_SCHEMAS):], timeout=30) as resp:
            caminho.write_bytes(resp.read())
    return json.loads(caminho.read_text(encoding="utf-8"))


REGISTRO = Registry(retrieve=lambda uri: Resource.from_contents(carregar_schema(uri), default_specification=DRAFT7))

# O Power BI Desktop grava estes arquivos sem "$schema"; o schema é deduzido pelo tipo do arquivo.
SCHEMA_POR_ARQUIVO = {
    ".pbip": URL_SCHEMAS + "fabric/pbip/pbipProperties/1.0.0/schema.json",
    ".pbir": URL_SCHEMAS + "fabric/item/report/definitionProperties/2.0.0/schema.json",
    ".pbism": URL_SCHEMAS + "fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
}


def schema_publicado(url: str) -> str | None:
    """Devolve a própria URL se o schema já foi publicado; senão, a versão publicada
    mais recente abaixo dela (o Desktop às vezes grava uma versão que a Microsoft
    ainda não publicou no GitHub)."""
    base, versao, arquivo = url.rsplit("/", 2)
    maior, menor, _ = (int(x) for x in versao.split("."))
    for m in range(menor, -1, -1):
        candidata = f"{base}/{maior}.{m}.0/{arquivo}"
        try:
            carregar_schema(candidata)
            return candidata
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
    return None


def checar_json(arquivos: list[Path]) -> None:
    secao("2. JSON bem formado e validado contra o schema oficial")
    for p in arquivos:
        if p.suffix == ".tmdl":
            continue
        try:
            doc = json.loads(p.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            ok(False, f"{rel(p)}: JSON inválido ({e})")
            continue
        url = doc.get("$schema") if isinstance(doc, dict) else None
        nota = ""
        if not url and p.suffix in SCHEMA_POR_ARQUIVO:
            url = SCHEMA_POR_ARQUIVO[p.suffix]
            nota = " (sem $schema no arquivo, como o Desktop grava; schema deduzido pelo tipo)"
        if not url:
            ok(True, f"{rel(p)}: JSON válido (arquivo sem $schema, gerado pelo Desktop)")
            continue
        usada = schema_publicado(url)
        if usada is None:
            ok(False, f"{rel(p)}: nenhum schema publicado para {url}")
            continue
        if usada != url:
            nota = f" (o arquivo pede {url.split('/')[-2]}, ainda não publicado; validado com a mais recente publicada)"
        doc = {**doc, "$schema": usada}
        validador = Draft7Validator(carregar_schema(usada), registry=REGISTRO)
        erros = sorted(validador.iter_errors(doc), key=lambda e: list(e.absolute_path))
        versao = "/".join(usada.split("/")[-3:-1])
        ok(not erros, f"{rel(p)}: schema {versao}{nota}"
           + "".join(f"\n           - {'/'.join(map(str, e.absolute_path)) or '(raiz)'}: {e.message[:200]}" for e in erros[:5]))


# ---------------------------------------------------------------- 3. TMDL
def ler_modelo() -> tuple[dict, dict]:
    return ler_tabela(TABELA)


def ler_tabela(arquivo: Path) -> tuple[dict, dict]:
    texto = arquivo.read_text(encoding="utf-8")
    colunas, medidas, atual = {}, {}, None
    for linha in texto.splitlines():
        if m := re.match(r"^\tcolumn ('[^']+'|\S+)(?: = (.+))?$", linha):
            atual = ("c", m.group(1).strip().strip("'"))
            colunas[atual[1]] = {"expressao": m.group(2).strip()} if m.group(2) else {}
        elif m := re.match(r"^\tmeasure (.+?) = (.+)$", linha):
            atual = ("m", m.group(1).strip().strip("'"))
            medidas[atual[1]] = {"expressao": m.group(2).strip()}
        elif re.match(r"^\t\S", linha):
            atual = None
        elif atual and (m := re.match(r"^\t\t(\w+): (.+)$", linha)):
            (colunas if atual[0] == "c" else medidas)[atual[1]][m.group(1)] = m.group(2).strip()
    return colunas, medidas


def checar_tmdl() -> tuple[dict, dict]:
    secao("3. Modelo (TMDL)")
    for p in sorted((MODELO / "definition").rglob("*.tmdl")):
        com_espaco = [i + 1 for i, l in enumerate(p.read_text(encoding="utf-8").splitlines()) if l.startswith(" ")]
        ok(not com_espaco, f"{rel(p)}: indentação só com tabs" + (f" (linhas com espaço: {com_espaco[:5]})" if com_espaco else ""))

    colunas, medidas = ler_modelo()
    for nome, tipo in COLUNAS_ESPERADAS.items():
        ok(colunas.get(nome, {}).get("dataType") == tipo, f"coluna '{nome}' ({tipo})")
    ok(set(colunas) == set(COLUNAS_ESPERADAS), f"tabela Vendas tem exatamente {len(COLUNAS_ESPERADAS)} colunas")
    for nome, (expr, fmt) in MEDIDAS_ESPERADAS.items():
        m = medidas.get(nome, {})
        ok(m.get("expressao") == expr and m.get("formatString") == fmt, f"medida '{nome}' = {expr}  [{fmt}]")
    ok(set(medidas) == set(MEDIDAS_ESPERADAS) | set(MEDIDAS_IMPACTO),
       f"tabela Vendas tem exatamente {len(MEDIDAS_ESPERADAS)} medidas originais + {len(MEDIDAS_IMPACTO)} da página Impacto")
    nomes = [n.casefold() for n in list(colunas) + list(medidas)]
    ok(len(nomes) == len(set(nomes)), "nenhum nome repetido entre colunas e medidas (ignorando maiúsculas)")
    return colunas, medidas


# ---------------------------------------------------------------- 4. dados
def checar_dados() -> None:
    secao("4. Dados embutidos x xlsx (pandas) x números do case")
    texto = TABELA.read_text(encoding="utf-8")
    linhas = re.findall(r'\{"([^"]*)", "([^"]*)", "([^"]*)", (\d+), ([\d.]+)\}', texto)
    emb = pd.DataFrame(linhas, columns=["Região", "Estado", "Canal", "Unid", "Preço"])
    emb["Unid"] = emb["Unid"].astype(int)
    emb["Preço"] = emb["Preço"].astype(float)
    ok(len(emb) == 75, f"tabela embutida tem {len(emb)} linhas (esperado 75)")

    xl = pd.read_excel(XLSX, sheet_name=0)
    xl.columns = ["Região", "Estado", "Canal", "Unid", "Preço"]
    xl["Preço"] = xl["Preço"].round(2)
    ok(emb.reset_index(drop=True).equals(xl.astype(emb.dtypes.to_dict()).reset_index(drop=True)),
       "as 75 linhas embutidas são idênticas às do xlsx")

    d = emb.copy()
    d["Receita"] = (d["Unid"] * d["Preço"]).round(2)  # mesma conta do M (R$ mil)
    d["EC"] = d["Estado"] + " - " + d["Canal"]
    receita, volume = d["Receita"].sum() / 1000, d["Unid"].sum() / 1000

    def confere(rotulo: str, obtido: str, esperado: str) -> None:
        ok(obtido == esperado, f"{rotulo}: {obtido} (esperado {esperado})")

    confere("Receita (R$ mi)", r1(receita), "683,5")
    confere("Volume (mi un.)", r1(volume), "141,2")
    confere("Preço Médio (R$) ponderado", r2(receita / volume), "4,84")

    reg = d.groupby("Região")["Receita"].sum().div(1000).sort_values(ascending=False)
    esperado_reg = {"Nordeste": "229,8", "Norte": "168,7", "Centro-Oeste": "110,4", "Sudeste": "100,4", "Sul": "74,3"}
    confere("Receita por região (ordem)", ", ".join(reg.index), ", ".join(esperado_reg))
    for r, v in esperado_reg.items():
        confere(f"  {r}", r1(reg[r]), v)
    confere("Norte + Nordeste (% da receita)", str(round(100 * (reg["Norte"] + reg["Nordeste"]) / receita)) + "%", "58%")

    blocos = d.groupby(["Região", "Canal"])["Receita"].sum().div(1000).sort_values(ascending=False)
    confere("Maior célula da matriz", f"{blocos.index[0][0]} x {blocos.index[0][1]} {r1(blocos.iloc[0])}", "Nordeste x Hospitalar 94,5")
    top5 = blocos.head(5)
    confere("5 maiores blocos (% do mercado)", str(round(100 * top5.sum() / receita)) + "%", "51%")
    ok(set(r for r, _ in top5.index) <= {"Norte", "Nordeste"}, "os 5 maiores blocos são todos do Norte/Nordeste")

    ec = d.set_index("EC")["Receita"].div(1000).sort_values(ascending=False)
    confere("Top 1 estado × canal", f"{ec.index[0]} {r1(ec.iloc[0])}", "SC - Varejo 17,6")
    ok(all(n.endswith("Varejo") for n in ec.index[:5]), "as 5 maiores combinações são de Varejo: " + ", ".join(ec.index[:5]))
    print("         Top 10 (para conferir o filtro Top N): " + "; ".join(f"{n} {r1(v)}" for n, v in ec.head(10).items()))
    ok(ec.iloc[9] > ec.iloc[10], f"sem empate na 10ª posição ({r2(ec.iloc[9])} > {r2(ec.iloc[10])})")

    can = d.groupby("Canal").agg(R=("Receita", "sum"), U=("Unid", "sum"), simples=("Preço", "mean"))
    can["P"] = can["R"] / can["U"]
    for c, v in {"Varejo": "5,01", "Hospitalar": "4,86", "Institucional": "4,67"}.items():
        confere(f"Preço médio {c}", r2(can.loc[c, "P"]), v)
    confere("Varejo acima do Institucional", str(round(100 * (can.loc["Varejo", "P"] / can.loc["Institucional", "P"] - 1))) + "%", "7%")
    ok(r2(can.loc["Varejo", "simples"]) != "5,01",
       f"teste anti-média-simples: média simples do Varejo daria {r2(can.loc['Varejo', 'simples'])}, não 5,01")


# ---------------------------------------------------------------- 5. relatório
def campos_usados(no, aliases: dict, achados: list) -> None:
    """Percorre o JSON do visual e coleta (tabela, campo, tipo) de cada Column/Measure."""
    if isinstance(no, dict):
        for item in no.get("From", []) if isinstance(no.get("From"), list) else []:
            if "Entity" in item:
                aliases[item["Name"]] = item["Entity"]
        for tipo in ("Column", "Measure"):
            alvo = no.get(tipo)
            if isinstance(alvo, dict) and "Property" in alvo:
                src = alvo.get("Expression", {}).get("SourceRef", {})
                tabela = src.get("Entity") or aliases.get(src.get("Source"), f"?{src.get('Source')}")
                achados.append((tabela, alvo["Property"], tipo))
        for v in no.values():
            campos_usados(v, aliases, achados)
    elif isinstance(no, list):
        for v in no:
            campos_usados(v, aliases, achados)


def checar_relatorio(colunas: dict, medidas: dict) -> None:
    secao("5. Relatório")
    pbip = json.loads((RAIZ / "Oportunidades.pbip").read_text(encoding="utf-8"))
    ok((RAIZ / pbip["artifacts"][0]["report"]["path"]).is_dir(), ".pbip aponta para a pasta do relatório")
    pbir = json.loads((RELATORIO / "definition.pbir").read_text(encoding="utf-8"))
    ok((RELATORIO / pbir["datasetReference"]["byPath"]["path"]).resolve() == MODELO.resolve(),
       "definition.pbir aponta (byPath) para Oportunidades.SemanticModel")

    rep = json.loads((RELATORIO / "definition" / "report.json").read_text(encoding="utf-8"))
    tema = rep["themeCollection"]["baseTheme"]["name"]
    caminhos = [i["path"] for pk in rep.get("resourcePackages", []) for i in pk["items"] if i["name"] == tema]
    ok(bool(caminhos) and (RELATORIO / "StaticResources" / "SharedResources" / caminhos[0]).is_file(),
       f"arquivo do tema base '{tema}' presente")

    pasta_paginas = RELATORIO / "definition" / "pages"
    meta = json.loads((pasta_paginas / "pages.json").read_text(encoding="utf-8"))
    pastas = sorted(p.name for p in pasta_paginas.iterdir() if p.is_dir())
    ok(meta["pageOrder"] == ORDEM_PAGINAS and pastas == sorted(ORDEM_PAGINAS), f"páginas na ordem {ORDEM_PAGINAS}")
    ok(meta.get("activePageName") == PAGINA["name"], f"o relatório abre na página '{PAGINA['displayName']}'")

    pagina = json.loads((pasta_paginas / PAGINA["name"] / "page.json").read_text(encoding="utf-8"))
    for k, v in PAGINA.items():
        ok(pagina.get(k) == v, f"page.json {k} = {v}")

    nome_valido = re.compile(r"^[\w-]+$")
    visuais = sorted((pasta_paginas / PAGINA["name"] / "visuals").glob("*/visual.json"))
    print(f"         {len(visuais)} visual(is) na página")
    for vj in visuais:
        v = json.loads(vj.read_text(encoding="utf-8"))
        pasta = vj.parent.name
        ok(v["name"] == pasta and bool(nome_valido.match(pasta)), f"{pasta}: nome do visual = nome da pasta")
        pos = v["position"]
        ok(pos["x"] >= 0 and pos["y"] >= 0 and pos["x"] + pos["width"] <= PAGINA["width"]
           and pos["y"] + pos["height"] <= PAGINA["height"], f"{pasta}: dentro da página {pos['x']},{pos['y']} {pos['width']}x{pos['height']}")
        achados: list = []
        campos_usados(v, {}, achados)
        for tabela, campo, tipo in sorted(set(achados)):
            existe = tabela == "Vendas" and (campo in medidas if tipo == "Measure" else campo in colunas)
            ok(existe, f"{pasta}: {tipo.lower()} {tabela}[{campo}] existe no modelo")
        checar_contra_especificacao(pasta, v)

    ok(sorted(vj.parent.name for vj in visuais) == sorted(VISUAIS_ESPERADOS),
       f"a página tem exatamente os {len(VISUAIS_ESPERADOS)} elementos da especificação")


def checar_contra_especificacao(pasta: str, v: dict) -> None:
    if pasta not in VISUAIS_ESPERADOS:
        ok(False, f"{pasta}: visual fora da especificação")
        return
    tipo, (x, y, w, h), papeis = VISUAIS_ESPERADOS[pasta]
    pos = v["position"]
    ok(v["visual"]["visualType"] == tipo and (pos["x"], pos["y"], pos["width"], pos["height"]) == (x, y, w, h),
       f"{pasta}: {tipo} em {x},{y} {w}x{h}")
    estado = v["visual"].get("query", {}).get("queryState", {})
    obtido = {papel: [next(iter(p["field"].values()))["Property"] for p in s["projections"]] for papel, s in estado.items()}
    ok(obtido == papeis, f"{pasta}: campos {papeis or '(nenhum)'}")
    objetos = v["visual"].get("objects", {})
    if pasta == "01_titulo":
        texto = "".join(r["value"] for p in objetos["general"][0]["properties"]["paragraphs"] for r in p["textRuns"])
        estilo = objetos["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["textStyle"]
        ok(texto == TITULO and estilo.get("fontSize") == FONTE_TITULO and estilo.get("fontWeight") == "bold",
           f"{pasta}: texto \"{texto}\" ({estilo.get('fontSize')}, {estilo.get('fontWeight')}; esperado {FONTE_TITULO}, bold)")
    if pasta in ESTILO_SEGMENTACAO:
        modo = objetos.get("data", [{}])[0].get("properties", {}).get("mode", {}).get("expr", {}).get("Literal", {}).get("Value")
        ok(modo == ESTILO_SEGMENTACAO[pasta], f"{pasta}: estilo {modo}")
        selecao = [g["properties"]["filter"] for g in objetos.get("general", []) if "filter" in g.get("properties", {})]
        valores = re.findall(r"'([^']*)'", json.dumps(selecao, ensure_ascii=False))
        ok(not selecao, f"{pasta}: abre sem seleção salva (mostra todos)" + (f" — está salvo com: {', '.join(valores)}" if selecao else ""))


# ---------------------------------------------------------------- 6. acabamento (Fase 3)
def valor(prop):
    """Lê um valor de formatação: Literal direto ou cor sólida."""
    if not isinstance(prop, dict):
        return None
    if "solid" in prop:
        return valor(prop["solid"].get("color"))
    return prop.get("expr", {}).get("Literal", {}).get("Value")


def props(objetos: dict, nome: str, com_seletor: bool = False) -> list[dict]:
    return [e for e in objetos.get(nome, []) if ("selector" in e) == com_seletor]


def checar_acabamento() -> None:
    secao("6. Acabamento (Fase 3)")
    pasta_pagina = RELATORIO / "definition" / "pages" / PAGINA["name"]
    pagina = json.loads((pasta_pagina / "page.json").read_text(encoding="utf-8"))
    fundo = (props(pagina.get("objects", {}), "background") or [{}])[0].get("properties", {})
    ok(valor(fundo.get("color")) == f"'{FUNDO_PAGINA}'" and valor(fundo.get("transparency")) == "0D",
       f"fundo da página {valor(fundo.get('color'))}, transparência {valor(fundo.get('transparency'))} (esperado '{FUNDO_PAGINA}', 0D)")

    for nome in COM_CAIXA_BRANCA:
        v = json.loads((pasta_pagina / "visuals" / nome / "visual.json").read_text(encoding="utf-8"))["visual"]
        vco = v.get("visualContainerObjects", {})
        bg = (props(vco, "background") or [{}])[0].get("properties", {})
        bd = (props(vco, "border") or [{}])[0].get("properties", {})
        raio = valor(bd.get("radius")) or "0D"
        ok(valor(bg.get("show")) == "true" and valor(bg.get("color")) == "'#FFFFFF'"
           and valor(bd.get("show")) == "true" and float(raio.rstrip("DL")) > 0
           and all(k in vco for k in ("padding", "visualHeader")),
           f"{nome}: caixa branca, borda com cantos arredondados (raio {raio}), margens e cabeçalho declarados")

        tit = (props(vco, "title") or [{}])[0].get("properties", {})
        sub = (props(vco, "subTitle") or [{}])[0].get("properties", {})
        esperado = "'" + TITULOS[nome].replace("'", "''") + "'"
        ok(valor(tit.get("show")) == "true" and valor(tit.get("text")) == esperado and valor(sub.get("show")) == "false",
           f"{nome}: título {valor(tit.get('text'))} (subtítulo automático desligado)")

        if nome in CORES:
            col, mapa = CORES[nome]
            obtido = {}
            for e in props(v.get("objects", {}), "dataPoint", com_seletor=True):
                comp = e["selector"]["data"][0].get("scopeId", {}).get("Comparison", {})
                if comp.get("Left", {}).get("Column", {}).get("Property") == col and comp.get("ComparisonKind") == 0:
                    obtido[comp["Right"]["Literal"]["Value"].strip("'")] = valor(e["properties"].get("fill", {})).strip("'")
            ok(obtido == mapa, f"{nome}: cores por {col} " + ", ".join(f"{k} {c}" for k, c in obtido.items()))

        campo_esp, dir_esp = ORDENACAO[nome]
        ordem = v.get("query", {}).get("sortDefinition", {}).get("sort", [{}])
        campo = next(iter(ordem[0].get("field", {"?": {"Property": None}}).values())).get("Property")
        ok(len(ordem) == 1 and campo == campo_esp and ordem[0].get("direction") == dir_esp,
           f"{nome}: ordenado por {campo} {ordem[0].get('direction')}")

        if nome in ROTULOS:
            rot = (props(v.get("objects", {}), "labels") or [{}])[0].get("properties", {})
            casas = valor(rot.get("labelPrecision"))
            ok(valor(rot.get("show")) == "true" and casas == ROTULOS[nome],
               f"{nome}: rótulos de dados ligados" + (f", {casas[:-1]} casas decimais" if casas else ""))
        if nome == "10_colunas_preco":
            eixo = (props(v.get("objects", {}), "valueAxis") or [{}])[0].get("properties", {})
            ok(valor(eixo.get("start")) == "0D", f"{nome}: eixo Y começa em {valor(eixo.get('start'))}")

    checar_top_n(pasta_pagina / "visuals" / "09_barras_top10" / "visual.json")
    checar_gradiente(pasta_pagina / "visuals" / "07_matriz" / "visual.json")


def checar_gradiente(arquivo: Path) -> None:
    v = json.loads(arquivo.read_text(encoding="utf-8"))["visual"]
    ref = v["query"]["queryState"]["Values"]["projections"][0]["queryRef"]
    regras = [e for e in v.get("objects", {}).get("values", [])
              if "FillRule" in e.get("properties", {}).get("backColor", {}).get("solid", {}).get("color", {}).get("expr", {})]
    if not ok(len(regras) == 1, f"07_matriz: um gradiente no fundo das células ({len(regras)} encontrado)"):
        return
    e = regras[0]
    regra = e["properties"]["backColor"]["solid"]["color"]["expr"]["FillRule"]
    grad = regra["FillRule"].get("linearGradient2", {})
    paradas = {k: grad.get(k, {}).get("color", {}) for k in ("min", "max")}
    ok(all("expr" not in c for c in paradas.values()), "07_matriz: cores de parada sem 'expr' duplicado")
    cores = {k: c.get("Literal", {}).get("Value") for k, c in paradas.items()}
    ok(cores == {"min": "'#F2F8F9'", "max": "'#4DB6AC'"}, f"07_matriz: gradiente menor {cores['min']} -> maior {cores['max']}")
    ok(regra["Input"].get("SelectRef", {}).get("ExpressionName") == ref and e["selector"].get("metadata") == ref,
       f"07_matriz: gradiente calculado sobre {ref}")
    ok(e["selector"]["data"][0]["dataViewWildcard"]["matchingOption"] == 1, "07_matriz: gradiente só nas células (totais sem cor)")


def checar_top_n(arquivo: Path) -> None:
    filtros = json.loads(arquivo.read_text(encoding="utf-8")).get("filterConfig", {}).get("filters", [])
    topn = [f for f in filtros if f.get("type") == "TopN"]
    if not ok(len(topn) == 1, f"09_barras_top10: um filtro Top N no visual ({len(topn)} encontrado)"):
        return
    f = topn[0]
    try:
        sub = next(x for x in f["filter"]["From"] if "Expression" in x)["Expression"]["Subquery"]["Query"]
        ordem = sub["OrderBy"][0]
        agg = ordem["Expression"]["Aggregation"]
        cond = f["filter"]["Where"][0]["Condition"]["In"]
        resumo = (f["field"]["Column"]["Property"], sub["Top"], ordem["Direction"], agg["Function"],
                  agg["Expression"]["Column"]["Property"], cond["Expressions"][0]["Column"]["Property"],
                  cond["Table"]["SourceRef"]["Source"])
    except (KeyError, IndexError, StopIteration, TypeError) as e:
        ok(False, f"09_barras_top10: estrutura do Top N incompleta ({e!r})")
        return
    ok(resumo == ("Estado - Canal", 10, 2, 0, "Receita (R$ mil)", "Estado - Canal", "subquery"),
       f"09_barras_top10: Top {resumo[1]} de '{resumo[0]}' pela soma de '{resumo[4]}' (direção {resumo[2]} = maiores)")


# ---------------------------------------------------------------- 7. modelo da página Impacto
def checar_modelo_impacto() -> dict:
    secao("7. Modelo da página Impacto (TMDL)")
    modelo = {}
    for arq in sorted((MODELO / "definition" / "tables").glob("*.tmdl")):
        nome = re.search(r"^table (.+)$", arq.read_text(encoding="utf-8"), re.M).group(1).strip().strip("'")
        modelo[nome] = dict(zip(("colunas", "medidas"), ler_tabela(arq)))
    ok(sorted(modelo) == ["Cenarios", "Descontos", "Vendas"], f"tabelas do modelo: {', '.join(sorted(modelo))}")
    ok(not (MODELO / "definition" / "relationships.tmdl").exists(),
       "nenhum relacionamento: Cenarios e Descontos ficam desconectadas")

    for tabela, esperado in COLUNAS_P2.items():
        cols = modelo.get(tabela, {}).get("colunas", {})
        ok(set(cols) == set(esperado), f"{tabela}: colunas {', '.join(cols)}")
        for nome, (tipo, fmt, ordem) in esperado.items():
            c = cols.get(nome, {})
            ok((c.get("dataType"), c.get("formatString"), c.get("sortByColumn"), c.get("summarizeBy")) == (tipo, fmt, ordem, "none"),
               f"{tabela}[{nome}]: {tipo}" + (f", formato {fmt}" if fmt else "") + (f", ordenada por {ordem}" if ordem else "")
               + ", não resume")
    share = modelo.get("Cenarios", {}).get("colunas", {}).get("Share total", {})
    ok(share.get("expressao") == EXPR_SHARE_TOTAL, f"coluna calculada Share total = {share.get('expressao')}")

    medidas = modelo["Vendas"]["medidas"]
    for nome, (expr, fmt) in MEDIDAS_IMPACTO.items():
        m = medidas.get(nome, {})
        ok(m.get("expressao") == expr and m.get("formatString") == fmt, f"medida '{nome}' = {expr}  [{fmt}]")
    com_numero = [n for n in MEDIDAS_IMPACTO if n not in ("Fator Preço Genérico", "Margem Genérico")
                  and re.search(r"\d\.\d", medidas.get(n, {}).get("expressao", ""))]
    ok(not com_numero, "premissas (0,65 e 0,30) escritas só em Fator Preço Genérico e Margem Genérico"
       + (f" — também em: {', '.join(com_numero)}" if com_numero else ""))
    todas_medidas = [n.casefold() for t in modelo.values() for n in t["medidas"]]
    todas_colunas = {n.casefold() for t in modelo.values() for n in t["colunas"]}
    ok(len(todas_medidas) == len(set(todas_medidas)) and not set(todas_medidas) & todas_colunas,
       "nomes de medidas únicos e diferentes de qualquer coluna (ignorando maiúsculas)")
    return modelo


# ---------------------------------------------------------------- 8. números da página Impacto
def pct1(x: Decimal) -> str:
    return str((x * 100).quantize(Decimal("0.1"), ROUND_HALF_UP)).replace(".", ",") + "%"


def checar_dados_impacto(modelo: dict) -> None:
    secao("8. Números da página Impacto (pandas, Região = Nordeste) x verificação do prompt")
    pasta = MODELO / "definition" / "tables"
    cen = [(c, int(o), p, s) for c, o, p, s in re.findall(
        r'\{"([^"]+)", (\d+), ([\d.]+), ([\d.]+)\}', (pasta / "Cenarios.tmdl").read_text(encoding="utf-8"))]
    des = [(d, r) for d, r in re.findall(r'\{([\d.]+), "([^"]+)"\}', (pasta / "Descontos.tmdl").read_text(encoding="utf-8"))]
    ok(cen == LINHAS_CENARIOS, f"Cenarios embutida: {cen}")
    ok(des == LINHAS_DESCONTOS, f"Descontos embutida: {des}")

    medidas = modelo["Vendas"]["medidas"]
    fator = Decimal(medidas["Fator Preço Genérico"]["expressao"])
    margem = Decimal(medidas["Margem Genérico"]["expressao"])
    xl = pd.read_excel(XLSX, sheet_name=0)
    xl.columns = ["Região", "Estado", "Canal", "Unid", "Preço"]
    ne = xl[xl["Região"] == "Nordeste"]
    unid = [Decimal(int(u)) for u in ne["Unid"]]
    rec_linha = [u * Decimal(str(round(p, 2))) for u, p in zip(unid, ne["Preço"])]  # R$ mil, exato
    receita, volume = sum(rec_linha) / 1000, sum(unid) / 1000
    preco = receita / volume
    generico = receita * fator
    share = {c: Decimal(p) * Decimal(s) for c, _, p, s in cen}
    sc = share.get(CENARIO_CARTOES_P2, Decimal(0))  # share do cenário usado nos cartões 3-4 e nas barras por canal
    print(f"         Nordeste: {len(ne)} linhas | receita {receita} | volume {volume} | preço médio {preco:.4f}"
          f" | premissas lidas do modelo: fator {fator}, margem {margem}, share {CENARIO_CARTOES_P2.lower()} {sc}")

    def confere(rotulo: str, obtido: str, esperado: str) -> None:
        ok(obtido == esperado, f"{rotulo}: {obtido} (esperado {esperado})")

    obtidos = (r1(receita), r1(generico), r1(generico * sc), r1(generico * sc * margem))
    rotulo_c = CENARIO_CARTOES_P2.lower()
    for nome, o, e in zip(("Mercado atual", "Mercado a preço de genérico", f"Receita EMS ({rotulo_c})",
                           f"Lucro bruto EMS ({rotulo_c})"), obtidos, ESPERADO_P2["cartoes"]):
        confere(f"cartão {nome} (R$ mi)", o, e)
    lucro_cen = {}
    for c, (e_share, e_unid, e_rec, e_luc) in ESPERADO_P2["cenarios"].items():
        s = share.get(c, Decimal(0))
        lucro_cen[c] = generico * s * margem
        confere(f"tabela {c}: share | unidades | receita | lucro",
                " | ".join((pct1(s), r1(volume * s), r1(generico * s), r1(lucro_cen[c]))),
                " | ".join((e_share, e_unid, e_rec, e_luc)))
    lucro_canal = {}
    for canal, e in ESPERADO_P2["canais"].items():
        rec_c = sum(r for r, c in zip(rec_linha, ne["Canal"]) if c == canal) / 1000
        lucro_canal[canal] = rec_c * fator * sc * margem
        confere(f"lucro bruto {rotulo_c} {canal} (R$ mi)", r1(lucro_canal[canal]), e)
    print(f"         soma exata dos canais {sum(lucro_canal.values()):.4f} = cartão {r1(generico * sc * margem)}"
          f" (os rótulos arredondados somam {sum(Decimal(r1(v).replace(',', '.')) for v in lucro_canal.values())})")
    custo = preco * fator * (1 - margem)
    lucro_un = {r: preco * (1 - Decimal(d)) - custo for d, r in des}
    for r, e in ESPERADO_P2["descontos"].items():
        confere(f"lucro por unidade com desconto {r} (R$)", r2(lucro_un.get(r, 0)), e)

    # Números escritos nos títulos x números calculados
    t7 = re.findall(r"R\$ (\d+,\d+)", TITULOS_P2["07_colunas_lucro_cenario"])
    confere("título 07: lucro conservador -> com foco", " -> ".join(t7),
            f"{r1(lucro_cen['Conservador'])} -> {r1(lucro_cen['Com foco no Nordeste'])}")
    maior = max(lucro_canal, key=lucro_canal.get)
    ok(maior in TITULOS_P2["08_barras_lucro_canal"], f"título 08: canal com maior lucro bruto = {maior}")

    # Textos fixos das legendas que dependem de premissas x valores calculados a partir do modelo
    def pct0(x: Decimal) -> str:
        return str((x * 100).quantize(Decimal("1"), ROUND_HALF_UP)) + "%"

    origens = [(pct0(fator), "Fator Preço Genérico"), (pct0(margem), "Margem Genérico"),
               (pct0(1 - margem), "1 - Margem Genérico (custo = 70% do preço de genérico)"),
               (pct1(sc), f"Share {CENARIO_CARTOES_P2} (Cenarios, linha {CENARIO_CARTOES_P2})"),
               ("R$ " + r2(custo), "Custo unitário (R$) no Nordeste")]
    for c, _, p, s in cen:  # penetração e participação de cada cenário (tabela Cenarios)
        origens += [(pct0(Decimal(p)), f"penetração do cenário {c}"), (pct0(Decimal(s)), f"participação EMS do cenário {c}")]
    print("         Textos fixos nas legendas que dependem de premissas (atualizar se a premissa mudar):")
    for nome, txt in LEGENDAS_P2.items():
        for numero in re.findall(r"R\$ \d+,\d+|\d+(?:,\d+)?%", txt):
            achou = sorted({d for n, d in origens if n == numero})
            ok(bool(achou), f"  {nome}: \"{numero}\" = " + (" ou ".join(achou) or "NÃO bate com nenhuma premissa do modelo"))


# ---------------------------------------------------------------- 9. página Impacto
def checar_pagina_impacto(modelo: dict) -> None:
    secao("9. Página Impacto (relatório)")
    pasta = RELATORIO / "definition" / "pages" / PAGINA2["name"]
    pagina = json.loads((pasta / "page.json").read_text(encoding="utf-8"))
    for k, v in PAGINA2.items():
        ok(pagina.get(k) == v, f"page.json {k} = {v}")
    fundo = (props(pagina.get("objects", {}), "background") or [{}])[0].get("properties", {})
    ok(valor(fundo.get("color")) == f"'{FUNDO_PAGINA}'" and valor(fundo.get("transparency")) == "0D",
       f"fundo da página {valor(fundo.get('color'))}, transparência {valor(fundo.get('transparency'))}")

    filtros = pagina.get("filterConfig", {}).get("filters", [])
    try:
        f = filtros[0]
        cond = f["filter"]["Where"][0]["Condition"]["In"]
        resumo = (len(filtros), f["type"], f["field"]["Column"]["Expression"]["SourceRef"]["Entity"],
                  f["field"]["Column"]["Property"], cond["Expressions"][0]["Column"]["Property"],
                  [v["Literal"]["Value"] for linha in cond["Values"] for v in linha], f.get("isLockedInViewMode"))
    except (KeyError, IndexError, TypeError) as e:
        resumo = repr(e)
    ok(resumo == (1, "Categorical", "Vendas", "Região", "Região", ["'Nordeste'"], True),
       f"filtro de página: Vendas[Região] = Nordeste (todos os canais), travado — {resumo}")

    nome_valido = re.compile(r"^[\w-]+$")
    visuais = sorted((pasta / "visuals").glob("*/visual.json"))
    print(f"         {len(visuais)} visual(is) na página Impacto (fase {FASE_PAGINA2})")
    for vj in visuais:
        v = json.loads(vj.read_text(encoding="utf-8"))
        nome = vj.parent.name
        ok(v["name"] == nome and bool(nome_valido.match(nome)), f"{nome}: nome do visual = nome da pasta")
        achados: list = []
        campos_usados(v, {}, achados)
        for tabela, campo, tipo in sorted(set(achados)):
            grupo = "medidas" if tipo == "Measure" else "colunas"
            ok(campo in modelo.get(tabela, {}).get(grupo, {}), f"{nome}: {tipo.lower()} {tabela}[{campo}] existe no modelo")
        if nome not in VISUAIS_P2:
            ok(False, f"{nome}: visual fora da especificação")
            continue
        tipo, (x, y, w, h), papeis = VISUAIS_P2[nome]
        pos = v["position"]
        ok(v["visual"]["visualType"] == tipo and (pos["x"], pos["y"], pos["width"], pos["height"]) == (x, y, w, h),
           f"{nome}: {tipo} em {x},{y} {w}x{h}")
        estado = v["visual"].get("query", {}).get("queryState", {})
        obtido = {p: [next(iter(pr["field"].values()))["Property"] for pr in s["projections"]] for p, s in estado.items()}
        ok(obtido == papeis, f"{nome}: campos {papeis or '(nenhum)'}")
        checar_extras_impacto(nome, v["visual"])
    esperados = ELEMENTOS_POR_FASE[FASE_PAGINA2]
    ok(sorted(vj.parent.name for vj in visuais) == sorted(esperados),
       f"fase {FASE_PAGINA2}: a página tem exatamente os {len(esperados)} elementos previstos para esta fase")
    for nome in REMOVIDOS_P2:
        ok(not (pasta / "visuals" / nome).exists(), f"{nome}: removido (a pasta não existe mais)")
    caixas_fontes = [vj.parent.name for vj in visuais
                     if json.loads(vj.read_text(encoding="utf-8"))["visual"]["visualType"] == "textbox"
                     and "Fontes" in vj.read_text(encoding="utf-8")]
    ok(not caixas_fontes, "não existe caixa de texto \"Fontes\"" + (f" — encontrada: {caixas_fontes}" if caixas_fontes else ""))

    # Geometria calculada a partir das posições gravadas: dentro de 1280x720 e sem sobreposição
    caixas = {}
    for vj in visuais:
        p = json.loads(vj.read_text(encoding="utf-8"))["position"]
        caixas[vj.parent.name] = (p["x"], p["y"], p["x"] + p["width"], p["y"] + p["height"])
    fora = [n for n, (x0, y0, x1, y1) in caixas.items()
            if x0 < 0 or y0 < 0 or x1 > PAGINA2["width"] or y1 > PAGINA2["height"]]
    ok(not fora, f"todos os {len(caixas)} visuais dentro de {PAGINA2['width']}x{PAGINA2['height']}"
       + (f" — fora: {fora}" if fora else ""))
    nomes = sorted(caixas)
    sobrepostos = [(a, b) for i, a in enumerate(nomes) for b in nomes[i + 1:]
                   if min(caixas[a][2], caixas[b][2]) > max(caixas[a][0], caixas[b][0])
                   and min(caixas[a][3], caixas[b][3]) > max(caixas[a][1], caixas[b][1])]
    ok(not sobrepostos, "nenhum par de visuais sobreposto" + (f" — sobrepostos: {sobrepostos}" if sobrepostos else ""))
    for n in nomes:
        x0, y0, x1, y1 = caixas[n]
        print(f"           {n:<26} x {x0:>4}-{x1:<4}  y {y0:>3}-{y1:<3}")


def checar_extras_impacto(nome: str, v: dict) -> None:
    objetos = v.get("objects", {})
    if nome in TEXTOS_P2:
        texto_esp, tamanho, negrito = TEXTOS_P2[nome]
        runs = [r for p in objetos.get("general", [{}])[0].get("properties", {}).get("paragraphs", []) for r in p["textRuns"]]
        texto = "".join(r["value"] for r in runs)
        estilo = runs[0].get("textStyle", {}) if runs else {}
        ok(texto == texto_esp, f"{nome}: texto exato da especificação ({len(texto)} caracteres)")
        if tamanho:
            ok(estilo.get("fontSize") == tamanho and (estilo.get("fontWeight") == "bold") == negrito,
               f"{nome}: fonte {estilo.get('fontSize')}, {estilo.get('fontWeight')} (esperado {tamanho}, negrito)")
    if nome == "06_tabela_cenarios":
        tot = (props(objetos, "total") or [{}])[0].get("properties", {})
        ok(valor(tot.get("totals")) == "false", f"{nome}: sem linha de total")
    if nome in ORDENACAO_P2:
        campo_esp, dir_esp = ORDENACAO_P2[nome]
        ordem = v.get("query", {}).get("sortDefinition", {}).get("sort", [{}])
        campo = next(iter(ordem[0].get("field", {"?": {"Property": None}}).values())).get("Property")
        ok(len(ordem) == 1 and campo == campo_esp and ordem[0].get("direction") == dir_esp,
           f"{nome}: ordenado por {campo} {ordem[0].get('direction')}")


# ---------------------------------------------------------------- 11. acabamento da página Impacto
def checar_acabamento_impacto() -> None:
    secao("11. Acabamento da página Impacto (Fase 3)")
    pasta = RELATORIO / "definition" / "pages" / PAGINA2["name"] / "visuals"

    def ler(nome: str) -> dict:
        return json.loads((pasta / nome / "visual.json").read_text(encoding="utf-8"))["visual"]

    def legenda_ok(nome: str, vco: dict, titulo_props: dict) -> None:
        sub = (props(vco, "subTitle") or [{}])[0].get("properties", {})
        tam_t = float((valor(titulo_props.get("fontSize")) or "0D").rstrip("D"))
        tam_s = float((valor(sub.get("fontSize")) or "99D").rstrip("D"))
        if nome in LEGENDAS_APLICADAS:
            esperado = "'" + LEGENDAS_P2[nome].replace("'", "''") + "'"
            ok(valor(sub.get("show")) == "true" and valor(sub.get("text")) == esperado,
               f"{nome}: legenda {valor(sub.get('text'))}")
            ok(tam_s < tam_t and valor(sub.get("fontColor")) == f"'{CINZA_LEGENDA}'",
               f"{nome}: legenda cinza {valor(sub.get('fontColor'))}, fonte {tam_s:g} < título {tam_t:g}")
        else:
            ok(valor(sub.get("show")) == "false", f"{nome}: (legenda ainda não aplicada) subtítulo desligado")

    def caixa_ok(nome: str, vco: dict) -> None:
        bg = (props(vco, "background") or [{}])[0].get("properties", {})
        bd = (props(vco, "border") or [{}])[0].get("properties", {})
        raio = valor(bd.get("radius")) or "0D"
        ok(valor(bg.get("show")) == "true" and valor(bg.get("color")) == "'#FFFFFF'" and valor(bd.get("show")) == "true"
           and float(raio.rstrip("DL")) > 0 and all(k in vco for k in ("padding", "visualHeader")),
           f"{nome}: caixa branca, borda arredondada (raio {raio}), margens e cabeçalho declarados")

    for nome, titulo in CARTOES_TITULOS_P2.items():
        v = ler(nome)
        vco = v.get("visualContainerObjects", {})
        tit = (props(vco, "title") or [{}])[0].get("properties", {})
        ok(valor(tit.get("show")) == "true" and valor(tit.get("text")) == "'" + titulo.replace("'", "''") + "'",
           f"{nome}: título {valor(tit.get('text'))} (fonte {valor(tit.get('fontSize'))}, negrito {valor(tit.get('bold'))})")
        legenda_ok(nome, vco, tit)
        internos = {o: valor((v.get("objects", {}).get(o) or [{}])[0].get("properties", {}).get("show")) for o in ("label", "outline")}
        ok(internos == {"label": "false", "outline": "false"}, f"{nome}: rótulo e borda internos do cartão desligados")
        caixa_ok(nome, vco)

    for nome, titulo in TITULOS_P2.items():
        vco = ler(nome).get("visualContainerObjects", {})
        bg = (props(vco, "background") or [{}])[0].get("properties", {})
        bd = (props(vco, "border") or [{}])[0].get("properties", {})
        tit = (props(vco, "title") or [{}])[0].get("properties", {})
        sub = (props(vco, "subTitle") or [{}])[0].get("properties", {})
        raio = valor(bd.get("radius")) or "0D"
        ok(valor(bg.get("show")) == "true" and valor(bg.get("color")) == "'#FFFFFF'" and valor(bd.get("show")) == "true"
           and float(raio.rstrip("DL")) > 0 and all(k in vco for k in ("padding", "visualHeader")),
           f"{nome}: caixa branca, borda arredondada (raio {raio}), margens e cabeçalho declarados")
        esperado = "'" + titulo.replace("'", "''") + "'"
        ok(valor(tit.get("show")) == "true" and valor(tit.get("text")) == esperado, f"{nome}: título {valor(tit.get('text'))}")
        legenda_ok(nome, vco, tit)

    # Legenda da tabela: fonte pequena (9 a 10) e a legenda antiga não existe em nenhum arquivo do relatório
    sub_tab = (props(ler("06_tabela_cenarios").get("visualContainerObjects", {}), "subTitle") or [{}])[0].get("properties", {})
    tam_tab = float((valor(sub_tab.get("fontSize")) or "0D").rstrip("D"))
    ok(9 <= tam_tab <= 10, f"06_tabela_cenarios: legenda em fonte {tam_tab:g} (pedido: 9 a 10)")
    vco_tab = ler("06_tabela_cenarios").get("visualContainerObjects", {})
    pad = (props(vco_tab, "padding") or [{}])[0].get("properties", {})
    esp = (props(vco_tab, "spacing") or [{}])[0].get("properties", {})
    margens = tuple(valor(pad.get(k)) for k in ("top", "bottom", "left", "right"))
    espacos = tuple(valor(esp.get(k)) for k in ("spaceBelowTitle", "spaceBelowSubTitle", "spaceBelowTitleArea"))
    ok(margens == ("8D", "8D", "16D", "16D") and espacos == ("2D", "2D", "2D"),
       f"06_tabela_cenarios: caixa compacta (margens cima/baixo/esq/dir {margens}, espaços {espacos}) — evita barra de rolagem")
    com_antiga = [rel(a) for a in RELATORIO.rglob("*.json") if LEGENDA_ANTIGA_TABELA in a.read_text(encoding="utf-8")]
    ok(not com_antiga, "a legenda antiga da tabela (\"Share = genéricos levam × ...\") não existe em nenhum arquivo do relatório"
       + (f" — encontrada em: {com_antiga}" if com_antiga else ""))

    # Botão do link da fonte: ação URL da Web com a URL exata, texto do botão, altura segura (40 a 56 px)
    bot = json.loads((pasta / "11_link_fonte" / "visual.json").read_text(encoding="utf-8"))
    link = (props(bot["visual"].get("visualContainerObjects", {}), "visualLink") or [{}])[0].get("properties", {})
    txt_bot = [e for e in bot["visual"].get("objects", {}).get("text", []) if e.get("selector", {}).get("id") == "default"]
    rotulo = valor(txt_bot[0]["properties"].get("text")) if txt_bot else None
    ok((valor(link.get("show")), valor(link.get("type")), valor(link.get("webUrl"))) == ("true", "'WebUrl'", f"'{URL_FONTE}'"),
       f"11_link_fonte: ação URL da Web -> {valor(link.get('webUrl'))}")
    ok(rotulo == f"'{ROTULO_BOTAO_FONTE}'" and 40 <= bot["position"]["height"] <= 56,
       f"11_link_fonte: texto do botão {rotulo}, altura {bot['position']['height']} px")

    # Cores (especificação): 07 #0E4A5C, 08 por canal, 09 #D9701A
    for nome, hexa in (("07_colunas_lucro_cenario", "#0E4A5C"),):
        dp = (props(ler(nome).get("objects", {}), "dataPoint") or [{}])[0].get("properties", {})
        ok(valor(dp.get("defaultColor")) == f"'{hexa}'", f"{nome}: cor única {valor(dp.get('defaultColor'))}")
    obtido = {}
    for e in props(ler("08_barras_lucro_canal").get("objects", {}), "dataPoint", com_seletor=True):
        comp = e["selector"]["data"][0].get("scopeId", {}).get("Comparison", {})
        if comp.get("Left", {}).get("Column", {}).get("Property") == "Canal" and comp.get("ComparisonKind") == 0:
            obtido[comp["Right"]["Literal"]["Value"].strip("'")] = valor(e["properties"].get("fill", {})).strip("'")
    ok(obtido == CORES_CANAL, "08_barras_lucro_canal: cores por canal " + ", ".join(f"{k} {c}" for k, c in obtido.items()))

    # Rótulos e eixos
    for nome, casas_esp in ROTULOS_P2.items():
        objetos = ler(nome).get("objects", {})
        rot = (props(objetos, "labels") or [{}])[0].get("properties", {})
        casas = valor(rot.get("labelPrecision"))
        ok(valor(rot.get("show")) == "true" and casas == casas_esp,
           f"{nome}: rótulos de dados ligados" + (f", {casas[:-1]} casas decimais" if casas else ""))
    eixo8 = (props(ler("08_barras_lucro_canal").get("objects", {}), "valueAxis") or [{}])[0].get("properties", {})
    ok(valor(eixo8.get("show")) == "false", "08_barras_lucro_canal: eixo de valores escondido (os rótulos mostram os números)")

    # Tabela: nomes de exibição, cabeçalhos com quebra de linha, larguras fixas que cabem sem rolagem
    tab = ler("06_tabela_cenarios")
    projs = tab["query"]["queryState"]["Values"]["projections"]
    cab = [p.get("displayName") for p in projs]
    ok(cab == CABECALHOS_P2, f"06_tabela_cenarios: cabeçalhos {cab}")
    ch = (props(tab.get("objects", {}), "columnHeaders") or [{}])[0].get("properties", {})
    vl = (props(tab.get("objects", {}), "values") or [{}])[0].get("properties", {})
    ok((valor(ch.get("autoSizeColumnWidth")), valor(ch.get("columnAdjustment")), valor(ch.get("wordWrap")),
        valor(ch.get("fontSize")), valor(vl.get("fontSize"))) == ("false", "'fixedWidth'", "true", "9D", "9D"),
       "06_tabela_cenarios: largura automática desligada (fixedWidth), quebra de linha nos cabeçalhos, fonte 9")
    larguras = {e["selector"].get("metadata"): float(valor(e["properties"].get("value")).rstrip("D"))
                for e in props(tab.get("objects", {}), "columnWidth", com_seletor=True)}
    refs = [p["queryRef"] for p in projs]
    ok(set(larguras) == set(refs), f"06_tabela_cenarios: largura fixa para as {len(refs)} colunas")
    total = sum(larguras.values())
    ok(total <= LARGURA_UTIL_TABELA, f"06_tabela_cenarios: soma das larguras {total:.0f} px <= {LARGURA_UTIL_TABELA} px"
       f" úteis (sem barra de rolagem) — " + ", ".join(f"{int(larguras.get(r, 0))}" for r in refs))


# ---------------------------------------------------------------- 12. itens do modelo sem uso (informativo)
def listar_sem_uso(modelo: dict) -> None:
    secao("12. Itens do modelo sem uso nos visuais (informativo, não é falha)")
    achados: list = []
    for arq in sorted((RELATORIO / "definition" / "pages").rglob("*.json")):
        if arq.name in ("visual.json", "page.json"):
            campos_usados(json.loads(arq.read_text(encoding="utf-8")), {}, achados)
    medidas = {n: m.get("expressao", "") for t in modelo.values() for n, m in t["medidas"].items()}
    colunas = {(t, c): d for t, tab in modelo.items() for c, d in tab["colunas"].items()}
    usadas_m, usadas_c = set(), {(t, c) for t, c, tipo in achados if tipo == "Column"}
    pendentes = [c for t, c, tipo in achados if tipo == "Measure"]
    while pendentes:  # medidas usadas por visuais + tudo o que elas referenciam
        m = pendentes.pop()
        if m in usadas_m or m not in medidas:
            continue
        usadas_m.add(m)
        pendentes += re.findall(r"(?<![\w\]'])\[([^\]]+)\]", medidas[m])
        usadas_c |= {(t, c) for t, c in re.findall(r"(\w+)\[([^\]]+)\]", medidas[m])}
    for (t, c) in list(usadas_c):  # ordenação por outra coluna e colunas calculadas também contam como uso
        d = colunas.get((t, c), {})
        if d.get("sortByColumn"):
            usadas_c.add((t, d["sortByColumn"]))
        usadas_c |= {(tt, cc) for tt, cc in re.findall(r"(\w+)\[([^\]]+)\]", d.get("expressao", ""))}
    sem_m = [m for m in medidas if m not in usadas_m]
    sem_c = [f"{t}[{c}]" for (t, c) in colunas if (t, c) not in usadas_c]
    print("         Medidas sem uso: " + (", ".join(sem_m) or "(nenhuma)"))
    print("         Colunas sem uso em visuais/medidas: " + (", ".join(sem_c) or "(nenhuma)"))
    tabelas_sem_uso = [t for t in modelo if not any(tt == t for tt, _ in usadas_c) and not modelo[t]["medidas"]]
    print("         Tabelas inteiras sem uso: " + (", ".join(tabelas_sem_uso) or "(nenhuma)"))


# ---------------------------------------------------------------- 10. página 1 intacta
def git(*args: str) -> str:
    r = subprocess.run(["git", "-C", str(RAIZ), *args], capture_output=True, text=True, encoding="utf-8")
    return r.stdout


def checar_pagina1_intacta() -> None:
    secao("10. Página 1 intacta (diff do git contra o branch main)")
    if not git("rev-parse", "--verify", "--quiet", "main").strip():
        ok(False, "branch main não encontrado")
        return
    mudou = git("-c", "core.quotepath=false", "diff", "--name-only", "main", "--", *PROTEGIDOS).split("\n")
    mudou = [m for m in mudou if m]
    ok(not mudou, "página 1 (visuais, page.json), report.json, tema e arquivos do projeto idênticos ao main"
       + (f" — mudaram: {', '.join(mudou)}" if mudou else ""))
    for caminho in SO_ACRESCIMOS:
        diff = git("diff", "--unified=0", "main", "--", caminho).splitlines()
        mais = [l for l in diff if l.startswith("+") and not l.startswith("+++")]
        menos = [l for l in diff if l.startswith("-") and not l.startswith("---")]
        ok(not menos, f"{caminho.split('/')[-1]}: só acréscimos ({len(mais)} linhas novas, {len(menos)} removidas/alteradas)")
    if git("rev-parse", "--verify", "--quiet", BASE_AJUSTES).strip():
        vendas = "Oportunidades.SemanticModel/definition/tables/Vendas.tmdl"
        modelo_mudou = [m for m in git("-c", "core.quotepath=false", "diff", "--name-only", BASE_AJUSTES, "--",
                                       "Oportunidades.SemanticModel").split("\n") if m and m != vendas]
        ok(not modelo_mudou, f"modelo (tabelas Cenarios/Descontos e arquivos do modelo) idêntico ao branch {BASE_AJUSTES}"
           + (f" — mudaram: {', '.join(modelo_mudou)}" if modelo_mudou else ""))
        diff = git("diff", "--unified=0", BASE_AJUSTES, "--", vendas).splitlines()
        menos = [l for l in diff if l.startswith("-") and not l.startswith("---")]
        mais = [l for l in diff if l.startswith("+") and not l.startswith("+++")]
        ok(not menos, f"Vendas.tmdl em relação a {BASE_AJUSTES}: só acréscimos ({len(mais)} linhas novas, "
           f"{len(menos)} removidas/alteradas) — medidas existentes intactas")
    print("         Tudo o que mudou em relação ao main (arquivos versionados + novos):")
    for linha in git("-c", "core.quotepath=false", "diff", "--stat", "main").splitlines():
        print("           " + linha)
    for linha in git("-c", "core.quotepath=false", "status", "--short", "--untracked-files=all").splitlines():
        if linha.startswith("??"):
            print("           novo: " + linha[3:])


def main() -> int:
    arquivos = arquivos_do_projeto()
    checar_codificacao(arquivos)
    checar_json(arquivos)
    colunas, medidas = checar_tmdl()
    checar_dados()
    checar_relatorio(colunas, medidas)
    checar_acabamento()
    modelo = checar_modelo_impacto()
    checar_dados_impacto(modelo)
    checar_pagina_impacto(modelo)
    checar_pagina1_intacta()
    checar_acabamento_impacto()
    listar_sem_uso(modelo)
    print("\n" + ("RESULTADO: TUDO OK" if not falhas else f"RESULTADO: {len(falhas)} FALHA(S)"))
    return 0 if not falhas else 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
