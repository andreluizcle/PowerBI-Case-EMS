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
"""
import json
import re
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
    texto = TABELA.read_text(encoding="utf-8")
    colunas, medidas, atual = {}, {}, None
    for linha in texto.splitlines():
        if m := re.match(r"^\tcolumn (.+)$", linha):
            atual = ("c", m.group(1).strip().strip("'"))
            colunas[atual[1]] = {}
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
    ok(set(medidas) == set(MEDIDAS_ESPERADAS), f"tabela Vendas tem exatamente {len(MEDIDAS_ESPERADAS)} medidas")
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
    ok(meta["pageOrder"] == pastas == [PAGINA["name"]], f"uma página só, pasta '{PAGINA['name']}'")
    ok(meta.get("activePageName") in meta["pageOrder"], "página ativa existe")

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


def main() -> int:
    arquivos = arquivos_do_projeto()
    checar_codificacao(arquivos)
    checar_json(arquivos)
    colunas, medidas = checar_tmdl()
    checar_dados()
    checar_relatorio(colunas, medidas)
    checar_acabamento()
    print("\n" + ("RESULTADO: TUDO OK" if not falhas else f"RESULTADO: {len(falhas)} FALHA(S)"))
    return 0 if not falhas else 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
