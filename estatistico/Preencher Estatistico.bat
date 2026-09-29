0<0# : ^
'''
@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
python "%~f0" %*
if errorlevel 1 (
  echo.
  echo Se reclamou de "xlrd" ou "openpyxl", rode uma vez no Prompt:
  echo    python -m pip install openpyxl xlrd
)
echo.
pause
exit /b
'''
"""
========================================================================
 PREENCHER ESTATISTICO + FOLHA DE ROSTO
========================================================================

Le quatro relatorios do FrontOffice e preenche as duas planilhas:

  RDS_RELATORIO        -> Lanc. Estatistico: ESTATISTICAS e RECEITA do dia
  on_the_book          -> Lanc. Estatistico: PREVISAO TRES DIAS POOL
                          Folha de Rosto:    previsao 3 dias + On The Book
  conta_pendente       -> Lanc. Estatistico: CONTAS PENDENTES
  situacao_uhs         -> Folha de Rosto:    foto atual (ocupadas, vagos,
                          walk-in, layover) + mensalistas e walk-ins

  Estatistico!C4       -> dia da auditoria (o resto da aba e' formula, se
                          vira sozinho)
  Folha de Rosto       -> Ocupacao / Diaria Media / Revpar / Receita
                          Hospedagem, em HOJE e ACUMULADO

Uso:
    Dois cliques neste arquivo .bat (Windows), ou:
    python "Preencher Estatistico.bat"

Depende de:  openpyxl  e  xlrd   (pip install openpyxl xlrd)
"""

import os
import re
import sys
import glob
import shutil
import datetime
import unicodedata

import openpyxl

# ======================================================================
# CONFIGURACAO
# ======================================================================

# item do RDS -> linha da aba "Lanc. Estatistico"
# a chave e' o nome do item sem acento e em maiuscula.
# se aparecer item novo no RDS o script avisa e voce acrescenta aqui.
MAPA_RECEITA = {
    "DIARIA":          17,   # Diarias
    "NO SHOW":         22,   # No Show
    "MENSALISTA":      24,   # Mensalista
    "DAY USE":         21,   # Day Use
    "EARLY CHECK IN":  19,   # Early check in
    "LATE CHECK OUT":  20,   # Late Check Out
    "TAXA ISS":        25,   # Taxa de ISS 5%
    "TAXA SERVICO":    26,   # Taxa de Servicos 10%
    "TAXA DE SERVICO": 26,
    "TELEFONE":        27,   # Telefonia
    "TELEFONIA":       27,
    "TAXA TURISMO":    28,   # Taxa de Turismo
    "ALUGUEL SALAO":   29,   # Locacao de Salas
    "EQUIPAMENTOS":    30,   # Equipamentos
    "LAVANDERIA":      31,   # Lavanderia
    "CAFE DA MANHA":   32,   # Cafe da Manha
    "FRIGOBAR":        33,   # Frigobar
    "RESTAURANTE":     34,   # Restaurante
    "ROOM SERVICE":    35,   # Room Service
    "BANQUETES":       36,   # Banquete
    "INTERNET":        37,   # Internet
    "ESTACIONAMENTO":  38,   # Aluguel de vagas
    "DIVERSOS":        39,   # Outras Receitas
}

LINHA_ESTORNO_HOSPEDAGEM = 23   # "Estornos"
LINHA_ESTORNO_OUTRAS     = 40   # "Estornos - Outras Receitas"

# grupo do RDS que NAO e' receita: sao pagamentos, ficam de fora
GRUPO_IGNORADO = "RECEBIMENTOS"

# estatisticas: linha da aba -> rotulo do RDS (busca por inicio do texto)
MAPA_ESTATISTICA = {
    5:  "UH'S DO HOTEL",           # UH's no Pool
    7:  "UH'S BLOQUEADAS",         # UH's Bloqueadas
    8:  "UH'S USO DA CASA",        # UH's Uso da Casa
    9:  "UH'S ALUGADAS",           # UH's Ocupadas Diarias (e' o total)
}
# linhas que o RDS nao traz: repete o valor do dia anterior e avisa
LINHAS_MANUAIS = {6: "UH's em Manutencao"}     # so' este ainda e' digitado

LINHA_PREV = {"pool": 44, "ocupadas": 45, "diaria": 46, "hospedes": 47}

# aba "Analise detalhada": dia N mora na linha N+4. Colunas que preenchemos:
COL_ANALISE = {
    21: "zero",       # U  Permanencia Estendida  -> sempre 0
    22: "day_use",    # V  Day Use
    23: "no_show",    # W  No show (numero da esquerda)
    24: "zero",       # X  Saida Antecipada       -> sempre 0
    28: "entradas",   # AB Check-ins realizados
    29: "saidas",     # AC Check-outs realizados
    30: "zero",       # AD Ocupados Grupos        -> sempre 0
}
# AA (Reservas do dia para o dia), Y e Z ficam intocados: nao tem no RDS
LINHA_CONTAS = 96                # Lanc. Estatistico

# "Check-in / Check-out Previstos" da Folha de Rosto saem das colunas Entrada
# e Saida do On The Book. O relatorio comeca no dia SEGUINTE ao da auditoria,
# entao 1 = o dia seguinte (que e' a manha em que a folha e' lida).
# Se um dia o On The Book passar a incluir o proprio dia da auditoria, troque
# para 0 aqui.
DIAS_A_FRENTE_PREVISTOS = 1

# True  = o arquivo passa a se chamar ..._06-09-2026.xlsx (a data da auditoria)
#         e o arquivo do dia anterior sai da pasta (a copia fica em backup/).
# False = o nome nunca muda, grava sempre por cima do mesmo arquivo.
RENOMEAR_COM_DATA = True

# Situacao das UHs: o que NAO e' pool do hotel (pool paralelo = condominio)
TIPOS_FORA_DO_POOL = ("COND",)
# Stat. Gov. que conta como "vago limpo" quando a UH esta VAGA
GOV_LIMPO = ("LIMPO", "INSPECAO", "ARRUMACAO")
GOV_SUJO = ("SUJO",)
COL_BLOCO_CONTAS = 37            # AK = dia 1; cada dia ocupa 3 colunas

# trecho do nome de arquivo que identifica cada planilha na pasta.
# use um trecho que so' apareca no nome daquela planilha.
NOME_ESTATISTICO = "ESTATISTICO_HOTEL"
NOME_FOLHA_ROSTO = "FOLHA_DE_ROSTO"


# ======================================================================
# UTILIDADES
# ======================================================================

def limpar(txt):
    if txt is None:
        return ""
    t = unicodedata.normalize("NFKD", str(txt))
    t = "".join(ch for ch in t if not unicodedata.combining(ch)).upper()
    t = t.replace("`", "'").replace("\u00b4", "'")
    t = re.sub(r"[^A-Z0-9'& ]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def brl(v):
    return ("R$ %s" % f"{v:,.2f}").replace(",", "@").replace(".", ",").replace("@", ".")


# ======================================================================
# LEITURA DOS RELATORIOS  ->  grade { (linha, coluna): valor }
# ======================================================================

def ler_grade(caminho):
    """Le .xls (xlrd) ou .xlsx (openpyxl) e devolve uma grade 1-indexada."""
    if caminho.lower().endswith(".xlsx"):
        wb = openpyxl.load_workbook(caminho, data_only=True)
        ws = wb.worksheets[0]
        return {(c.row, c.column): c.value
                for row in ws.iter_rows() for c in row if c.value not in (None, "")}
    try:
        import xlrd
    except ImportError:
        raise SystemExit(
            "Falta o xlrd pra ler arquivos .xls antigos.\n"
            "Abra o Prompt de Comando e rode:  python -m pip install xlrd"
        )
    livro = xlrd.open_workbook(caminho)
    aba = livro.sheet_by_index(0)
    grade = {}
    for r in range(aba.nrows):
        for c in range(aba.ncols):
            cel = aba.cell(r, c)
            if cel.ctype == xlrd.XL_CELL_EMPTY:
                continue
            v = cel.value
            if cel.ctype == xlrd.XL_CELL_DATE:
                v = datetime.datetime(*xlrd.xldate_as_tuple(v, livro.datemode))
            if v == "":
                continue
            grade[(r + 1, c + 1)] = v
    return grade


def linha_de(grade, r):
    return {c: v for (rr, c), v in grade.items() if rr == r}


def achar_linha(grade, texto, col=None):
    """Primeira linha cujo conteudo comeca com `texto` (ja normalizado)."""
    alvo = limpar(texto)
    for (r, c), v in sorted(grade.items()):
        if col is not None and c != col:
            continue
        if limpar(v).startswith(alvo):
            return r, c
    return None, None


def valor_a_direita(grade, r, c0, quantos=1):
    """Pega os primeiros `quantos` numeros a direita da coluna c0 na linha r."""
    achados = []
    for c in sorted(k for (rr, k) in grade if rr == r and k > c0):
        v = num(grade[(r, c)])
        if v is not None:
            achados.append(v)
            if len(achados) == quantos:
                break
    return achados


# ---------------------------------------------------------------- RDS
def ler_rds(caminho):
    g = ler_grade(caminho)
    out = {"receitas": {}, "estorno_hospedagem": 0.0, "estorno_outras": 0.0,
           "itens_nao_mapeados": [], "estatisticas": {}}

    # --- blocos de receita -------------------------------------------
    grupo, col_bruto, col_estorno = None, None, None
    for r in sorted({rr for rr, _ in g}):
        linha = linha_de(g, r)
        a = limpar(linha.get(1))

        if a.startswith("GRUPO"):
            grupo = limpar(linha.get(2))
            continue
        if a == "ITEM":                       # cabecalho: descobre as colunas
            for c, v in linha.items():
                t = limpar(v)
                if "BRUTO" in t:
                    col_bruto = c
                elif t == "ESTORNOS":
                    col_estorno = c
            continue
        if not grupo or grupo == GRUPO_IGNORADO or not a:
            continue
        if a.startswith("SUBTOTAL") or a.startswith("TOTAL"):
            if a.startswith("TOTAL"):
                grupo = None
            continue

        bruto = num(linha.get(col_bruto))
        est = num(linha.get(col_estorno)) or 0.0
        if bruto is None:
            continue

        if grupo.startswith("HOSPEDAGEM"):
            out["estorno_hospedagem"] += est
        else:
            out["estorno_outras"] += est

        chave = a
        if chave in MAPA_RECEITA:
            out["receitas"][MAPA_RECEITA[chave]] = \
                out["receitas"].get(MAPA_RECEITA[chave], 0.0) + bruto
        elif bruto or est:
            out["itens_nao_mapeados"].append((grupo, a, bruto))

    # --- estatisticas -------------------------------------------------
    for linha_planilha, rotulo in MAPA_ESTATISTICA.items():
        r, c = achar_linha(g, rotulo, col=1)
        if r:
            v = valor_a_direita(g, r, c)
            out["estatisticas"][linha_planilha] = v[0] if v else 0.0

    # cortesia + permuta somam na mesma linha da planilha
    total_cp = 0.0
    for rot in ("UH'S CORTESIA", "UH'S PERMUTA"):
        r, c = achar_linha(g, rot, col=1)
        if r:
            v = valor_a_direita(g, r, c)
            total_cp += v[0] if v else 0.0
    out["estatisticas"][11] = total_cp

    # adultos / criancas vem juntos: "122/2/0"
    r, c = achar_linha(g, "HOSPEDES ADULTOS", col=1)
    if r:
        for cc in sorted(k for (rr, k) in g if rr == r and k > c):
            partes = str(g[(r, cc)]).split("/")
            if len(partes) >= 2 and partes[0].strip().isdigit():
                out["estatisticas"][13] = float(partes[0])
                out["estatisticas"][14] = sum(float(p) for p in partes[1:]
                                              if p.strip().isdigit())
                break

    # --- campos avulsos usados na Analise detalhada e na Folha de Rosto ---
    def simples(rotulo):
        r, c = achar_linha(g, rotulo, col=1)
        if not r:
            return 0.0
        v = valor_a_direita(g, r, c)
        return v[0] if v else 0.0

    def antes_da_barra(rotulo):
        """'0/1' -> 0. O RDS junta No Show e No Show Cobrados numa celula."""
        r, c = achar_linha(g, rotulo, col=1)
        if not r:
            return 0.0
        for cc in sorted(k for (rr, k) in g if rr == r and k > c):
            m = re.match(r"^\s*(\d+)\s*/", str(g[(r, cc)]))
            if m:
                return float(m.group(1))
        return 0.0

    out["extra"] = {
        "day_use":     simples("HOSPEDES DAY USE"),
        "no_show":     antes_da_barra("UH'S NO SHOW"),
        "entradas":    simples("QUANTIDADE DE ENTRADAS"),
        "saidas":      simples("QUANTIDADE DE SAIDAS"),
        "cortesia":    simples("UH'S CORTESIA"),
        "uso_casa":    simples("UH'S USO DA CASA"),
        "zero":        0.0,
    }

    # total de debitos: e' a conferencia final
    r, c = achar_linha(g, "TOTAL DOS DEBITOS", col=1)
    v = valor_a_direita(g, r, c) if r else []
    out["total_debitos"] = v[0] if v else None

    # data do relatorio
    r, c = achar_linha(g, "PERIODO")
    out["data"] = None
    if r:
        for cc, v in sorted(linha_de(g, r).items()):
            m = re.findall(r"(\d{2}/\d{2}/\d{4})", str(v))
            if m:
                out["data"] = datetime.datetime.strptime(m[-1], "%d/%m/%Y")
                break
    return out


# ---------------------------------------------------------- ON THE BOOK
def ler_on_the_book(caminho):
    g = ler_grade(caminho)
    r_cab, _ = achar_linha(g, "DATA", col=1)
    cab = {}
    for c, v in linha_de(g, r_cab).items():
        cab[limpar(v)] = c
    # "%" some no limpar(); o rotulo "Ocupacao" mora na linha de cima
    col_ocup = None
    for c, v in linha_de(g, r_cab - 1).items():
        if limpar(v).startswith("OCUPACAO"):
            col_ocup = c
    col = {"ocup": col_ocup, "ocupadas": cab.get("OCUPADAS"),
           "adultos": cab.get("ADULTOS"), "diaria": cab.get("MEDIA"),
           "receita": cab.get("TOT REC"), "entrada": cab.get("ENTRADA"),
           "saida": cab.get("SAIDA")}

    dias, total_rec = {}, None
    for r in sorted({rr for rr, _ in g if rr > r_cab}):
        linha = linha_de(g, r)
        a = linha.get(1)
        if limpar(a).startswith("TOTALIZACAO"):
            total_rec = num(linha.get(col["receita"]))
            break
        if not isinstance(a, datetime.datetime):
            continue
        dias[a.date()] = {
            "ocupacao": (num(linha.get(col["ocup"])) or 0) / 100.0,
            "ocupadas": num(linha.get(col["ocupadas"])) or 0,
            "adultos":  num(linha.get(col["adultos"])) or 0,
            "diaria":   num(linha.get(col["diaria"])) or 0,
            "entrada":  num(linha.get(col["entrada"])) or 0,
            "saida":    num(linha.get(col["saida"])) or 0,
        }
    return {"dias": dias, "total_receita": total_rec}


# ------------------------------------------------------- CONTA PENDENTE
def ler_conta_pendente(caminho):
    g = ler_grade(caminho)
    r_tot, c_tot = achar_linha(g, "TOTAL GERAL")

    # empilha as linhas de cabecalho pra achar "Saldo" e "Creditos gerais"
    r_cab, _ = achar_linha(g, "UH", col=1)
    pilha = {}
    for (r, c), v in g.items():
        if r <= r_cab and isinstance(v, str):
            pilha.setdefault(c, []).append((r, v))
    col_saldo = col_credito = None
    for c, itens in pilha.items():
        txt = limpar(" ".join(t for _, t in sorted(itens)))
        if txt == "SALDO":
            col_saldo = c
        elif "GERAIS" in txt and "CREDITO" in txt:
            col_credito = c

    linha_tot = linha_de(g, r_tot)
    debito = num(linha_tot.get(col_saldo)) or 0.0
    credito = abs(num(linha_tot.get(col_credito)) or 0.0)

    # conta as linhas de conta: tem UH na coluna A e status na E
    qtd = 0
    for r in sorted({rr for rr, _ in g if r_cab < rr < r_tot}):
        linha = linha_de(g, r)
        if linha.get(1) and num(linha.get(col_saldo)) is not None:
            if limpar(linha.get(1)) not in ("UH", ""):
                qtd += 1
    return {"qtd": qtd, "debito": debito, "credito": credito}


# ------------------------------------------------------ SITUACAO DAS UHS
def ler_situacao_uhs(caminho, data_auditoria):
    """
    Foto do hotel na hora em que o relatorio saiu. Conta so' o pool
    (tipo != COND). Cada UH conta UMA vez mesmo com varios hospedes.

    Walk-in e layover: chegada == data da auditoria, identificados por
    "WALK IN" / "LAY OVER" / "LAYOVER" na Empresa ou na linha de Obs.
    Mensalista: "Tipo Hosp." = Mensalista.
    """
    g = ler_grade(caminho)
    r_cab, _ = achar_linha(g, "UH", col=1)
    cab = {limpar(v): c for (r, c), v in g.items() if r == r_cab}
    col = {"sit": cab.get("SITUACAO"), "tipo": cab.get("TIPO"),
           "gov": cab.get("STAT GOV"), "thosp": cab.get("TIPO HOSP"),
           "cheg": cab.get("CHEGADA"), "emp": cab.get("EMPRESA")}

    linhas = sorted({rr for rr, _ in g if rr > r_cab})
    # "Situacao" no cabecalho fica uma coluna a direita do dado (Bloco vem vazio).
    # Confere onde os valores realmente estao antes de usar.
    for tentativa in (col["sit"], (col["sit"] or 0) - 1):
        if tentativa and any(limpar(g.get((r, tentativa))) in ("OCUPADO", "VAGO", "BLOQUEADO")
                             for r in linhas[:60]):
            col["sit"] = tentativa
            break

    uhs, atual = {}, None
    for r in linhas:
        a = g.get((r, 1))
        txt = limpar(a)
        if txt.startswith("OBS"):
            if atual and uhs[atual]["reservas"]:
                uhs[atual]["reservas"][-1]["obs"] = str(a)
            continue
        if not txt or txt == "UH" or txt.startswith("FRONTOFFICE"):
            continue
        tipo = limpar(g.get((r, col["tipo"])))
        if not tipo or tipo == "TIPO":
            continue
        uh = txt
        rec = uhs.setdefault(uh, {"tipo": tipo, "sit": limpar(g.get((r, col["sit"]))),
                                   "gov": limpar(g.get((r, col["gov"]))), "reservas": []})
        cheg = g.get((r, col["cheg"]))
        rec["reservas"].append({
            "thosp": limpar(g.get((r, col["thosp"]))),
            "cheg": cheg.date() if isinstance(cheg, datetime.datetime) else None,
            "emp": limpar(g.get((r, col["emp"]))),
            "obs": "",
        })
        atual = uh

    pool = {k: v for k, v in uhs.items() if v["tipo"] not in TIPOS_FORA_DO_POOL}
    dia = data_auditoria.date() if hasattr(data_auditoria, "date") else data_auditoria

    def marcado(rec, *palavras):
        return any(r["cheg"] == dia and any(p in (r["emp"] + " " + limpar(r["obs"]))
                                            for p in palavras)
                   for r in rec["reservas"])

    out = {"total_pool": len(pool), "ocupadas": 0, "bloqueadas": 0,
           "vago_limpo": 0, "vago_sujo": 0, "vago_outro": [], "walk_in": [],
           "layover": [], "mensalistas": [], "gov_desconhecido": set()}
    for uh, rec in pool.items():
        if rec["sit"] == "OCUPADO":
            out["ocupadas"] += 1
            if any("MENSALISTA" in r["thosp"] for r in rec["reservas"]):
                out["mensalistas"].append(uh)
        elif rec["sit"] == "BLOQUEADO":
            out["bloqueadas"] += 1
        elif rec["sit"] == "VAGO":
            if rec["gov"] in GOV_LIMPO:
                out["vago_limpo"] += 1
            elif rec["gov"] in GOV_SUJO:
                out["vago_sujo"] += 1
            else:
                out["vago_outro"].append((uh, rec["gov"]))
                out["gov_desconhecido"].add(rec["gov"])
        if marcado(rec, "WALK IN", "WALKIN", "WALK-IN"):
            out["walk_in"].append(uh)
        if marcado(rec, "LAY OVER", "LAYOVER"):
            out["layover"].append(uh)
    return out


# ======================================================================
# PREENCHIMENTO
# ======================================================================

def _merge_de(ws, r, c):
    """Devolve o intervalo mesclado que contem (r, c), ou None."""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
            return rng
    return None


def _gravavel(ws, r, c):
    """Numa celula mesclada so' da' pra escrever na do canto superior esquerdo."""
    rng = _merge_de(ws, r, c)
    return ws.cell(rng.min_row, rng.min_col) if rng else ws.cell(r, c)


def celula_a_direita(ws, texto, col_rotulo=12):
    """Acha o rotulo na coluna dada e devolve a celula ao lado."""
    alvo = limpar(texto)
    for r in range(1, ws.max_row + 1):
        if limpar(ws.cell(r, col_rotulo).value) == alvo:
            return _gravavel(ws, r, col_rotulo + 1)
    return None


def celula_abaixo(ws, texto, linhas=(14, 16)):
    """Acha o rotulo no cabecalho e devolve a celula de valor logo abaixo.

    Rotulo mesclado em varias linhas (o "Check-in Previstos" ocupa E14:E16)
    tem o valor abaixo do FIM da mescla, nao abaixo da primeira linha.
    """
    alvo = limpar(texto)
    for r in linhas:
        for c in range(1, ws.max_column + 1):
            if limpar(ws.cell(r, c).value) == alvo:
                rng = _merge_de(ws, r, c)
                return _gravavel(ws, (rng.max_row if rng else r) + 1, c)
    return None


RE_DATA_NO_NOME = re.compile(r"\d{1,2}([-_.])\d{1,2}\1\d{2,4}")


def nome_com_data(caminho, data):
    """Poe a data da auditoria no nome, respeitando o padrao que ja' existe.

    'ESTATISTICO__05-09-2026.xlsx' -> 'ESTATISTICO__06-09-2026.xlsx'
    'ESTATISTICO.xlsx'             -> 'ESTATISTICO_06-09-2026.xlsx'
    """
    pasta = os.path.dirname(caminho)
    base, ext = os.path.splitext(os.path.basename(caminho))
    achado = RE_DATA_NO_NOME.search(base)
    if achado:
        sep = achado.group(1)
        ano = "%y" if len(achado.group(0).split(sep)[-1]) == 2 else "%Y"
        nova = data.strftime(f"%d{sep}%m{sep}{ano}")
        base = base[:achado.start()] + nova + base[achado.end():]
    else:
        # nome sem data ainda: separa com espaco ou underscore, o que o nome usar
        sep = " " if (" " in base and "_" not in base) else "_"
        base = base + sep + data.strftime("%d-%m-%Y")
    return os.path.join(pasta, base + ext)


def fazer_backup(caminho):
    """Copia o arquivo pra pasta backup/ antes de sobrescrever."""
    pasta = os.path.join(os.path.dirname(caminho), "backup")
    os.makedirs(pasta, exist_ok=True)
    carimbo = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    base, ext = os.path.splitext(os.path.basename(caminho))
    destino = os.path.join(pasta, f"{base}__{carimbo}{ext}")
    shutil.copy2(caminho, destino)
    return destino


def preencher(est_path, folha_path, rds, otb, cp, uhs):
    wb = openpyxl.load_workbook(est_path)
    lc, ec = wb["Lanç. Estatístico"], wb["Estatístico"]

    dia = rds["data"].day
    col = dia + 1                      # dia 1 = coluna B
    avisos = []

    # ---- 1. data ------------------------------------------------------
    ec["C4"] = dia

    ja_tinha = any(lc.cell(row=r, column=col).value not in (None, "")
                   for r in list(range(5, 15)) + list(range(17, 41)))
    if ja_tinha:
        avisos.append(f"O dia {dia} JA tinha lancamento nesta planilha. "
                      "Sobrescrevi. Se rodou com o relatorio errado, "
                      "restaura da pasta backup/.")

    # ---- 2. estatisticas ---------------------------------------------
    for linha, valor in rds["estatisticas"].items():
        lc.cell(row=linha, column=col).value = valor
    for linha, nome in LINHAS_MANUAIS.items():
        anterior = lc.cell(row=linha, column=col - 1).value if col > 2 else 0
        lc.cell(row=linha, column=col).value = anterior or 0
        avisos.append(f"{nome}: o RDS nao traz. Repeti o dia anterior ({anterior or 0}).")
    # da Situacao das UHs
    lc.cell(row=10, column=col).value = len(uhs["mensalistas"])   # Mensalistas
    lc.cell(row=12, column=col).value = len(uhs["walk_in"])       # Walk-ins

    # ---- 3. receitas --------------------------------------------------
    for linha in range(17, 41):
        lc.cell(row=linha, column=col).value = 0
    for linha, valor in rds["receitas"].items():
        lc.cell(row=linha, column=col).value = valor
    lc.cell(row=LINHA_ESTORNO_HOSPEDAGEM, column=col).value = rds["estorno_hospedagem"]
    lc.cell(row=LINHA_ESTORNO_OUTRAS, column=col).value = rds["estorno_outras"]

    # ---- 4. previsao tres dias ---------------------------------------
    pool = rds["estatisticas"].get(5, 0)
    prev_datas = []
    for n in (1, 2, 3):
        d = (rds["data"] + datetime.timedelta(days=n)).date()
        cprev = d.day + 1
        info = otb["dias"].get(d)
        prev_datas.append((d, info))
        if not info:
            avisos.append(f"{d:%d/%m} nao esta no On The Book. Deixei em branco.")
            continue
        lc.cell(row=LINHA_PREV["pool"], column=cprev).value = pool
        lc.cell(row=LINHA_PREV["ocupadas"], column=cprev).value = info["ocupadas"]
        lc.cell(row=LINHA_PREV["diaria"], column=cprev).value = round(info["diaria"], 2)
        lc.cell(row=LINHA_PREV["hospedes"], column=cprev).value = info["adultos"]

    # ---- 5. contas pendentes -----------------------------------------
    lc.cell(row=LINHA_CONTAS, column=col).value = cp["qtd"]
    base = COL_BLOCO_CONTAS + (dia - 1) * 3
    lc.cell(row=LINHA_CONTAS, column=base).value = cp["qtd"]
    lc.cell(row=LINHA_CONTAS, column=base + 1).value = cp["debito"]
    lc.cell(row=LINHA_CONTAS, column=base + 2).value = cp["credito"]

    # ---- 5b. Analise detalhada ---------------------------------------
    ad = wb["Análise detalhada"]
    linha_ad = dia + 4
    for c, chave in COL_ANALISE.items():
        ad.cell(row=linha_ad, column=c).value = rds["extra"][chave]

    # ---- 6. conferencia ----------------------------------------------
    receitas_dia = sum(lc.cell(row=r, column=col).value or 0 for r in range(17, 41))
    hosp_dia = sum(lc.cell(row=r, column=col).value or 0 for r in range(17, 25))
    dif = receitas_dia - (rds["total_debitos"] or 0)

    # ---- 7. folha de rosto -------------------------------------------
    pool_ac = sum(lc.cell(5, x).value or 0 for x in range(2, col + 1))
    ocup_ac = sum(lc.cell(9, x).value or 0 for x in range(2, col + 1))
    hosp_ac = sum(sum(lc.cell(r, x).value or 0 for r in range(17, 25))
                  for x in range(2, col + 1))
    ocup_dia = lc.cell(9, col).value or 0
    pool_dia = lc.cell(5, col).value or 0

    fr = {
        "ocupacao":  (ocup_dia / pool_dia if pool_dia else 0, ocup_ac / pool_ac if pool_ac else 0),
        "diaria":    (hosp_dia / ocup_dia if ocup_dia else 0, hosp_ac / ocup_ac if ocup_ac else 0),
        "revpar":    (hosp_dia / pool_dia if pool_dia else 0, hosp_ac / pool_ac if pool_ac else 0),
        "hospedagem": (hosp_dia, hosp_ac),
    }
    on_the_book = hosp_ac + (otb["total_receita"] or 0)

    wf = openpyxl.load_workbook(folha_path)
    wsf = wf["Resumo Estatistico"]
    wsf["F7"] = dia
    wsf["H3"] = pool_dia
    wsf["F8"], wsf["G8"] = fr["ocupacao"]
    wsf["F9"], wsf["G9"] = round(fr["diaria"][0], 2), round(fr["diaria"][1], 2)
    wsf["F10"], wsf["G10"] = round(fr["revpar"][0], 2), round(fr["revpar"][1], 2)
    wsf["F11"], wsf["G11"] = round(fr["hospedagem"][0], 2), round(fr["hospedagem"][1], 2)
    wsf["G12"] = round(on_the_book, 2)
    ocupadas = rds["estatisticas"].get(9, 0)
    bloqueadas = rds["estatisticas"].get(7, 0)
    adultos = rds["estatisticas"].get(13, 0)
    hospedes = adultos + rds["estatisticas"].get(14, 0)
    e = rds["extra"]


    # Check-in/out previstos: colunas Entrada e Saida do On The Book
    d_prev = (rds["data"] + datetime.timedelta(days=DIAS_A_FRENTE_PREVISTOS)).date()
    prev_info = otb["dias"].get(d_prev)
    if prev_info:
        checkin_prev, checkout_prev = prev_info["entrada"], prev_info["saida"]
        avisos.append(
            f"Check-in/out Previstos vieram do On The Book de {d_prev:%d/%m} "
            f"({checkin_prev:.0f} entradas / {checkout_prev:.0f} saidas). "
            "Confere se e' esse o dia que a folha de rosto espera.")
    else:
        checkin_prev = checkout_prev = 0
        avisos.append(f"{d_prev:%d/%m} nao esta no On The Book: "
                      "Check-in/out Previstos ficaram zerados.")

    # bloco verde da direita: foto mais atual, da Situacao das UHs.
    # CHECK-IN APOS AS 00h nao e' tocado: e' escrito na passagem de plantao.
    for rotulo, valor in (
            ("OCUPADAS", uhs["ocupadas"]),
            ("WALK IN", len(uhs["walk_in"])),
            ("NO SHOW", e["no_show"]),
            ("BLOQUEADA/ INTERDITADO", uhs["bloqueadas"]),
            ("VAGO LIMPO", uhs["vago_limpo"]),
            ("VAGO SUJO", uhs["vago_sujo"]),
            ("LAYOVER", len(uhs["layover"])),
            ("USO DA CASA/ CORTESIA", e["uso_casa"] + e["cortesia"])):
        cel = celula_a_direita(wsf, rotulo)
        if cel is not None:
            cel.value = valor
        else:
            avisos.append(f"Nao achei '{rotulo}' na Folha de Rosto.")

    # bloco Estatistica do centro
    for rotulo, valor in (
            ("Número de Hóspedes", hospedes),
            # so' adultos, igual ao "Indice de Frequencia" do RDS
            ("Hóspedes/UH's Ocupadas", adultos / ocupadas if ocupadas else 0),
            ("No Show", e["no_show"]),
            ("Aptos bloq manutenção", bloqueadas),
            ("UH's Ocupadas", ocupadas),
            ("Check-in Previstos", checkin_prev),
            ("Check-out Previstos", checkout_prev),
            ("Aptos Mercado Informal", 0), ("Prorrogações", 0)):
        cel = celula_abaixo(wsf, rotulo)
        if cel is not None:
            cel.value = round(valor, 2) if isinstance(valor, float) else valor

    for i, (d, info) in enumerate(prev_datas):
        wsf.cell(row=20, column=6 + i).value = datetime.datetime(d.year, d.month, d.day)
        wsf.cell(row=21, column=6 + i).value = info["ocupacao"] if info else 0

    if uhs["gov_desconhecido"]:
        avisos.append("Stat. Gov. que nao conheco em UH vaga: " +
                      ", ".join(sorted(uhs["gov_desconhecido"])) +
                      f" (UHs {', '.join(u for u, _ in uhs['vago_outro'])}). "
                      "Nao contei nem como limpo nem como sujo - me manda o nome.")
    if uhs["total_pool"] != pool:
        avisos.append(f"Pool na Situacao das UHs = {uhs['total_pool']}, "
                      f"mas o RDS diz {pool:.0f} UH's do Hotel. Confere.")

    # ---- 8. salva, com backup antes -----------------------------------
    backups = [fazer_backup(est_path), fazer_backup(folha_path)]

    # o openpyxl injeta um <workbookProtection/> vazio que nao existia no
    # original. Nao trava nada, mas evita sujeira no arquivo.
    wb.security = None
    wf.security = None

    saidas = []
    for origem, livro in ((est_path, wb), (folha_path, wf)):
        destino = nome_com_data(origem, rds["data"]) if RENOMEAR_COM_DATA else origem
        livro.save(destino)
        # o arquivo do dia anterior sai de cena; a copia dele ja' esta em backup/
        if os.path.abspath(destino) != os.path.abspath(origem):
            try:
                os.remove(origem)
            except OSError as erro:
                avisos.append(f"Salvei como {os.path.basename(destino)} mas nao "
                              f"consegui apagar o {os.path.basename(origem)}: {erro}. "
                              "Apague na mao, senao amanha vai ter dois na pasta.")
        saidas.append(destino)

    return saidas[0], saidas[1], fr, on_the_book, dif, receitas_dia, avisos, prev_datas, backups


# ======================================================================
# MAIN
# ======================================================================

def achar(pasta, *chaves):
    """Devolve (caminho, quantos_bateram). Empate: o modificado mais recente."""
    achados = []
    for f in glob.glob(os.path.join(pasta, "*.xls")) + glob.glob(os.path.join(pasta, "*.xlsx")):
        if os.path.basename(f).startswith("~$"):      # temporario do Excel aberto
            continue
        nome = limpar(os.path.basename(f)).replace(" ", "")
        if any(limpar(k).replace(" ", "") in nome for k in chaves):
            achados.append(f)
    if not achados:
        return None, 0
    achados.sort(key=lambda f: -os.path.getmtime(f))
    return achados[0], len(achados)


def main():
    pasta = os.path.dirname(os.path.abspath(sys.argv[0])) or "."

    busca = {
        "RDS": ("RDS",),
        "On The Book": ("on_the_book", "onthebook"),
        "Conta Pendente": ("conta_pendente", "contapendente"),
        "Situação UHs": ("situacao", "situação"),
        "Estatístico": (NOME_ESTATISTICO,),
        "Folha de Rosto": (NOME_FOLHA_ROSTO,),
    }
    arqs, duplicados = {}, []
    for rotulo, chaves in busca.items():
        caminho, quantos = achar(pasta, *chaves)
        arqs[rotulo] = caminho
        if quantos > 1:
            duplicados.append(rotulo)

    faltando = [k for k, v in arqs.items() if not v]
    if faltando:
        raise SystemExit("Nao achei nesta pasta: " + ", ".join(faltando))

    print("=" * 68)
    print(" PREENCHENDO ESTATISTICO + FOLHA DE ROSTO")
    print("=" * 68)
    print("Arquivos que vou usar:")
    for rotulo, caminho in arqs.items():
        marca = "  <-- TEM MAIS DE UM! confere" if rotulo in duplicados else ""
        print(f"   {rotulo:<16} {os.path.basename(caminho)}{marca}")
    if duplicados:
        print("\n   Peguei o modificado mais recente. Se nao for esse, tira o")
        print("   arquivo velho da pasta antes de rodar de novo.")
    print("-" * 68)

    rds = ler_rds(arqs["RDS"])
    otb = ler_on_the_book(arqs["On The Book"])
    cp = ler_conta_pendente(arqs["Conta Pendente"])
    uhs = ler_situacao_uhs(arqs["Situação UHs"], rds["data"])

    print(f"Data da auditoria : {rds['data']:%d/%m/%Y}  (dia {rds['data'].day})")
    print(f"Total de Debitos  : {brl(rds['total_debitos'] or 0)}")
    print("-" * 68)

    r = preencher(arqs["Estatístico"], arqs["Folha de Rosto"], rds, otb, cp, uhs)
    out_est, out_fol, fr, otb_val, dif, receitas, avisos, prev, backups = r

    print("ESTATISTICAS")
    for linha, v in sorted(rds["estatisticas"].items()):
        print(f"   linha {linha:>2}  {v:>12,.0f}")
    print("\nRECEITAS")
    for linha, v in sorted(rds["receitas"].items()):
        print(f"   linha {linha:>2}  {brl(v):>16}")
    print(f"   estornos hospedagem  {brl(rds['estorno_hospedagem']):>16}")
    print(f"   estornos outras      {brl(rds['estorno_outras']):>16}")

    print("\nPREVISAO TRES DIAS")
    for d, info in prev:
        if info:
            print(f"   {d:%d/%m}  ocupadas={info['ocupadas']:>4.0f}  "
                  f"diaria={brl(info['diaria']):>12}  adultos={info['adultos']:>4.0f}  "
                  f"ocup={info['ocupacao']:>6.2%}")

    print(f"\nCONTAS PENDENTES  qtd={cp['qtd']}  debito={brl(cp['debito'])}  "
          f"credito={brl(cp['credito'])}")

    print("\nFOLHA DE ROSTO" + " " * 14 + "HOJE" + " " * 12 + "ACUMULADO")
    for nome, chave in (("Ocupacao", "ocupacao"), ("Diaria Media", "diaria"),
                        ("Revpar", "revpar"), ("Receita Hospedagem", "hospedagem")):
        h, a = fr[chave]
        if chave == "ocupacao":
            print(f"   {nome:<20}{h:>14.2%}{a:>21.2%}")
        else:
            print(f"   {nome:<20}{brl(h):>14}{brl(a):>21}")
    print(f"   {'On The Book':<20}{'':>14}{brl(otb_val):>21}")

    e = rds["extra"]
    print("\nANALISE DETALHADA (linha do dia)")
    print(f"   Day Use={e['day_use']:.0f}  No show={e['no_show']:.0f}  "
          f"Check-ins={e['entradas']:.0f}  Check-outs={e['saidas']:.0f}")
    print("   Permanencia Estendida, Saida Antecipada e Ocupados Grupos = 0")
    print("   'Reservas do dia para o dia' NAO foi tocada (nao existe no RDS)")

    print("\nSITUACAO DAS UHS (pool = %d)" % uhs["total_pool"])
    print(f"   Ocupadas={uhs['ocupadas']}  Bloqueadas={uhs['bloqueadas']}  "
          f"Vago limpo={uhs['vago_limpo']}  Vago sujo={uhs['vago_sujo']}")
    print(f"   Walk-in={len(uhs['walk_in'])} {uhs['walk_in'] or ''}  "
          f"Layover={len(uhs['layover'])} {uhs['layover'] or ''}  "
          f"Mensalistas={len(uhs['mensalistas'])} {uhs['mensalistas'] or ''}")

    print("\nFOLHA DE ROSTO - bloco Estatistica (RDS)")
    ocup = rds["estatisticas"].get(9, 0)
    hosp_qtd = rds["estatisticas"].get(13, 0) + rds["estatisticas"].get(14, 0)
    print(f"   UH's Ocupadas={ocup:.0f}  No Show={e['no_show']:.0f}  "
          f"Bloq manutencao={rds['estatisticas'].get(7, 0):.0f}  "
          f"Uso da Casa/Cortesia={e['uso_casa'] + e['cortesia']:.0f}")
    adt = rds["estatisticas"].get(13, 0)
    print(f"   Numero de Hospedes={hosp_qtd:.0f}  "
          f"Hospedes/UH={adt / ocup if ocup else 0:.2f} (so' adultos)")

    print("\n" + "-" * 68)
    print("CONFERENCIA")
    print(f"   TOTAL DE RECEITAS do dia : {brl(receitas)}")
    print(f"   Total de Debitos (RDS)   : {brl(rds['total_debitos'] or 0)}")
    if abs(dif) < 0.01:
        print("   >>> BATEU CERTINHO <<<")
    else:
        print(f"   >>> DIFERENCA DE {brl(dif)} - CONFERE ANTES DE ENVIAR <<<")

    if rds["itens_nao_mapeados"]:
        print("\n   *** item do RDS que nao tem linha na planilha: ***")
        for grupo, item, v in rds["itens_nao_mapeados"]:
            print(f"      [{grupo}] {item} = {brl(v)}")
        print("      Me manda esse nome que eu acrescento no MAPA_RECEITA.")

    if avisos:
        print("\n   Confere na mao:")
        for a in avisos:
            print("      -", a)

    print("\nGerados:" if RENOMEAR_COM_DATA else "\nAtualizados (mesmo nome, por cima):")
    print("  ", os.path.basename(out_est))
    print("  ", os.path.basename(out_fol))
    print("Backup do estado anterior em backup/:")
    for b in backups:
        print("  ", os.path.basename(b))
    print("=" * 68)


if __name__ == "__main__":
    main()
