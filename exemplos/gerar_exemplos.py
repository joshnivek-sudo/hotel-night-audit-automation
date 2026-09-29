#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera os arquivos de EXEMPLO (100% fictícios) usados pra demonstrar e testar
as automações sem expor nenhum dado real de hotel.

Cria:
  exemplos/cartoes/
      RELACAO_DE_COMANDAS_DEMO.xlsx      relatório de comandas no formato do FrontOffice
      MODELO_CONTROLE_DE_CARTOES.xlsx    modelo de planilha de cartões (layout próprio)
  exemplos/estatistico/
      RDS_RELATORIO_DEMO.xlsx            Resumo Diário de Situação
      on_the_book_demo.xlsx              On The Book (próximos dias)
      conta_pendente_demo.xlsx           Contas pendentes
      situacao_uhs_demo.xlsx             Situação das UHs (foto atual)
      ESTATISTICO_HOTEL_DEMO.xlsx        Estatístico com o mês em andamento
      FOLHA_DE_ROSTO_DEMO.xlsx           Folha de rosto (layout próprio)

Os relatórios imitam só a ESTRUTURA que os scripts leem (onde fica cada
rótulo e cada número). Os modelos de planilha foram desenhados do zero pra
demonstração: não são cópia das planilhas usadas em nenhum hotel.

Uso:
    python exemplos/gerar_exemplos.py
"""
import os
import random
import zipfile
import datetime as dt
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PASTA = os.path.dirname(os.path.abspath(__file__))
PASTA_CARTOES = os.path.join(PASTA, "cartoes")
PASTA_EST = os.path.join(PASTA, "estatistico")

DATA_AUDITORIA = dt.datetime(2026, 9, 28)
rnd = random.Random(42)   # semente fixa: sempre gera os mesmos números

# ----------------------------------------------------------------------
# estilo
# ----------------------------------------------------------------------
AZUL = PatternFill("solid", fgColor="1F4E78")
AZUL_CLARO = PatternFill("solid", fgColor="DDEBF7")
CINZA = PatternFill("solid", fgColor="F2F2F2")
AMARELO = PatternFill("solid", fgColor="FFF2CC")
BRANCO_NEGRITO = Font(bold=True, color="FFFFFF")
NEGRITO = Font(bold=True)
TITULO = Font(bold=True, size=14, color="1F4E78")
AVISO = Font(italic=True, size=9, color="7F7F7F")
FINA = Side(style="thin", color="BFBFBF")
BORDA = Border(left=FINA, right=FINA, top=FINA, bottom=FINA)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
BRL = '"R$" #,##0.00'
PCT = "0.00%"
DATA_FMT = "dd/mm/yyyy"


def cabecalho(ws, linha, textos, col_inicial=1, fill=AZUL, fonte=BRANCO_NEGRITO):
    for i, t in enumerate(textos):
        c = ws.cell(row=linha, column=col_inicial + i, value=t)
        c.fill, c.font, c.alignment, c.border = fill, fonte, CENTRO, BORDA


def larguras(ws, mapa):
    for col, w in mapa.items():
        ws.column_dimensions[col].width = w


def pagina_paisagem(ws):
    """Imprime em paisagem, cabendo na largura de uma página."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def aviso_demo(ws, celula):
    ws[celula] = "Dados 100% fictícios, gerados por exemplos/gerar_exemplos.py"
    ws[celula].font = AVISO


def serial_excel(d):
    return (d - dt.datetime(1899, 12, 30)).days


# ======================================================================
# CARTÕES
# ======================================================================
LANCAMENTOS_CARTAO = [
    # (descrição no sistema, cód. débito, tipo de documento)
    ("Cielo - Visa Credito", "VC", "pos"),
    ("Cielo - Visa Credito", "VC", "pos"),
    ("Cielo - Visa Credito", "VC", "bee2pay"),
    ("Cielo - Visa Debito", "VD", "pos"),
    ("Cielo - Visa Debito", "VD", "pos"),
    ("Cielo - Master Credito", "MC", "pos"),
    ("Cielo - Master Credito", "MC", "b2pay"),
    ("Mastercard Parcelado", "MP", "pos"),
    ("Cielo - Master Debito", "MD", "pos"),
    ("Cielo - Elo Credito", "EC", "pos"),
    ("Cielo - Elo Credito", "EC", "link"),
    ("Cielo - Elo Debito", "ED", "pos"),
    ("American Express", "AX", "pos"),
    ("Amex", "AX", "b2b"),
    ("Pix", "PX", "pos"),
    ("Pix", "PX", "pos"),
    ("Pix", "PX", "link"),
    ("Diners", "DN", "pos"),
    ("Hipercard Credito", "HC", "pos"),   # bandeira que o modelo não tem
    ("Cielo - Visa Credito", "VC", "pos"),
    ("Cielo - Master Credito", "MC", "pos"),
    ("Cielo - Elo Credito", "EC", "bee2pay"),
]


def documento_ficticio(tipo):
    n = rnd.randint(10000, 99999)
    return {
        "pos": rnd.choice([f"{rnd.choice('38')}{n}", ""]),   # às vezes vem vazio
        "bee2pay": f"BEE2PAY {n}",
        "b2pay": "B2PAY",                                    # erro de digitação comum
        "link": f"LINK PGTO {n}",
        "b2b": f"B2B FATURADO {n}",
    }[tipo]


def salvar_xlsx_como_frontoffice(linhas, destino):
    """Grava um .xlsx mínimo com os textos numa tabela compartilhada
    (sharedStrings.xml), do jeito que o FrontOffice exporta. O openpyxl grava
    textos "inline", formato que o leitor do script de cartões não usa."""
    textos, indice = [], {}
    xml_linhas = []
    for r, linha in enumerate(linhas, start=1):
        celulas = []
        for c, v in enumerate(linha, start=1):
            if v is None:
                continue
            ref = f"{get_column_letter(c)}{r}"
            if isinstance(v, str):
                if v not in indice:
                    indice[v] = len(textos)
                    textos.append(v)
                celulas.append(f'<c r="{ref}" t="s"><v>{indice[v]}</v></c>')
            else:
                celulas.append(f'<c r="{ref}"><v>{v}</v></c>')
        xml_linhas.append(f'<row r="{r}">{"".join(celulas)}</row>')

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    arquivos = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '</Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        "xl/workbook.xml":
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{ns}" xmlns:r="{rel}"><sheets>'
            f'<sheet name="Relacao de Comandas" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            '</Relationships>',
        "xl/worksheets/sheet1.xml":
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{ns}"><sheetData>{"".join(xml_linhas)}</sheetData></worksheet>',
        "xl/sharedStrings.xml":
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<sst xmlns="{ns}" count="{len(textos)}" uniqueCount="{len(textos)}">'
            + "".join(f"<si><t>{escape(t)}</t></si>" for t in textos) + "</sst>",
    }
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for nome, conteudo in arquivos.items():
            z.writestr(nome, conteudo)


def gerar_relacao_comandas(destino):
    """Formato do export do FrontOffice: sem cabeçalho, dado começa na linha 1.
    A=UH B=Tipo C=Reserva D=Conta E=Cód F=Descrição G=Nota H=Valor(-)
    I=Documento J=Data(serial) K=Hora(fração) L=Usuário M=Designação"""
    usuarios = ["RECEPCAO1", "RECEPCAO2", "RECEPCAO3", "AUDITORIA"]
    linhas = []
    for i, (desc, cod, tipo) in enumerate(LANCAMENTOS_CARTAO, start=1):
        uh = rnd.choice(["", f"{rnd.randint(1, 4)}{rnd.randint(1, 10):02d}"]) if i % 7 == 0 \
            else f"{rnd.randint(1, 4)}{rnd.randint(1, 10):02d}"
        valor = round(rnd.uniform(120, 1400), 2)
        hora = rnd.randint(6 * 60, 23 * 60 + 50) / (24 * 60)
        linhas.append([uh or None, rnd.choice(["STD", "LUXO"]) if uh else None,
                       700000 + rnd.randint(1, 9999), 1 + rnd.randint(0, 3), cod, desc,
                       None, -valor, documento_ficticio(tipo) or None,
                       serial_excel(DATA_AUDITORIA), hora, rnd.choice(usuarios),
                       f"HOSPEDE FICTICIO {i:02d}"])
    salvar_xlsx_como_frontoffice(linhas, destino)


def gerar_modelo_cartoes(destino):
    wb = openpyxl.Workbook()
    ctrl = wb.active
    ctrl.title = "Controle"
    ctrl["B2"] = "Controle de Cartões — modelo de demonstração"
    ctrl["B2"].font = TITULO
    ctrl["B3"].number_format = DATA_FMT      # o script grava a data aqui
    ctrl["B3"].font = Font(bold=True, size=12)
    ctrl["C3"] = "← data da auditoria (preenchida pelo script)"
    ctrl["C3"].font = AVISO

    cabecalho(ctrl, 5, ["Bandeira", "Total relatório", "POS 1", "Cancel. POS 1",
                        "POS 2", "Cancel. POS 2", "POS 3", "Cancel. POS 3",
                        "Online (BEE2PAY / B2B / LINK)", "Cancel. online",
                        "Total fechamentos", "Diferença", "Status"], col_inicial=2)
    ctrl.row_dimensions[5].height = 32
    bandeiras = ["AMEX", "MASTERCARD DÉBITO", "MASTERCARD CRÉDITO", "VISA DÉBITO",
                 "VISA CRÉDITO", "ELO CRÉDITO", "ELO DÉBITO", "PIX", "DINERS"]
    primeira = 6
    for i, b in enumerate(bandeiras):
        r = primeira + i
        ctrl.cell(row=r, column=2, value=b).font = NEGRITO
        ctrl.cell(row=r, column=3,
                  value=f"=ROUND(SUMIFS(Comandas!$K:$K,Comandas!$H:$H,$B{r}),2)")
        for c in range(4, 10):
            ctrl.cell(row=r, column=c, value=0).fill = AMARELO     # digitado à mão
        ctrl.cell(row=r, column=10,
                  value=f'=ROUND(SUMIFS(Comandas!$K:$K,Comandas!$H:$H,$B{r},'
                        f'Comandas!$I:$I,"ONLINE"),2)')
        ctrl.cell(row=r, column=11, value=0).fill = AMARELO
        ctrl.cell(row=r, column=12, value=f"=SUM(D{r}:K{r})")
        ctrl.cell(row=r, column=13, value=f"=L{r}-C{r}")
        ctrl.cell(row=r, column=14, value=f'=IF(ROUND(M{r},2)=0,"OK","Justificar")')
        for c in range(2, 15):
            cel = ctrl.cell(row=r, column=c)
            cel.border = BORDA
            if 3 <= c <= 13:
                cel.number_format = BRL
    tot = primeira + len(bandeiras)
    ctrl.cell(row=tot, column=2, value="TOTAIS").font = NEGRITO
    for c in range(3, 14):
        L = get_column_letter(c)
        cel = ctrl.cell(row=tot, column=c, value=f"=SUM({L}{primeira}:{L}{tot - 1})")
        cel.number_format, cel.font, cel.fill, cel.border = BRL, NEGRITO, AZUL_CLARO, BORDA
    ctrl.cell(row=tot, column=2).fill = AZUL_CLARO
    ctrl.cell(row=tot + 2, column=2,
              value="Células amarelas: fechamento de cada maquininha, digitado pelo auditor. "
                    "Total do relatório e coluna online vêm da aba Comandas.").font = AVISO
    aviso_demo(ctrl, f"B{tot + 3}")
    pagina_paisagem(ctrl)
    larguras(ctrl, {"A": 2, "B": 22, "C": 15, "D": 12, "E": 12, "F": 12, "G": 12,
                    "H": 12, "I": 12, "J": 17, "K": 12, "L": 15, "M": 13, "N": 12})

    cm = wb.create_sheet("Comandas")
    cabecalho(cm, 1, ["Data", "UH", "Tipo UH", "Reserva", "Conta", "Cód. Déb.",
                      "Descrição", "BANDEIRA", "CANAL", "Valor (relatório)",
                      "Valor (+)", "Documento", "Hora", "Usuário", "Designação"])
    larguras(cm, {"A": 11, "B": 6, "C": 8, "D": 10, "E": 7, "F": 9, "G": 24, "H": 20,
                  "I": 9, "J": 14, "K": 12, "L": 18, "M": 7, "N": 12, "O": 22})
    wb.save(destino)


# ======================================================================
# ESTATÍSTICO — relatórios
# ======================================================================
POOL = 40
UHS_HOTEL = [f"{andar}{n:02d}" for andar in range(1, 5) for n in range(1, 11)]
UHS_COND = [f"5{n:02d}" for n in range(1, 9)]

# números do dia da auditoria (usados em mais de um relatório, pra tudo bater)
DIA = {
    "alugadas": 31, "bloqueadas": 1, "uso_casa": 1, "cortesia": 1, "permuta": 0,
    "adultos": 52, "criancas": (3, 1), "day_use": 2, "no_show": (1, 1),
    "entradas": 12, "saidas": 10,
}
RECEITAS_DIA = {
    # grupo -> [(item, bruto, estorno)]
    "HOSPEDAGEM": [("Diária", 8959.00, -180.00), ("Early Check In", 120.00, 0.0),
                   ("Late Check Out", 90.00, 0.0), ("Day Use", 260.00, 0.0),
                   ("No Show", 289.00, 0.0), ("Mensalista", 210.00, 0.0)],
    "TAXAS": [("Taxa ISS", 447.95, 0.0), ("Taxa Serviço", 895.90, -18.00),
              ("Taxa Turismo", 62.00, 0.0)],
    "ALIMENTOS E BEBIDAS": [("Café da Manhã", 640.00, 0.0), ("Frigobar", 184.50, -12.00),
                            ("Restaurante", 512.30, 0.0), ("Room Service", 98.00, 0.0)],
    "OUTRAS RECEITAS": [("Lavanderia", 75.00, 0.0), ("Estacionamento", 350.00, 0.0),
                        ("Internet", 0.00, 0.0), ("Diversos", 40.00, 0.0)],
    "RECEBIMENTOS": [("Dinheiro", 350.00, 0.0), ("Cartão de Crédito", 9820.40, 0.0),
                     ("Pix", 1480.00, 0.0)],   # pagamento, não é receita
}


def gerar_rds(destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RDS"
    d = DATA_AUDITORIA.strftime("%d/%m/%Y")
    ws["A1"] = "FrontOffice - Resumo Diário de Situação (DEMONSTRAÇÃO)"
    ws["A1"].font = TITULO
    ws["A2"] = f"Período: {d} a {d}"
    r = 4
    total_debitos = 0.0
    for grupo, itens in RECEITAS_DIA.items():
        ws.cell(row=r, column=1, value="Grupo:").font = NEGRITO
        ws.cell(row=r, column=2, value=grupo).font = NEGRITO
        r += 1
        cabecalho(ws, r, ["Item", "Valor Bruto", "Estornos", "Valor Líquido"])
        r += 1
        sub = 0.0
        for item, bruto, est in itens:
            ws.cell(row=r, column=1, value=item)
            for c, v in ((2, bruto), (3, est), (4, bruto + est)):
                ws.cell(row=r, column=c, value=v).number_format = BRL
            sub += bruto + est
            r += 1
        ws.cell(row=r, column=1, value=f"Total {grupo.title()}").font = NEGRITO
        ws.cell(row=r, column=4, value=round(sub, 2)).number_format = BRL
        if grupo != "RECEBIMENTOS":
            total_debitos += sub
        r += 2

    ws.cell(row=r, column=1, value="Total dos Débitos").font = NEGRITO
    ws.cell(row=r, column=2, value=round(total_debitos, 2)).number_format = BRL
    r += 2

    ws.cell(row=r, column=1, value="Estatísticas").font = NEGRITO
    r += 1
    est = [
        ("UH's do Hotel", POOL),
        ("UH's Bloqueadas", DIA["bloqueadas"]),
        ("UH's Uso da Casa", DIA["uso_casa"]),
        ("UH's Alugadas", DIA["alugadas"]),
        ("UH's Cortesia", DIA["cortesia"]),
        ("UH's Permuta", DIA["permuta"]),
        ("Hóspedes Adultos/Criança1/Criança2",
         f"{DIA['adultos']}/{DIA['criancas'][0]}/{DIA['criancas'][1]}"),
        ("Hóspedes Day Use", DIA["day_use"]),
        ("UH's No Show / No Show Cobrados", f"{DIA['no_show'][0]}/{DIA['no_show'][1]}"),
        ("Quantidade de Entradas", DIA["entradas"]),
        ("Quantidade de Saídas", DIA["saidas"]),
        ("Índice de Frequência", round(DIA["adultos"] / DIA["alugadas"], 2)),
    ]
    for rot, v in est:
        ws.cell(row=r, column=1, value=rot)
        ws.cell(row=r, column=2, value=v)
        r += 1
    aviso_demo(ws, f"A{r + 1}")
    larguras(ws, {"A": 38, "B": 16, "C": 14, "D": 16})
    wb.save(destino)
    return round(total_debitos, 2)


def gerar_on_the_book(destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "On The Book"
    ws["A1"] = "FrontOffice - On The Book (DEMONSTRAÇÃO)"
    ws["A1"].font = TITULO
    inicio = DATA_AUDITORIA + dt.timedelta(days=1)
    fim = inicio + dt.timedelta(days=13)
    ws["A2"] = f"Período: {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}"
    ws["F4"] = "Ocupação"
    ws["F4"].font = NEGRITO
    cabecalho(ws, 5, ["Data", "Dia", "Ocupadas", "Adultos", "Crianças", "%", "Média",
                      "Tot. Rec", "Entrada", "Saída"])
    dias_sem = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
    r, total = 6, 0.0
    ocupadas = DIA["alugadas"] + 1
    for n in range(14):
        d = inicio + dt.timedelta(days=n)
        entrada = rnd.randint(4, 14)
        saida = rnd.randint(4, 14)
        ocupadas = max(12, min(POOL, ocupadas + entrada - saida))
        media = round(rnd.uniform(265, 320), 2)
        rec = round(ocupadas * media, 2)
        total += rec
        valores = [d, dias_sem[d.weekday()], ocupadas, round(ocupadas * 1.65),
                   rnd.randint(0, 4), round(ocupadas / POOL * 100, 2), media, rec,
                   entrada, saida]
        for c, v in enumerate(valores, start=1):
            cel = ws.cell(row=r, column=c, value=v)
            cel.border = BORDA
        ws.cell(row=r, column=1).number_format = DATA_FMT
        ws.cell(row=r, column=7).number_format = BRL
        ws.cell(row=r, column=8).number_format = BRL
        r += 1
    ws.cell(row=r, column=1, value="Totalização").font = NEGRITO
    ws.cell(row=r, column=8, value=round(total, 2)).number_format = BRL
    aviso_demo(ws, f"A{r + 2}")
    larguras(ws, {"A": 13, "B": 6, "C": 10, "D": 9, "E": 9, "F": 8, "G": 12, "H": 14,
                  "I": 9, "J": 9})
    wb.save(destino)


NOMES = ["ANA", "BRUNO", "CARLA", "DIEGO", "ELISA", "FABIO", "GABRIELA", "HUGO",
         "ISABELA", "JOAO", "KARINA", "LUCAS", "MARINA", "NATAN", "OLIVIA", "PEDRO"]


def nome_ficticio(i):
    return f"{NOMES[i % len(NOMES)]} EXEMPLO {i:02d}"


def gerar_conta_pendente(destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contas Pendentes"
    ws["A1"] = "FrontOffice - Contas Pendentes (DEMONSTRAÇÃO)"
    ws["A1"].font = TITULO
    ws["A2"] = f"Posição em {DATA_AUDITORIA:%d/%m/%Y}"
    cabecalho(ws, 4, ["UH", "Hóspede", "Chegada", "Saída", "Status", "Débitos",
                      "Créditos Gerais", "Saldo"])
    r = 5
    tot_d = tot_c = 0.0
    for i in range(7):
        uh = rnd.choice(UHS_HOTEL)
        cheg = DATA_AUDITORIA - dt.timedelta(days=rnd.randint(1, 6))
        deb = round(rnd.uniform(300, 2500), 2)
        cred = -round(rnd.uniform(0, deb * 0.8), 2)
        valores = [uh, nome_ficticio(i + 20), cheg, cheg + dt.timedelta(days=rnd.randint(2, 7)),
                   rnd.choice(["HOSPEDADO", "CHECK-OUT"]), deb, cred, round(deb + cred, 2)]
        for c, v in enumerate(valores, start=1):
            ws.cell(row=r, column=c, value=v).border = BORDA
        for c in (3, 4):
            ws.cell(row=r, column=c).number_format = DATA_FMT
        for c in (6, 7, 8):
            ws.cell(row=r, column=c).number_format = BRL
        tot_d += deb
        tot_c += cred
        r += 1
    ws.cell(row=r, column=1, value="Total Geral").font = NEGRITO
    for c, v in ((6, tot_d), (7, tot_c), (8, tot_d + tot_c)):
        cel = ws.cell(row=r, column=c, value=round(v, 2))
        cel.number_format, cel.font = BRL, NEGRITO
    aviso_demo(ws, f"A{r + 2}")
    larguras(ws, {"A": 7, "B": 24, "C": 11, "D": 11, "E": 12, "F": 13, "G": 15, "H": 13})
    wb.save(destino)


def gerar_situacao_uhs(destino):
    """Foto de agora (madrugada seguinte à auditoria)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Situacao das UHs"
    ws["A1"] = "FrontOffice - Situação das UHs (DEMONSTRAÇÃO)"
    ws["A1"].font = TITULO
    ws["A2"] = f"Emitido em {DATA_AUDITORIA + dt.timedelta(days=1):%d/%m/%Y} 02:10"
    cabecalho(ws, 4, ["UH", "Tipo", "Situação", "Stat. Gov.", "Hóspede", "Tipo Hósp.",
                      "Chegada", "Saída", "Empresa"])

    hoje = DATA_AUDITORIA
    # plano da foto: 32 ocupadas, 1 bloqueada, 5 vago limpo, 2 vago sujo = 40
    especiais = {
        "103": ("walk_in",), "208": ("walk_in",),
        "305": ("layover",), "306": ("layover",),
        "410": ("mensalista",),
    }
    bloqueadas = ["402"]
    vagos_limpos = ["104", "110", "207", "309", "408"]
    vagos_sujos = ["201", "310"]
    gov_limpo = ["LIMPO", "INSPECAO", "ARRUMACAO"]

    r, i = 5, 0
    for uh in UHS_HOTEL + UHS_COND:
        tipo = "COND" if uh in UHS_COND else ("LUXO" if uh.endswith(("09", "10")) else "STD")
        base = [uh, tipo]
        obs = None
        if uh in bloqueadas:
            linha = base + ["BLOQUEADO", "SUJO", None, None, None, None, None]
        elif uh in vagos_limpos:
            linha = base + ["VAGO", rnd.choice(gov_limpo), None, None, None, None, None]
        elif uh in vagos_sujos:
            linha = base + ["VAGO", "SUJO", None, None, None, None, None]
        elif uh in UHS_COND and uh in ("503", "507"):
            linha = base + ["VAGO", "LIMPO", None, None, None, None, None]
        else:
            i += 1
            marca = especiais.get(uh, ("normal",))[0]
            cheg = hoje if marca in ("walk_in", "layover") else \
                hoje - dt.timedelta(days=rnd.randint(1, 5))
            if marca == "mensalista":
                cheg = hoje - dt.timedelta(days=20)
            empresa = {"walk_in": "WALK IN", "layover": "CIA AEREA EXEMPLO",
                       "mensalista": "EMPRESA EXEMPLO LTDA"}.get(
                marca, rnd.choice(["PARTICULAR", "PARTICULAR", "AGENCIA EXEMPLO",
                                   "EMPRESA EXEMPLO LTDA"]))
            if uh in UHS_COND:
                empresa = "PROPRIETARIO"
            linha = base + ["OCUPADO", "SUJO", nome_ficticio(i),
                            "MENSALISTA" if marca == "mensalista" else "NORMAL",
                            cheg, cheg + dt.timedelta(days=rnd.randint(1, 4)), empresa]
            if marca == "layover":
                obs = "Obs: LAYOVER - tripulação, saída às 11h"
        for c, v in enumerate(linha, start=1):
            if v is not None:
                ws.cell(row=r, column=c, value=v)
        for c in (7, 8):
            ws.cell(row=r, column=c).number_format = DATA_FMT
        r += 1
        if obs:
            ws.cell(row=r, column=1, value=obs).font = AVISO
            r += 1
    aviso_demo(ws, f"A{r + 1}")
    larguras(ws, {"A": 8, "B": 7, "C": 11, "D": 11, "E": 22, "F": 12, "G": 11, "H": 11,
                  "I": 22})
    wb.save(destino)


# ======================================================================
# ESTATÍSTICO — modelos de planilha (layout próprio de demonstração)
# ======================================================================
ROTULOS_ESTATISTICAS = {
    5: "UH's no pool", 6: "UH's em manutenção", 7: "UH's bloqueadas",
    8: "UH's uso da casa", 9: "UH's ocupadas", 10: "Mensalistas",
    11: "Cortesia + permuta", 12: "Walk-ins", 13: "Hóspedes adultos",
    14: "Hóspedes crianças",
}
ROTULOS_RECEITA = {
    17: "Diárias", 18: "Diárias - outras", 19: "Early check-in", 20: "Late check-out",
    21: "Day use", 22: "No show", 23: "Estornos (hospedagem)", 24: "Mensalista",
    25: "Taxa de ISS", 26: "Taxa de serviço", 27: "Telefonia", 28: "Taxa de turismo",
    29: "Locação de salas", 30: "Equipamentos", 31: "Lavanderia", 32: "Café da manhã",
    33: "Frigobar", 34: "Restaurante", 35: "Room service", 36: "Banquete",
    37: "Internet", 38: "Aluguel de vagas", 39: "Outras receitas",
    40: "Estornos (outras receitas)",
}
ROTULOS_PREVISAO = {44: "UH's no pool", 45: "UH's ocupadas", 46: "Diária média",
                    47: "Hóspedes (adultos)"}
LINHA_CONTAS = 96
COL_BLOCO_CONTAS = 37


def dia_ficticio():
    """Um dia qualquer do mês, com números coerentes entre si."""
    ocup = rnd.randint(22, 36)
    diaria = rnd.uniform(260, 315)
    hosp = round(ocup * diaria, 2)
    return {
        "est": {5: POOL, 6: rnd.choice([0, 0, 1]), 7: rnd.randint(0, 2),
                8: rnd.randint(0, 1), 9: ocup, 10: 1, 11: rnd.randint(0, 1),
                12: rnd.randint(0, 3), 13: round(ocup * rnd.uniform(1.5, 1.8)),
                14: rnd.randint(0, 4)},
        "rec": {17: hosp, 19: rnd.choice([0, 120]), 20: rnd.choice([0, 90]),
                21: rnd.choice([0, 130, 260]), 22: rnd.choice([0, 0, 289]),
                23: rnd.choice([0, 0, -150]), 24: 210, 25: round(hosp * 0.05, 2),
                26: round(hosp * 0.10, 2), 28: round(ocup * 2, 2),
                31: round(rnd.uniform(0, 150), 2), 32: round(ocup * 20, 2),
                33: round(rnd.uniform(60, 260), 2), 34: round(rnd.uniform(200, 700), 2),
                35: round(rnd.uniform(0, 200), 2), 38: rnd.choice([250, 300, 350]),
                39: rnd.choice([0, 40]), 40: rnd.choice([0, 0, -12])},
        "contas": (rnd.randint(3, 9), round(rnd.uniform(2000, 9000), 2),
                   round(rnd.uniform(500, 4000), 2)),
        "analise": {"day_use": rnd.randint(0, 2), "no_show": rnd.randint(0, 2),
                    "entradas": rnd.randint(6, 15), "saidas": rnd.randint(6, 15)},
    }


def gerar_estatistico(destino):
    dia_anterior = DATA_AUDITORIA.day - 1
    wb = openpyxl.Workbook()

    # --- aba Estatístico: C4 = dia; o resto é fórmula ------------------
    ec = wb.active
    ec.title = "Estatístico"
    ec["B2"] = "Estatístico — modelo de demonstração"
    ec["B2"].font = TITULO
    ec["B4"] = "Dia do mês"
    ec["B4"].font = NEGRITO
    ec["C4"] = dia_anterior
    ec["C4"].fill = AMARELO
    ec["D4"] = "← o script só troca este número; o resto da aba é fórmula"
    ec["D4"].font = AVISO
    cabecalho(ec, 6, ["Indicador", "Hoje", "Acumulado no mês"], col_inicial=2)
    L = "'Lanç. Estatístico'!"

    def hoje(linha):
        return f"INDEX({L}$B${linha}:$AF${linha},$C$4)"

    def acum(linha):
        return f"SUM(OFFSET({L}$B${linha},0,0,1,$C$4))"

    def receita_hosp_hoje():
        return f"SUM(INDEX({L}$B$17:$AF$24,0,$C$4))"

    def receita_hosp_acum():
        return f"SUM(OFFSET({L}$B$17,0,0,8,$C$4))"

    linhas = [
        ("UH's no pool", f"={hoje(5)}", f"={acum(5)}", "0"),
        ("UH's ocupadas", f"={hoje(9)}", f"={acum(9)}", "0"),
        ("Ocupação", "=IF(C7=0,0,C8/C7)", "=IF(D7=0,0,D8/D7)", PCT),
        ("Receita de hospedagem", f"={receita_hosp_hoje()}", f"={receita_hosp_acum()}", BRL),
        ("Diária média", "=IF(C8=0,0,C10/C8)", "=IF(D8=0,0,D10/D8)", BRL),
        ("RevPAR", "=IF(C7=0,0,C10/C7)", "=IF(D7=0,0,D10/D7)", BRL),
        ("Receita total", f"=SUM(INDEX({L}$B$17:$AF$40,0,$C$4))",
         f"=SUM(OFFSET({L}$B$17,0,0,24,$C$4))", BRL),
    ]
    for i, (rot, f_hoje, f_acum, fmt) in enumerate(linhas, start=7):
        ec.cell(row=i, column=2, value=rot).border = BORDA
        for c, f in ((3, f_hoje), (4, f_acum)):
            cel = ec.cell(row=i, column=c, value=f)
            cel.number_format, cel.border = fmt, BORDA
    aviso_demo(ec, "B15")
    larguras(ec, {"A": 2, "B": 26, "C": 16, "D": 18})

    # --- aba Lanç. Estatístico ------------------------------------------
    lc = wb.create_sheet("Lanç. Estatístico")
    lc["A1"] = "Lançamentos do mês — dia N fica na coluna N+1 (B = dia 1)"
    lc["A1"].font = TITULO
    cabecalho(lc, 3, ["Indicador"] + list(range(1, 32)))
    lc.freeze_panes = "B4"

    def secao(linha, texto):
        cel = lc.cell(row=linha, column=1, value=texto)
        cel.font, cel.fill = NEGRITO, AZUL_CLARO

    secao(4, "ESTATÍSTICAS")
    secao(16, "RECEITAS")
    secao(43, "PREVISÃO TRÊS DIAS")
    secao(94, "CONTAS PENDENTES")
    for linha, rot in {**ROTULOS_ESTATISTICAS, **ROTULOS_RECEITA,
                       **ROTULOS_PREVISAO}.items():
        lc.cell(row=linha, column=1, value=rot).border = BORDA
    lc.cell(row=41, column=1, value="Total de receitas").font = NEGRITO
    lc.cell(row=LINHA_CONTAS, column=1, value="Contas pendentes (qtd)").border = BORDA
    for dia in range(1, 32):
        col = dia + 1
        L2 = get_column_letter(col)
        tot = lc.cell(row=41, column=col, value=f"=SUM({L2}17:{L2}40)")
        tot.number_format, tot.font = BRL, NEGRITO
        for linha in list(ROTULOS_RECEITA) + [46]:
            lc.cell(row=linha, column=col).number_format = BRL
        base = COL_BLOCO_CONTAS + (dia - 1) * 3
        for k, t in enumerate(("Qtd", "Débito", "Crédito")):
            h = lc.cell(row=95, column=base + k, value=f"Dia {dia} - {t}")
            h.font, h.fill, h.alignment = NEGRITO, CINZA, CENTRO
            if k:
                lc.cell(row=LINHA_CONTAS, column=base + k).number_format = BRL

    # dias anteriores já lançados (mês em andamento)
    analise_ant = {}
    for dia in range(1, dia_anterior + 1):
        col = dia + 1
        f = dia_ficticio()
        for linha, v in {**f["est"], **f["rec"]}.items():
            lc.cell(row=linha, column=col, value=v)
        for linha in ROTULOS_RECEITA:
            if lc.cell(row=linha, column=col).value is None:
                lc.cell(row=linha, column=col, value=0)
        q, d_, c_ = f["contas"]
        lc.cell(row=LINHA_CONTAS, column=col, value=q)
        base = COL_BLOCO_CONTAS + (dia - 1) * 3
        for k, v in enumerate((q, d_, c_)):
            lc.cell(row=LINHA_CONTAS, column=base + k, value=v)
        analise_ant[dia] = f["analise"]
    larguras(lc, {"A": 28})
    for col in range(2, 33):
        lc.column_dimensions[get_column_letter(col)].width = 11

    # --- aba Análise detalhada ------------------------------------------
    ad = wb.create_sheet("Análise detalhada")
    ad["A1"] = "Análise detalhada — dia N fica na linha N+4"
    ad["A1"].font = TITULO
    cab = {1: "Dia", 2: "Data", 3: "UH's ocupadas", 4: "Ocupação",
           21: "Permanência estendida", 22: "Day use", 23: "No show",
           24: "Saída antecipada", 25: "(manual)", 26: "(manual)",
           27: "Reservas do dia para o dia", 28: "Check-ins realizados",
           29: "Check-outs realizados", 30: "Ocupados grupos"}
    for c, t in cab.items():
        cel = ad.cell(row=4, column=c, value=t)
        cel.fill, cel.font, cel.alignment, cel.border = AZUL, BRANCO_NEGRITO, CENTRO, BORDA
    ad.row_dimensions[4].height = 45
    ad.column_dimensions.group("E", "T", hidden=True)
    for dia in range(1, 32):
        r = dia + 4
        colL = get_column_letter(dia + 1)
        ad.cell(row=r, column=1, value=dia)
        ad.cell(row=r, column=2,
                value=f"=DATE({DATA_AUDITORIA.year},{DATA_AUDITORIA.month},A{r})"
                ).number_format = DATA_FMT
        ad.cell(row=r, column=3, value=f"='Lanç. Estatístico'!{colL}9")
        ad.cell(row=r, column=4,
                value=f"=IF('Lanç. Estatístico'!{colL}5=0,0,"
                      f"'Lanç. Estatístico'!{colL}9/'Lanç. Estatístico'!{colL}5)"
                ).number_format = PCT
        if dia in analise_ant:
            a = analise_ant[dia]
            for c, v in ((21, 0), (22, a["day_use"]), (23, a["no_show"]), (24, 0),
                         (28, a["entradas"]), (29, a["saidas"]), (30, 0)):
                ad.cell(row=r, column=c, value=v)
    larguras(ad, {"A": 6, "B": 11, "C": 10, "D": 10})
    wb.save(destino)


def gerar_folha_de_rosto(destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumo Estatistico"
    ws["B2"] = "RESUMO ESTATÍSTICO — modelo de demonstração"
    ws["B2"].font = TITULO
    ws["G3"] = "UH's no pool"
    ws["G3"].font = NEGRITO

    # bloco principal: HOJE x ACUMULADO
    cabecalho(ws, 6, ["", "HOJE", "ACUMULADO"], col_inicial=5)
    blocos = [(7, "Dia", "0"), (8, "Ocupação", PCT), (9, "Diária média", BRL),
              (10, "RevPAR", BRL), (11, "Receita hospedagem", BRL),
              (12, "On the Book (mês)", BRL)]
    for r, rot, fmt in blocos:
        ws.cell(row=r, column=5, value=rot).font = NEGRITO
        for c in (5, 6, 7):
            cel = ws.cell(row=r, column=c)
            cel.border = BORDA
            if c > 5:
                cel.number_format = fmt

    # bloco da direita: foto atual (rótulo na coluna L, valor na M)
    cabecalho(ws, 4, ["SITUAÇÃO ATUAL", ""], col_inicial=12)
    direita = ["Ocupadas", "Walk in", "No show", "Bloqueada / Interditado", "Vago limpo",
               "Vago sujo", "Layover", "Uso da casa / Cortesia", "Check-in após 00h"]
    for i, rot in enumerate(direita, start=5):
        ws.cell(row=i, column=12, value=rot).border = BORDA
        ws.cell(row=i, column=13).border = BORDA
    ws.cell(row=13, column=13).fill = AMARELO        # preenchido na passagem de plantão

    # bloco central: rótulo na linha 14, valor logo abaixo
    ws.cell(row=13, column=2, value="ESTATÍSTICA").font = NEGRITO
    centro = ["Número de Hóspedes", "Hóspedes/UH's Ocupadas", "No Show",
              "Check-in Previstos", "Check-out Previstos", "UH's Ocupadas",
              "Aptos bloq manutenção", "Aptos Mercado Informal", "Prorrogações"]
    cabecalho(ws, 14, centro, col_inicial=2, fill=AZUL_CLARO, fonte=NEGRITO)
    ws.row_dimensions[14].height = 48
    for c in range(2, 2 + len(centro)):
        ws.cell(row=15, column=c).border = BORDA
        ws.cell(row=15, column=c).alignment = CENTRO

    # previsão três dias
    ws["E19"] = "PREVISÃO 3 DIAS"
    ws["E19"].font = NEGRITO
    ws["E20"] = "Data"
    ws["E21"] = "Ocupação prevista"
    for c in (6, 7, 8):
        ws.cell(row=20, column=c).number_format = "dd/mm"
        ws.cell(row=21, column=c).number_format = PCT
        for r in (20, 21):
            ws.cell(row=r, column=c).border = BORDA
    aviso_demo(ws, "B24")
    pagina_paisagem(ws)
    larguras(ws, {"A": 2, "B": 13, "C": 13, "D": 10, "E": 20, "F": 13, "G": 14,
                  "H": 13, "I": 12, "J": 15, "K": 3, "L": 24, "M": 10})
    wb.save(destino)


# ======================================================================
def main():
    os.makedirs(PASTA_CARTOES, exist_ok=True)
    os.makedirs(PASTA_EST, exist_ok=True)

    gerar_relacao_comandas(os.path.join(PASTA_CARTOES, "RELACAO_DE_COMANDAS_DEMO.xlsx"))
    gerar_modelo_cartoes(os.path.join(PASTA_CARTOES, "MODELO_CONTROLE_DE_CARTOES.xlsx"))

    total = gerar_rds(os.path.join(PASTA_EST, "RDS_RELATORIO_DEMO.xlsx"))
    gerar_on_the_book(os.path.join(PASTA_EST, "on_the_book_demo.xlsx"))
    gerar_conta_pendente(os.path.join(PASTA_EST, "conta_pendente_demo.xlsx"))
    gerar_situacao_uhs(os.path.join(PASTA_EST, "situacao_uhs_demo.xlsx"))
    gerar_estatistico(os.path.join(PASTA_EST, "ESTATISTICO_HOTEL_DEMO.xlsx"))
    gerar_folha_de_rosto(os.path.join(PASTA_EST, "FOLHA_DE_ROSTO_DEMO.xlsx"))

    print("Exemplos gerados em:")
    print("  ", PASTA_CARTOES)
    print("  ", PASTA_EST)
    print(f"Total dos débitos do RDS de exemplo: R$ {total:,.2f}")


if __name__ == "__main__":
    main()
