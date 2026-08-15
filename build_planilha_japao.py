# -*- coding: utf-8 -*-
"""
Gera a planilha de controle de gastos da viagem ao Japao (30/04/2027 a 16/05/2027).

Principio: a planilha so contem numeros que o viajante informou ou que vieram de
fonte publica citada. Nada de orcamento "chutado" - a coluna Orcamento nasce vazia
e a aba Roteiro traz faixas de custo de referencia, com fonte, para o viajante
preencher por conta propria.

Abas:
  Painel   -> resumo, cotacoes, orcamento por categoria (a preencher)
  Gastos   -> lancamento dos gastos
  Roteiro  -> roteiro sugerido, o que ver, o que reservar, dicas e fontes
  Por Dia  -> gasto de cada dia da viagem
  Listas   -> listas que alimentam os menus suspensos
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from datetime import date

OUT = "Controle_de_Gastos_Japao.xlsx"

# ---------------------------------------------------------------- paleta
INK      = "1C2541"
ACCENT   = "C1121F"
TEAL     = "1B7F5A"
AMBER    = "B26B00"
BLUE     = "2A5DB0"
GREY     = "6B7280"
CARD_BG  = "F3F5F9"
BAND     = "FAFBFD"
LINE     = "DCE0E8"
INPUT_BG = "FFF6D6"
INPUT_FG = "0000FF"
CALC_BG  = "EEF1F6"
OK_BG    = "DDF0E5"
WARN_BG  = "FBEBD2"
INFO_BG  = "E1E9F7"
MUTE_BG  = "ECEDF0"
ALERT_BG = "FBE3E4"

FONT = "Arial"

F_BRL  = 'R$ #,##0.00;-R$ #,##0.00;"-"'
F_NUM  = '#,##0.00;-#,##0.00;"-"'
F_INT  = '#,##0;-#,##0;"-"'
F_PCT  = '0.0%'
F_DATE = 'DD/MM/YYYY'
F_RATE = '#,##0.0000'

thin = Side(style="thin", color=LINE)


def f(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color,
                italic=italic, underline=underline)


def fill(hexcolor):
    return PatternFill("solid", fgColor=hexcolor)


def center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


def left(wrap=False, indent=0, top=False):
    return Alignment(horizontal="left", vertical="top" if top else "center",
                     wrap_text=wrap, indent=indent)


def right():
    return Alignment(horizontal="right", vertical="center")


def outline(ws, c1, r1, c2, r2, color=LINE, style="thin"):
    s = Side(style=style, color=color)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            cell.border = Border(
                left=s if c == c1 else b.left,
                right=s if c == c2 else b.right,
                top=s if r == r1 else b.top,
                bottom=s if r == r2 else b.bottom,
            )


def paint(ws, c1, r1, c2, r2, patt):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).fill = patt


def top_accent(ws, c1, r1, c2, color):
    s = Side(style="medium", color=color)
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=r1, column=c)
        b = cell.border
        cell.border = Border(left=b.left, right=b.right, top=s, bottom=b.bottom)


def banner(ws, c1, c2, titulo, subtitulo):
    """Faixa vermelha + tarja azul com titulo e subtitulo. Ocupa as linhas 1 a 5."""
    ws.row_dimensions[1].height = 6
    ws.row_dimensions[2].height = 7
    paint(ws, c1, 2, c2, 2, fill(ACCENT))
    ws.row_dimensions[3].height = 36
    ws.row_dimensions[4].height = 20
    ws.merge_cells(start_row=3, start_column=c1, end_row=3, end_column=c2)
    ws.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c2)
    paint(ws, c1, 3, c2, 4, fill(INK))
    t = ws.cell(row=3, column=c1, value=titulo)
    t.font = f(18, True, "FFFFFF"); t.alignment = left(indent=1)
    s = ws.cell(row=4, column=c1, value=subtitulo)
    s.font = f(10, False, "C9CEDD"); s.alignment = left(indent=1)
    ws.row_dimensions[5].height = 12


def section(ws, r, c1, c2, texto, cor=INK, altura=20):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(row=r, column=c1, value=texto)
    cell.font = f(10, True, "FFFFFF"); cell.alignment = left(indent=1)
    paint(ws, c1, r, c2, r, fill(cor))
    ws.row_dimensions[r].height = altura


# ================================================================== dados
# --- o unico gasto ja realizado, informado pelo viajante
PASSAGENS_BRL = 8218.93

IDA, VOLTA = date(2027, 4, 30), date(2027, 5, 16)
VIAJANTES = 2

# --- cotacoes: apuradas em 15/08/2026, com a fonte registrada na propria planilha
COTACOES = [
    ("JPY", 0.0325, "Iene japonês  -  Wise, faixa 0,0322-0,0326 na semana de 04-14/08/2026"),
    ("BRL", 1.0000, "Real brasileiro  -  moeda base da planilha"),
    ("USD", 5.1900, "Dólar americano  -  fechamento de 13/08/2026 (R$ 5,1896)"),
]

CATEGORIAS = [
    "Passagens Aéreas",
    "Hospedagem",
    "Transporte no Japão (trens/shinkansen)",
    "Transporte local (metrô/ônibus/táxi)",
    "Alimentação",
    "Passeios e Ingressos",
    "Tokyo Disney",
    "Compras e Souvenirs",
    "Seguro Viagem",
    "Chip / Internet",
    "Documentos (passaporte etc.)",
    "Bagagem e Equipamentos",
    "Outros",
]

PAGAMENTOS = [
    "Dinheiro (iene)",
    "Dinheiro (real)",
    "Cartão de crédito",
    "Cartão de débito",
    "Cartão internacional (Wise/Nomad)",
    "IC Card (Suica/Pasmo)",
    "PIX",
    "Transferência / Boleto",
]

STATUS = ["Pago", "Pendente", "Parcelado", "Reembolsado"]
PESSOAS = ["Viajante 1", "Viajante 2", "Dividido"]
CIDADES = [
    "Tóquio", "Quioto", "Osaka", "Nara", "Kawaguchiko / Monte Fuji",
    "Tokyo Disney Resort", "Em trânsito", "Brasil (pré-viagem)", "Outro",
]
STATUS_RESERVA = ["A fazer", "Pesquisando", "Reservado", "Pago"]

FIRST, LAST = 7, 400          # linhas de lancamento na aba Gastos (cabecalho na 6)
CAP = 20                      # linhas reservadas por lista na aba Listas

wb = Workbook()

# =================================================================== PAINEL
pa = wb.active
pa.title = "Painel"
pa.sheet_view.showGridLines = False
pa.sheet_properties.tabColor = INK

for col, w in (("A", 2.5), ("B", 32), ("C", 15), ("D", 15), ("E", 15),
               ("F", 15), ("G", 22), ("H", 16), ("I", 16), ("J", 2.5)):
    pa.column_dimensions[col].width = w

banner(pa, 2, 9, "CONTROLE DE GASTOS  •  JAPÃO 2027",
       "30 de abril a 16 de maio de 2027  •  Tóquio, Quioto e Osaka  •  17 dias, 16 noites")

section(pa, 6, 2, 3, "DADOS DA VIAGEM")
section(pa, 6, 5, 7, "CÂMBIO — quanto vale 1 unidade em R$", ACCENT)

campos = [
    ("Destino", "Japão", "General"),
    ("Data de ida", IDA, F_DATE),
    ("Data de volta", VOLTA, F_DATE),
    ("Nº de viajantes", VIAJANTES, F_INT),
]
r = 7
for label, val, fmt in campos:
    pa.cell(row=r, column=2, value=label).font = f(10, False, GREY)
    pa.cell(row=r, column=2).alignment = left(indent=1)
    c = pa.cell(row=r, column=3, value=val)
    c.font = f(10, True, INPUT_FG); c.fill = fill(INPUT_BG)
    c.alignment = center(); c.number_format = fmt
    pa.row_dimensions[r].height = 19
    r += 1

for label, formula, fmt in (
    ("Duração (dias)", '=IFERROR($C$9-$C$8+1,"")', F_INT),
    ("Orçamento total (R$)", "=SUM($C$25:$C$37)", F_BRL),
):
    pa.cell(row=r, column=2, value=label).font = f(10, False, GREY)
    pa.cell(row=r, column=2).alignment = left(indent=1)
    c = pa.cell(row=r, column=3, value=formula)
    c.font = f(10, True, INK); c.fill = fill(CALC_BG)
    c.alignment = center(); c.number_format = fmt
    pa.row_dimensions[r].height = 19
    r += 1
outline(pa, 2, 7, 3, 12)

for col, txt in ((5, "Moeda"), (6, "Cotação (R$)"), (7, "Fonte da cotação")):
    cell = pa.cell(row=7, column=col, value=txt)
    cell.font = f(9, True, GREY); cell.alignment = center(); cell.fill = fill(BAND)
pa.row_dimensions[7].height = 19

for i, (cod, taxa, fonte) in enumerate(COTACOES):
    rr = 8 + i
    pa.cell(row=rr, column=5, value=cod).font = f(10, True, INK)
    pa.cell(row=rr, column=5).alignment = center()
    cc = pa.cell(row=rr, column=6, value=taxa)
    cc.font = f(10, True, INPUT_FG); cc.fill = fill(INPUT_BG)
    cc.alignment = center(); cc.number_format = F_RATE
    fc = pa.cell(row=rr, column=7, value=fonte)
    fc.font = f(8, False, GREY); fc.alignment = left(indent=1)
    pa.row_dimensions[rr].height = 19
outline(pa, 5, 7, 7, 10)

pa.merge_cells("E11:G12")
pa["E11"] = ("Cotações apuradas em 15/08/2026. Como a viagem é em 2027, trate como referência "
             "e atualize antes de comprar e durante a viagem.")
pa["E11"].font = f(9, False, ACCENT, italic=True)
pa["E11"].alignment = left(wrap=True, indent=1)
pa.row_dimensions[13].height = 12

# --- KPIs
section(pa, 14, 2, 9, "RESUMO")

GASTO = f'SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$G${FIRST}:$G${LAST},"<>Reembolsado")'
KPIS = [
    (16, 2, "TOTAL GASTO", f"={GASTO}", F_BRL, INK, CARD_BG),
    (16, 4, "JÁ PAGO",
     f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$G${FIRST}:$G${LAST},"Pago")', F_BRL, TEAL, OK_BG),
    (16, 6, "EM ABERTO",
     f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$G${FIRST}:$G${LAST},"Pendente")'
     f'+SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$G${FIRST}:$G${LAST},"Parcelado")', F_BRL, AMBER, WARN_BG),
    (16, 8, "SALDO DO ORÇAMENTO", f'=IF($C$12=0,"",$C$12-$B$17)', F_BRL, BLUE, INFO_BG),
    (19, 2, "GASTO POR PESSOA", '=IFERROR($B$17/$C$10,0)', F_BRL, INK, CARD_BG),
    (19, 4, "GASTO POR DIA DE VIAGEM", '=IFERROR($B$17/$C$11,0)', F_BRL, INK, CARD_BG),
    (19, 6, "GASTOS ANTES DE EMBARCAR",
     f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$A${FIRST}:$A${LAST},"<"&$C$8)', F_BRL, GREY, MUTE_BG),
    (19, 8, "GASTOS DURANTE A VIAGEM",
     f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$A${FIRST}:$A${LAST},">="&$C$8)', F_BRL, ACCENT, CARD_BG),
]
for rl, c1, titulo, formula, fmt, cor, bg in KPIS:
    c2 = c1 + 1
    pa.merge_cells(start_row=rl, start_column=c1, end_row=rl, end_column=c2)
    pa.merge_cells(start_row=rl + 1, start_column=c1, end_row=rl + 1, end_column=c2)
    paint(pa, c1, rl, c2, rl + 1, fill(bg))
    t = pa.cell(row=rl, column=c1, value=titulo)
    t.font = f(8, True, GREY); t.alignment = center(wrap=True)
    v = pa.cell(row=rl + 1, column=c1, value=formula)
    v.font = f(15, True, cor); v.alignment = center(); v.number_format = fmt
    outline(pa, c1, rl, c2, rl + 1)
    top_accent(pa, c1, rl, c2, cor)
    pa.row_dimensions[rl].height = 20
    pa.row_dimensions[rl + 1].height = 30
pa.row_dimensions[18].height = 8
pa.row_dimensions[21].height = 14

# --- por categoria (orcamento em branco: quem preenche e o viajante)
section(pa, 23, 2, 9, "GASTOS POR CATEGORIA  —  preencha a coluna Orçamento com os seus valores")

for i, h in enumerate(["Categoria", "Orçamento (R$)", "Gasto (R$)", "Já pago", "Em aberto",
                       "% do total", "Saldo do orçamento", "Lançamentos"]):
    cell = pa.cell(row=24, column=2 + i, value=h)
    cell.font = f(9, True, "FFFFFF"); cell.fill = fill("3B4A6B")
    cell.alignment = center(wrap=True)
pa.row_dimensions[24].height = 28

for i in range(len(CATEGORIAS)):
    rr = 25 + i
    pa.row_dimensions[rr].height = 18
    band = fill(BAND) if i % 2 == 0 else fill("FFFFFF")

    a = pa.cell(row=rr, column=2, value=f'=IF(Listas!$B${7+i}="","",Listas!$B${7+i})')
    a.font = f(10, False, INK); a.alignment = left(indent=1); a.fill = band

    b = pa.cell(row=rr, column=3)          # <- vazia de proposito
    b.font = f(10, False, INPUT_FG); b.fill = fill(INPUT_BG)
    b.alignment = right(); b.number_format = F_BRL

    for col, formula in (
        (4, f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$B${FIRST}:$B${LAST},$B{rr},'
            f'Gastos!$G${FIRST}:$G${LAST},"<>Reembolsado")'),
        (5, f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$B${FIRST}:$B${LAST},$B{rr},'
            f'Gastos!$G${FIRST}:$G${LAST},"Pago")'),
        (6, f'=$D{rr}-$E{rr}'),
        (7, f'=IFERROR($D{rr}/$D$38,0)'),
        (8, f'=IF($C{rr}="","",$C{rr}-$D{rr})'),
        (9, f'=COUNTIFS(Gastos!$B${FIRST}:$B${LAST},$B{rr})'),
    ):
        cell = pa.cell(row=rr, column=col, value=formula)
        cell.font = f(10, False, INK); cell.alignment = right(); cell.fill = band
        cell.number_format = F_PCT if col == 7 else (F_INT if col == 9 else F_BRL)

tr = 38
pa.row_dimensions[tr].height = 22
pa.cell(row=tr, column=2, value="TOTAL").font = f(10, True, "FFFFFF")
pa.cell(row=tr, column=2).alignment = left(indent=1)
for col in range(3, 10):
    L = get_column_letter(col)
    cell = pa.cell(row=tr, column=col, value=f"=SUM({L}25:{L}37)")
    cell.font = f(10, True, "FFFFFF"); cell.alignment = right()
    cell.number_format = F_PCT if col == 7 else (F_INT if col == 9 else F_BRL)
paint(pa, 2, tr, 9, tr, fill(INK))
outline(pa, 2, 24, 9, tr)

pa.conditional_formatting.add("D25:D37", DataBarRule(
    start_type="num", start_value=0, end_type="max", color=ACCENT, showValue=True))
pa.conditional_formatting.add("H25:H37", CellIsRule(
    operator="lessThan", formula=["0"], font=f(10, True, ACCENT)))
pa.row_dimensions[39].height = 12

# --- forma de pagamento | pessoa
section(pa, 40, 2, 5, "POR FORMA DE PAGAMENTO", "3B4A6B")
section(pa, 40, 7, 9, "POR PESSOA", "3B4A6B")
for i, h in enumerate(["Forma de pagamento", "Total (R$)", "Já pago", "Em aberto"]):
    cell = pa.cell(row=41, column=2 + i, value=h)
    cell.font = f(9, True, GREY); cell.fill = fill(BAND); cell.alignment = center(wrap=True)
for i, h in enumerate(["Quem pagou", "Total (R$)", "Em aberto"]):
    cell = pa.cell(row=41, column=7 + i, value=h)
    cell.font = f(9, True, GREY); cell.fill = fill(BAND); cell.alignment = center(wrap=True)
pa.row_dimensions[41].height = 24

for i in range(len(PAGAMENTOS)):
    rr = 42 + i
    pa.row_dimensions[rr].height = 18
    a = pa.cell(row=rr, column=2, value=f'=IF(Listas!$D${7+i}="","",Listas!$D${7+i})')
    a.font = f(10, False, INK); a.alignment = left(indent=1)
    for col, formula in (
        (3, f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$F${FIRST}:$F${LAST},$B{rr})'),
        (4, f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$F${FIRST}:$F${LAST},$B{rr},'
            f'Gastos!$G${FIRST}:$G${LAST},"Pago")'),
        (5, f'=$C{rr}-$D{rr}'),
    ):
        cell = pa.cell(row=rr, column=col, value=formula)
        cell.font = f(10, False, INK); cell.alignment = right(); cell.number_format = F_BRL
outline(pa, 2, 41, 5, 41 + len(PAGAMENTOS))

for i in range(len(PESSOAS)):
    rr = 42 + i
    a = pa.cell(row=rr, column=7, value=f'=IF(Listas!$H${7+i}="","",Listas!$H${7+i})')
    a.font = f(10, False, INK); a.alignment = left(indent=1)
    for col, formula in (
        (8, f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$E${FIRST}:$E${LAST},$G{rr})'),
        (9, f'=SUMIFS(Gastos!$J${FIRST}:$J${LAST},Gastos!$E${FIRST}:$E${LAST},$G{rr},'
            f'Gastos!$G${FIRST}:$G${LAST},"Pendente")'),
    ):
        cell = pa.cell(row=rr, column=col, value=formula)
        cell.font = f(10, False, INK); cell.alignment = right(); cell.number_format = F_BRL
outline(pa, 7, 41, 9, 41 + len(PESSOAS))

guia = 42 + len(PAGAMENTOS) + 1
section(pa, guia, 2, 9, "COMO USAR ESTA PLANILHA", ACCENT)
LINHAS_GUIA = [
    ("1.  Atualize as cotações (célula amarela) sempre que for comprar algo em iene.", False),
    ("2.  Defina o orçamento de cada categoria na coluna \"Orçamento (R$)\" acima. Ela nasce vazia "
     "de propósito — a aba ROTEIRO traz faixas de custo reais para você se basear.", False),
    ("3.  Lance cada gasto na aba GASTOS, uma linha por gasto. As colunas com seta abrem menu suspenso.", False),
    ("4.  Digite o valor na moeda em que pagou e escolha a moeda — a conversão para real é automática.", False),
    ("5.  Marque o STATUS (Pago / Pendente / Parcelado / Reembolsado). Este painel se atualiza sozinho.", False),
    ("ATENÇÃO: a viagem começa dentro da Golden Week japonesa (29/04 a 05/05/2027), "
     "o feriadão mais movimentado do Japão. Veja o alerta na aba ROTEIRO antes de reservar qualquer coisa.", True),
    ("Amarelo = você preenche.    Cinza = calculado automaticamente, não precisa mexer.", True),
]
for i, (t, destaque) in enumerate(LINHAS_GUIA):
    rr = guia + 1 + i
    pa.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=9)
    cell = pa.cell(row=rr, column=2, value=t)
    cell.font = f(9, destaque, ACCENT if "ATENÇÃO" in t else GREY)
    cell.alignment = left(indent=1, wrap=True)
    pa.row_dimensions[rr].height = 26 if len(t) > 120 else 17
    paint(pa, 2, rr, 9, rr, fill(ALERT_BG if "ATENÇÃO" in t else BAND))
outline(pa, 2, guia + 1, 9, guia + len(LINHAS_GUIA))
PAINEL_FIM = guia + len(LINHAS_GUIA)

# =================================================================== GASTOS
gs = wb.create_sheet("Gastos")
gs.sheet_view.showGridLines = False
gs.sheet_properties.tabColor = ACCENT

COLS = [
    ("Data", 12, F_DATE, "in"), ("Categoria", 32, "General", "in"),
    ("Descrição", 36, "General", "in"), ("Cidade / Local", 20, "General", "in"),
    ("Quem pagou", 15, "General", "in"), ("Forma de pagamento", 24, "General", "in"),
    ("Status", 13, "General", "in"), ("Moeda", 9, "General", "in"),
    ("Valor pago", 14, F_NUM, "in"), ("Valor em R$", 15, F_BRL, "calc"),
    ("Observações", 40, "General", "in"),
]
for i, (h, w, _, _) in enumerate(COLS):
    gs.column_dimensions[get_column_letter(i + 1)].width = w

banner(gs, 1, 11, "LANÇAMENTO DE GASTOS  •  JAPÃO 2027",
       "Uma linha por gasto. \"Valor em R$\" é calculado pela cotação do Painel — não digite nada nessa coluna.")

gs.row_dimensions[6].height = 26
for i, (h, _, _, kind) in enumerate(COLS):
    cell = gs.cell(row=6, column=i + 1, value=h)
    cell.font = f(9, True, "FFFFFF")
    cell.fill = fill(ACCENT if kind == "calc" else "3B4A6B")
    cell.alignment = center(wrap=True)

HEAD = 6
FIRST_G = FIRST

for rr in range(FIRST_G, LAST + 1):
    gs.row_dimensions[rr].height = 18
    band = fill("FFFFFF") if (rr - FIRST_G) % 2 == 0 else fill(BAND)
    for i, (_, _, fmt, kind) in enumerate(COLS):
        col = i + 1
        cell = gs.cell(row=rr, column=col)
        cell.font = f(10, False, INK)
        cell.number_format = fmt
        cell.fill = fill(CALC_BG) if kind == "calc" else band
        cell.alignment = right() if col in (9, 10) else (
            center() if col in (1, 7, 8) else left(indent=1))
        cell.border = Border(bottom=thin)
    gs.cell(row=rr, column=10).value = (
        f'=IF($I{rr}="","",IFERROR($I{rr}*INDEX(TAXAS,MATCH($H{rr},MOEDAS,0)),""))')

# gasto real ja informado pelo viajante
real = [None, "Passagens Aéreas", "Passagens aéreas Brasil - Tóquio (ida e volta, 2 pessoas)",
        "Brasil (pré-viagem)", "Dividido", None, "Pago", "BRL", PASSAGENS_BRL, None,
        "Valor informado por você. Preencha a data da compra e a forma de pagamento."]
for i, v in enumerate(real):
    if v is not None:
        gs.cell(row=FIRST_G, column=i + 1).value = v
for col in (1, 6):                                   # data e forma de pagamento em aberto
    gs.cell(row=FIRST_G, column=col).fill = fill(INPUT_BG)
gs.cell(row=FIRST_G, column=11).font = f(9, False, AMBER, italic=True)

# unica linha de exemplo, marcada para ser apagada
exemplo = [date(2027, 5, 1), "Alimentação", "EXEMPLO — APAGUE ESTA LINHA", "Tóquio",
           "Viajante 1", "Dinheiro (iene)", "Pago", "JPY", 3200, None,
           "Mostra a conversão automática: 3.200 ienes viram reais sozinhos."]
ex_row = FIRST_G + 1
for i, v in enumerate(exemplo):
    if v is not None:
        gs.cell(row=ex_row, column=i + 1).value = v
for col in range(1, 12):
    cell = gs.cell(row=ex_row, column=col)
    cell.fill = fill(ALERT_BG)
    cell.font = f(10 if col != 11 else 9, col == 3, ACCENT, italic=True)

gs.freeze_panes = f"A{FIRST_G}"
gs.auto_filter.ref = f"A{HEAD}:K{LAST}"

for nome, ref, msg in (
    ("CATEGORIAS", f"B{FIRST_G}:B{LAST}", "Escolha uma categoria (editável na aba Listas)."),
    ("CIDADES", f"D{FIRST_G}:D{LAST}", "Onde foi o gasto?"),
    ("PESSOAS", f"E{FIRST_G}:E{LAST}", "Quem pagou?"),
    ("PAGAMENTOS", f"F{FIRST_G}:F{LAST}", "Como foi pago?"),
    ("STATUS", f"G{FIRST_G}:G{LAST}", "Pago, Pendente, Parcelado ou Reembolsado."),
    ("MOEDAS", f"H{FIRST_G}:H{LAST}", "Moeda em que o valor foi pago."),
):
    dv = DataValidation(type="list", formula1=f"={nome}", allow_blank=True, showDropDown=False)
    dv.error = "Escolha uma das opções da lista."; dv.errorTitle = "Valor inválido"
    dv.prompt = msg; dv.promptTitle = nome.capitalize()
    gs.add_data_validation(dv); dv.add(ref)

for texto, bg, fg in (("Pago", OK_BG, TEAL), ("Pendente", WARN_BG, AMBER),
                      ("Parcelado", INFO_BG, BLUE), ("Reembolsado", MUTE_BG, GREY)):
    gs.conditional_formatting.add(f"G{FIRST_G}:G{LAST}", CellIsRule(
        operator="equal", formula=[f'"{texto}"'], fill=fill(bg), font=f(10, True, fg)))

# =================================================================== ROTEIRO
ro = wb.create_sheet("Roteiro")
ro.sheet_view.showGridLines = False
ro.sheet_properties.tabColor = TEAL

ro.column_dimensions["A"].width = 2.5
ro.column_dimensions["B"].width = 11
ro.column_dimensions["C"].width = 6
for col in "DEFGHIJK":
    ro.column_dimensions[col].width = 13
ro.column_dimensions["L"].width = 2.5
C1, C2 = 2, 11                                   # conteudo vai de B ate K

banner(ro, C1, C2, "ROTEIRO E DICAS  •  JAPÃO 2027",
       "Sugestão de 17 dias entrando e saindo por Tóquio, com Quioto, Osaka, Monte Fuji e Tokyo Disney")

r = 6


def bloco_texto(texto, bg=BAND, cor=GREY, bold=False, altura=None, tamanho=9):
    global r
    ro.merge_cells(start_row=r, start_column=C1, end_row=r, end_column=C2)
    cell = ro.cell(row=r, column=C1, value=texto)
    cell.font = f(tamanho, bold, cor)
    cell.alignment = left(indent=1, wrap=True)
    paint(ro, C1, r, C2, r, fill(bg))
    ro.row_dimensions[r].height = altura or (28 if len(texto) > 150 else 17)
    r += 1


def tabela(cabecalhos, larguras, linhas, alturas=None, wrap=True):
    """cabecalhos/larguras: listas paralelas; larguras = nº de colunas mescladas."""
    global r
    col = C1
    for h, w in zip(cabecalhos, larguras):
        ro.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + w - 1)
        cell = ro.cell(row=r, column=col, value=h)
        cell.font = f(9, True, "FFFFFF"); cell.fill = fill("3B4A6B")
        cell.alignment = center(wrap=True)
        col += w
    ro.row_dimensions[r].height = 20
    topo = r
    r += 1
    for i, linha in enumerate(linhas):
        col = C1
        band = fill(BAND) if i % 2 == 0 else fill("FFFFFF")
        for j, (valor, w) in enumerate(zip(linha, larguras)):
            ro.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + w - 1)
            cell = ro.cell(row=r, column=col, value=valor)
            cell.font = f(9, j == 0, INK)
            cell.alignment = left(indent=1, wrap=wrap, top=True)
            paint(ro, col, r, col + w - 1, r, band)
            col += w
        ro.row_dimensions[r].height = (alturas[i] if alturas else 30)
        r += 1
    outline(ro, C1, topo, C2, r - 1)
    return topo


# --- alerta Golden Week
section(ro, r, C1, C2, "ATENÇÃO — SUA VIAGEM COMEÇA DENTRO DA GOLDEN WEEK", ACCENT); r += 1
for t in [
    "A Golden Week de 2027 vai de quinta 29/04 a quarta 05/05 — quatro feriados nacionais emendados. "
    "É o feriadão em que o Japão inteiro viaja ao mesmo tempo.",
    "Pico de saída das cidades: 29 e 30/04 (vocês pousam justamente no dia 30).  "
    "Pico de volta: 04 e 05/05.  Nesses dias, evite deslocamento entre cidades.",
    "Por isso o roteiro abaixo mantém vocês parados em Tóquio até 07/05 e só então desce para Quioto: "
    "shinkansen lotado e hotel caro nessas datas, e a cidade fica mais tranquila porque muito japonês viaja para fora.",
    "Consequência prática: a hospedagem em Tóquio de 30/04 a 07/05 é a reserva mais urgente de todas. "
    "Hotel bom em Golden Week esgota com meses de antecedência.",
]:
    bloco_texto(t, ALERT_BG, ACCENT, altura=30 if len(t) > 150 else 20)
r += 1

# --- resumo do roteiro
section(ro, r, C1, C2, "RESUMO DO ROTEIRO"); r += 1
tabela(
    ["Trecho", "Datas", "Noites", "Por quê"],
    [4, 3, 1, 2],
    [
        ["Tóquio", "30/04 a 07/05", 7, "Atravessar a Golden Week sem trocar de cidade"],
        ["Monte Fuji (bate-volta)", "06/05", "—", "Feriadão acabou: estrada e ônibus liberam"],
        ["Tokyo Disney", "07/05", "—", "Sexta pós-feriado: dia mais calmo do período"],
        ["Quioto", "08/05 a 12/05", 4, "Templos, Arashiyama e bate-volta a Nara"],
        ["Osaka", "12/05 a 15/05", 3, "Comida, Dotonbori e Universal Studios (opcional)"],
        ["Tóquio (última noite)", "15/05 a 16/05", 1, "Volta de shinkansen e voo no dia 16"],
    ],
    alturas=[26] * 6)
r += 1

# --- dia a dia
section(ro, r, C1, C2, "ROTEIRO DIA A DIA  —  sugestão, ajuste como quiser"); r += 1
DIAS = [
    ("30/04", "sex", "Tóquio", "Chegada. Comprar o IC card (Welcome Suica / Tourist Pasmo) no aeroporto, check-in e um jantar leve perto do hotel.", "Pico da Golden Week. Não marque nada exigente no dia da chegada."),
    ("01/05", "sáb", "Tóquio", "Shibuya (cruzamento e Shibuya Sky), Harajuku e Takeshita, Meiji Jingu, Omotesando.", "Shibuya Sky é ingresso datado — compre online semanas antes."),
    ("02/05", "dom", "Tóquio", "Asakusa e o templo Senso-ji, rua Nakamise, Tokyo Skytree, parque e museus de Ueno.", "Chegue em Senso-ji antes das 9h para escapar do fluxo."),
    ("03/05", "seg", "Tóquio", "Shinjuku: jardim Gyoen, Omoide Yokocho, Kabukicho e o mirante grátis do Tocho.", "Feriado da Constituição. Prefira bairro a cartão-postal."),
    ("04/05", "ter", "Tóquio", "teamLab, Odaiba, ou Ginza com o mercado externo de Tsukiji pela manhã.", "Feriado. teamLab é ingresso datado, esgota."),
    ("05/05", "qua", "Tóquio", "Bairros com menos turista: Shimokitazawa, Nakameguro, Yanaka, Koenji. Akihabara à noite.", "Último feriado do bloco e pico de volta às cidades."),
    ("06/05", "qui", "Monte Fuji", "Bate-volta a Kawaguchiko: pagode Chureito, Oishi Park, lago Kawaguchi. Se o festival estiver aberto, Fuji Shibazakura.", "Ônibus do terminal Busta Shinjuku, ~2h, reserve antes. Festival costuma ir de meados de abril a fim de maio — confirme as datas de 2027."),
    ("07/05", "sex", "Tokyo Disney", "Um dia inteiro: DisneySea (único no mundo, área Fantasy Springs) ou Disneyland (o clássico).", "Ingresso entra à venda exatamente 2 meses antes, 14h no Japão = 02h de Brasília do dia 07/03/2027. Datas boas esgotam em horas."),
    ("08/05", "sáb", "Tóquio → Quioto", "Shinkansen de manhã (~2h15). À tarde, Higashiyama: Kiyomizu-dera e Gion no fim do dia.", "Reserve assento pelo smartEX. Considere mandar a mala por takkyubin."),
    ("09/05", "dom", "Quioto", "Arashiyama: bambuzal, templo Tenryu-ji e o parque dos macacos. À tarde, Kinkaku-ji (Pavilhão Dourado).", "Bambuzal antes das 8h ou depois das 17h — no meio do dia é impraticável."),
    ("10/05", "seg", "Quioto", "Fushimi Inari (os mil torii), templo Tofuku-ji e mercado Nishiki.", "Fushimi Inari é aberto 24h e de graça: vá bem cedo ou no fim da tarde."),
    ("11/05", "ter", "Nara", "Bate-volta de ~45 min: Todai-ji e o Grande Buda, parque dos cervos, Kasuga Taisha.", "Trem Kintetsu saindo de Quioto. Meio dia dá, mas o dia inteiro rende mais."),
    ("12/05", "qua", "Quioto → Osaka", "Manhã em Quioto (Ginkaku-ji e o Caminho do Filósofo). À tarde, trem para Osaka e Dotonbori à noite.", "Quioto-Osaka é ~15 min de shinkansen ou ~30 min de trem comum."),
    ("13/05", "qui", "Osaka", "Castelo de Osaka, Umeda Sky Building, Shinsekai e Tsutenkaku, mercado Kuromon.", "Alternativa: trocar o dia inteiro pelo Universal Studios Japan (ingresso antecipado)."),
    ("14/05", "sex", "Osaka", "Dia livre: compras em Shinsaibashi e Namba, ou bate-volta a Kobe ou ao castelo de Himeji.", "Bom dia para compras: tax-free acima de ¥5.000 apresentando o passaporte."),
    ("15/05", "sáb", "Osaka → Tóquio", "Shinkansen (~2h30). Últimas compras e hotel perto da estação ou do aeroporto.", "Sábado é movimentado: reserve o trem com antecedência."),
    ("16/05", "dom", "Voo de volta", "Traslado ao aeroporto: Narita ~1h a 1h30, Haneda ~40 min.", "Chegue 3h antes. Dá para devolver o IC card e resgatar o saldo."),
]
tabela(["Data", "Dia", "Base", "Programa sugerido", "Observação"],
       [1, 1, 2, 4, 2],
       [list(d) for d in DIAS],
       alturas=[44] * len(DIAS))
r += 1

# --- o que ver
section(ro, r, C1, C2, "O QUE VER EM CADA LUGAR"); r += 1
tabela(["Lugar", "Destaques"], [3, 7], [
    ["Tóquio", "Shibuya e o cruzamento • Shibuya Sky • Shinjuku e o Gyoen • Asakusa e Senso-ji • Tokyo Skytree • Ueno • "
               "Akihabara • Ginza e Tsukiji • Meiji Jingu e Harajuku • teamLab • Yanaka e Shimokitazawa para fugir da multidão"],
    ["Monte Fuji / Kawaguchiko", "Pagode Chureito (a foto clássica) • Oishi Park à beira do lago Kawaguchi • teleférico Kachi Kachi • "
                                 "Fuji Shibazakura, o tapete de flores rosa, que costuma abrir de meados de abril a fim de maio"],
    ["Tokyo Disney", "DisneySea: não existe igual no mundo, e a área Fantasy Springs é a mais nova • "
                     "Disneyland: o parque clássico, melhor com crianças • Premier Access no app para furar fila em atração específica"],
    ["Quioto", "Kiyomizu-dera • Fushimi Inari • Arashiyama e o bambuzal • Kinkaku-ji • Ginkaku-ji e o Caminho do Filósofo • "
               "Gion ao entardecer • mercado Nishiki • Ryoan-ji e seu jardim zen"],
    ["Nara (bate-volta)", "Todai-ji e o Grande Buda de bronze • parque dos cervos soltos • Kasuga Taisha e suas lanternas"],
    ["Osaka", "Dotonbori e o letreiro do Glico • Namba • Castelo de Osaka • Umeda Sky Building • Shinsekai e Tsutenkaku • "
              "mercado Kuromon • Universal Studios Japan"],
], alturas=[40, 34, 34, 40, 26, 34])
r += 1

# --- o que reservar
section(ro, r, C1, C2, "O QUE RESERVAR E QUANDO  —  em ordem de urgência", AMBER); r += 1
topo_check = tabela(["O que reservar", "Quando", "Status"], [4, 4, 2], [
    ["Hospedagem em Tóquio (30/04 a 07/05)", "Assim que possível — é Golden Week, esgota com meses de antecedência", None],
    ["Hospedagem em Quioto (08 a 12/05) e Osaka (12 a 15/05)", "4 a 6 meses antes", None],
    ["Ingresso do Tokyo Disney", "07/03/2027, às 02h de Brasília (14h no Japão) — exatamente 2 meses antes", None],
    ["Shinkansen Tóquio→Quioto e Osaka→Tóquio", "Abre 1 mês antes, pelo smartEX ou nos sites da JR", None],
    ["Ônibus Shinjuku → Kawaguchiko (Monte Fuji)", "Semanas antes; é ônibus com assento marcado", None],
    ["Shibuya Sky e teamLab", "Ingressos datados, semanas antes", None],
    ["Universal Studios Japan (se decidirem ir)", "Semanas antes, com Express Pass se quiserem furar fila", None],
    ["Seguro viagem", "Antes de embarcar", None],
    ["Passaporte eletrônico (com chip) válido", "Já — a isenção de visto exige passaporte biométrico", None],
    ["Chip / eSIM ou pocket wi-fi", "Semanas antes, ou eSIM na véspera", None],
    ["Cartão internacional e algum iene em espécie", "Semanas antes", None],
    ["Hotel da última noite perto do aeroporto", "Junto com o resto", None],
], alturas=[26] * 12)

status_col = C1 + 8
dv_res = DataValidation(type="list", formula1="=STATUS_RESERVA", allow_blank=True, showDropDown=False)
dv_res.prompt = "A fazer, Pesquisando, Reservado ou Pago."; dv_res.promptTitle = "Status da reserva"
ro.add_data_validation(dv_res)
for i in range(12):
    rr = topo_check + 1 + i
    cell = ro.cell(row=rr, column=status_col)
    cell.value = "A fazer"
    cell.fill = fill(INPUT_BG); cell.font = f(9, True, INPUT_FG); cell.alignment = center()
dv_res.add(f"{get_column_letter(status_col)}{topo_check+1}:{get_column_letter(status_col)}{topo_check+12}")
r += 1

# --- referencias de custo
section(ro, r, C1, C2, "FAIXAS DE CUSTO PARA VOCÊ SE BASEAR  —  fontes públicas consultadas em ago/2026", BLUE); r += 1
bloco_texto("Estes NÃO são o orçamento de vocês — são faixas publicadas por guias de viagem, para servir de ponto de partida "
            "ao preencher a coluna Orçamento do Painel. Valores em iene; converta pela cotação do Painel.",
            INFO_BG, BLUE, altura=28)
tabela(["Item", "Faixa de referência", "Observação"], [4, 3, 3], [
    ["Diária por pessoa, perfil intermediário", "¥20.000 a ¥35.000", "Tudo incluído: hotel, comida e transporte"],
    ["Diária por pessoa, perfil econômico", "¥10.000 a ¥15.000", "Hostel, konbini e transporte local"],
    ["Hotel business (quarto)", "¥6.000 a ¥12.000 / noite", "Padrão em Tóquio, Quioto e Osaka"],
    ["Hostel ou cápsula", "¥3.000 a ¥5.000 / noite", "Por pessoa"],
    ["Almoço (ramen, teishoku)", "¥1.000 a ¥1.500", "Por pessoa"],
    ["Jantar em izakaya", "¥2.500 a ¥4.500", "Por pessoa"],
    ["Transporte local por dia", "¥1.000 a ¥2.000", "Metrô e ônibus, por pessoa"],
    ["Shinkansen Tóquio ↔ Osaka (ida e volta)", "cerca de ¥29.000", "Bilhete avulso, por pessoa"],
    ["Tokyo Disney — 1-Day Passport", "¥7.900 a ¥10.900", "Preço varia conforme a lotação do dia"],
    ["Tokyo Disney — Premier Access", "¥1.500 a ¥2.000 por atração", "Opcional, comprado no app dentro do parque"],
    ["Entrada do Fuji Shibazakura", "¥1.000 a ¥1.300", "Mais caro na janela de pico do feriadão"],
    ["Ônibus + entrada do Shibazakura (combo)", "cerca de ¥3.000", "Shibazakura Liner, saindo de Kawaguchiko"],
    ["JR Pass 7 dias / 14 dias", "¥50.000 / ¥80.000", "Veja a dica abaixo: para este roteiro NÃO compensa"],
], alturas=[22] * 13)
r += 1

# --- dicas
section(ro, r, C1, C2, "DICAS PRÁTICAS", TEAL); r += 1
DICAS = [
    ("VISTO", "Brasileiro está isento de visto de turismo para o Japão até 29/09/2029, para estadas de até 90 dias — "
              "mas exige passaporte eletrônico, com chip. A isenção não garante entrada: a imigração pode pedir "
              "comprovante de hospedagem, passagem de volta e recursos financeiros."),
    ("JR PASS", "Para este roteiro, não compensa. O passe de 7 dias custa ¥50.000 e o trecho Tóquio→Quioto→Osaka→Tóquio "
                "avulso sai por volta de ¥29.000. Comprem bilhetes avulsos e, se quiserem, um passe regional do Kansai "
                "(a partir de ¥2.400) para os dias em Quioto e Osaka."),
    ("IC CARD", "Peguem o Welcome Suica ou o Tourist Pasmo no balcão do aeroporto: sem depósito, valem 28 dias. "
                "Carreguem ¥2.000 a ¥3.000 em dinheiro e usem em trem, metrô, ônibus, armário e loja de conveniência."),
    ("SHINKANSEN", "O IC card não cobre o shinkansen. Reservem pelo smartEX, que abre 1 mês antes, e vinculem o "
                   "Suica para passar direto na catraca."),
    ("DINHEIRO", "O Japão mistura cartão e dinheiro vivo. Levem iene em espécie. Caixas eletrônicos do 7-Eleven e do "
                 "Lawson funcionam 24h e aceitam cartão estrangeiro."),
    ("TAX-FREE", "Compras acima de ¥5.000 em loja credenciada saem sem imposto, apresentando o passaporte na hora do pagamento."),
    ("DISNEY", "As bilheterias dos parques estão fechadas: só entra com e-ticket comprado antes, no site ou no app oficial. "
               "Datas boas esgotam em horas — coloque o despertador para 07/03/2027, 02h de Brasília."),
    ("MONTE FUJI", "A montanha aparece ou não conforme o tempo. A melhor visibilidade é de novembro a fevereiro; em maio "
                   "é loteria. Compensa deixar o dia flexível e escolher uma manhã de céu limpo."),
    ("HORÁRIOS", "Templos de Quioto lotam entre 10h e 16h. Amanhecer e fim de tarde mudam completamente a experiência."),
    ("JESTA", "O Japão estuda exigir uma autorização eletrônica prévia (JESTA) para países isentos de visto. "
              "Ainda não vale, mas vale checar no site da embaixada uns 3 meses antes de embarcar."),
]
for idx, (tag, texto) in enumerate(DICAS):
    ro.merge_cells(start_row=r, start_column=C1, end_row=r, end_column=C1 + 1)
    ro.merge_cells(start_row=r, start_column=C1 + 2, end_row=r, end_column=C2)
    t = ro.cell(row=r, column=C1, value=tag)
    t.font = f(9, True, TEAL); t.alignment = left(indent=1, top=True)
    d = ro.cell(row=r, column=C1 + 2, value=texto)
    d.font = f(9, False, INK); d.alignment = left(indent=1, wrap=True, top=True)
    paint(ro, C1, r, C2, r, fill(BAND if idx % 2 == 0 else "FFFFFF"))
    ro.row_dimensions[r].height = 40 if len(texto) > 190 else 30
    r += 1
outline(ro, C1, r - len(DICAS), C2, r - 1)
r += 1

# --- fontes
section(ro, r, C1, C2, "FONTES CONSULTADAS", GREY); r += 1
FONTES = [
    ("japan-guide.com — roteiros sugeridos pelo Japão", "https://www.japan-guide.com/e/e2400.html"),
    ("japan-guide.com — Fuji Shibazakura Festival", "https://www.japan-guide.com/e/e6919.html"),
    ("JNTO (turismo oficial do Japão) — Rota Dourada", "https://www.japan.travel/en/gc/itineraries/long-plan/"),
    ("Tokyo Disney Resort — compra oficial de ingressos", "https://www.tokyodisneyresort.jp/en/ticket/purchase.html"),
    ("Rakuten Travel — feriados nacionais e Golden Week", "https://travel.rakuten.com/contents/usa/en-us/guide/japanese-national-holidays-golden-week/"),
    ("Panrotas — isenção de visto Brasil-Japão prorrogada até 2029", "https://www.panrotas.com.br/mercado/destinos/2026/07/brasil-e-japao-prorrogam-isencao-de-vistos-para-turistas-ate-2029_230682.html"),
    ("Tokyo Cheapo — bate-volta de Tóquio a Kawaguchiko", "https://tokyocheapo.com/entertainment/tokyo-to-kawaguchiko-day-trip-guide/"),
    ("Wise — histórico da cotação iene / real", "https://wise.com/us/currency-converter/jpy-to-brl-rate/history"),
]
for i, (titulo, url) in enumerate(FONTES):
    ro.merge_cells(start_row=r, start_column=C1, end_row=r, end_column=C2)
    cell = ro.cell(row=r, column=C1, value=f"{titulo}  —  {url}")
    cell.hyperlink = url
    cell.font = f(8, False, BLUE, underline="single")
    cell.alignment = left(indent=1)
    paint(ro, C1, r, C2, r, fill(BAND if i % 2 == 0 else "FFFFFF"))
    ro.row_dimensions[r].height = 16
    r += 1
outline(ro, C1, r - len(FONTES), C2, r - 1)
bloco_texto("Pesquisa feita em 15/08/2026. Preços, datas de festival e regras de visto mudam — reconfirme "
            "tudo cerca de 3 meses antes de embarcar.", "FFFFFF", GREY, altura=17)
ROTEIRO_FIM = r - 1

# =================================================================== POR DIA
pd_ = wb.create_sheet("Por Dia")
pd_.sheet_view.showGridLines = False
pd_.sheet_properties.tabColor = BLUE

DIA_COLS = [("Data", 13), ("Total do dia (R$)", 16), ("Acumulado (R$)", 16),
            ("Alimentação", 15), ("Transporte local", 16), ("Passeios e Ingressos", 18),
            ("Compras", 19), ("Demais categorias", 17), ("Lançamentos", 12)]
for i, (h, w) in enumerate(DIA_COLS):
    pd_.column_dimensions[get_column_letter(i + 1)].width = w

banner(pd_, 1, 9, "GASTOS DIA A DIA",
       "As datas saem da ida e da volta informadas no Painel. Tudo aqui é calculado automaticamente.")

pd_.row_dimensions[6].height = 28
for i, (h, _) in enumerate(DIA_COLS):
    cell = pd_.cell(row=6, column=i + 1, value=h)
    cell.font = f(9, True, "FFFFFF"); cell.fill = fill("3B4A6B"); cell.alignment = center(wrap=True)
pd_["D6"] = "=Listas!$B$11"     # Alimentação
pd_["E6"] = "=Listas!$B$10"     # Transporte local
pd_["F6"] = "=Listas!$B$12"     # Passeios e Ingressos
pd_["G6"] = "=Listas!$B$14"     # Compras e Souvenirs

D1, D2 = 7, 46
for rr in range(D1, D2 + 1):
    pd_.row_dimensions[rr].height = 18
    band = fill("FFFFFF") if (rr - D1) % 2 == 0 else fill(BAND)
    if rr == D1:
        pd_.cell(row=rr, column=1, value='=IF(Painel!$C$8="","",Painel!$C$8)')
    else:
        pd_.cell(row=rr, column=1,
                 value=f'=IF($A{rr-1}="","",IF($A{rr-1}+1>Painel!$C$9,"",$A{rr-1}+1))')
    pd_.cell(row=rr, column=2, value=f'=IF($A{rr}="","",SUMIFS(Gastos!$J${FIRST}:$J${LAST},'
                                     f'Gastos!$A${FIRST}:$A${LAST},$A{rr}))')
    pd_.cell(row=rr, column=3, value=f'=IF($A{rr}="","",SUM($B${D1}:$B{rr}))')
    for col in (4, 5, 6, 7):
        L = get_column_letter(col)
        pd_.cell(row=rr, column=col,
                 value=f'=IF($A{rr}="","",SUMIFS(Gastos!$J${FIRST}:$J${LAST},'
                       f'Gastos!$A${FIRST}:$A${LAST},$A{rr},Gastos!$B${FIRST}:$B${LAST},{L}$6))')
    pd_.cell(row=rr, column=8, value=f'=IF($A{rr}="","",$B{rr}-SUM($D{rr}:$G{rr}))')
    pd_.cell(row=rr, column=9, value=f'=IF($A{rr}="","",COUNTIFS(Gastos!$A${FIRST}:$A${LAST},$A{rr}))')
    for col in range(1, 10):
        cell = pd_.cell(row=rr, column=col)
        cell.font = f(10, col == 2, INK); cell.fill = band
        cell.border = Border(bottom=thin)
        cell.alignment = center() if col in (1, 9) else right()
        cell.number_format = F_DATE if col == 1 else (F_INT if col == 9 else F_BRL)

tot = D2 + 1
pd_.row_dimensions[tot].height = 22
pd_.cell(row=tot, column=1, value="TOTAL").font = f(10, True, "FFFFFF")
pd_.cell(row=tot, column=1).alignment = center()
for col in range(2, 10):
    if col == 3:
        continue
    L = get_column_letter(col)
    cell = pd_.cell(row=tot, column=col, value=f"=SUM({L}{D1}:{L}{D2})")
    cell.font = f(10, True, "FFFFFF"); cell.alignment = right()
    cell.number_format = F_INT if col == 9 else F_BRL
paint(pd_, 1, tot, 9, tot, fill(INK))
outline(pd_, 1, 6, 9, tot)
pd_.conditional_formatting.add(f"B{D1}:B{D2}", DataBarRule(
    start_type="num", start_value=0, end_type="max", color=BLUE, showValue=True))
pd_.freeze_panes = f"A{D1}"

# =================================================================== LISTAS
ls = wb.create_sheet("Listas")
ls.sheet_view.showGridLines = False
ls.sheet_properties.tabColor = GREY

for col, w in (("A", 2.5), ("B", 38), ("C", 2.5), ("D", 32), ("E", 2.5), ("F", 16),
               ("G", 2.5), ("H", 16), ("I", 2.5), ("J", 26), ("K", 2.5), ("L", 18), ("M", 2.5)):
    ls.column_dimensions[col].width = w

banner(ls, 2, 12, "LISTAS DOS MENUS SUSPENSOS",
       "Edite, acrescente ou renomeie itens aqui — os menus e as tabelas do Painel acompanham sozinhos.")

for col, titulo, itens in ((2, "CATEGORIAS", CATEGORIAS), (4, "FORMAS DE PAGAMENTO", PAGAMENTOS),
                           (6, "STATUS DO GASTO", STATUS), (8, "QUEM PAGOU", PESSOAS),
                           (10, "CIDADES / LOCAIS", CIDADES), (12, "STATUS DA RESERVA", STATUS_RESERVA)):
    cell = ls.cell(row=6, column=col, value=titulo)
    cell.font = f(9, True, "FFFFFF"); cell.fill = fill("3B4A6B"); cell.alignment = center()
    for i in range(CAP):
        rr = 7 + i
        c = ls.cell(row=rr, column=col, value=itens[i] if i < len(itens) else None)
        c.font = f(10, False, INPUT_FG if i < len(itens) else INK)
        c.fill = fill(INPUT_BG if i < len(itens) else "FFFFFF")
        c.alignment = left(indent=1); c.border = Border(bottom=thin)
        ls.row_dimensions[rr].height = 18
    outline(ls, col, 6, col, 6 + CAP)
ls.row_dimensions[6].height = 22
ls.merge_cells(start_row=6 + CAP + 2, start_column=2, end_row=6 + CAP + 2, end_column=12)
nota = ls.cell(row=6 + CAP + 2, column=2,
               value="Não deixe linha em branco no meio de uma lista: o menu suspenso para na primeira lacuna.")
nota.font = f(9, False, GREY, italic=True); nota.alignment = left(indent=1)

# =================================================================== nomes
def dinamica(col):
    return f"OFFSET(Listas!${col}$7,0,0,MAX(1,COUNTA(Listas!${col}$7:${col}${6+CAP})),1)"


for nome, ref in (
    ("CATEGORIAS", dinamica("B")), ("PAGAMENTOS", dinamica("D")), ("STATUS", dinamica("F")),
    ("PESSOAS", dinamica("H")), ("CIDADES", dinamica("J")), ("STATUS_RESERVA", dinamica("L")),
    ("MOEDAS", f"Painel!$E$8:$E${7+len(COTACOES)}"),
    ("TAXAS", f"Painel!$F$8:$F${7+len(COTACOES)}"),
):
    wb.defined_names.add(DefinedName(nome, attr_text=ref))

# =================================================================== impressao
for ws, titulos in ((pa, None), (gs, "6:6"), (ro, None), (pd_, "6:6"), (ls, None)):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    if titulos:
        ws.print_title_rows = titulos

pa.print_area = f"A1:I{PAINEL_FIM}"
gs.print_area = "A1:K206"
ro.print_area = f"A1:K{ROTEIRO_FIM}"
pd_.print_area = f"A1:I{tot}"
ls.print_area = f"A1:M{6+CAP+2}"

wb.active = 0
wb.save(OUT)
print("ok ->", OUT)
