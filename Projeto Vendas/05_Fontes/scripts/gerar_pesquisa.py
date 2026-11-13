# -*- coding: utf-8 -*-
"""Gera a planilha, os gráficos e os números do relatório da pesquisa de mercado.
Uso:  python gerar_pesquisa.py   (rodar a partir de qualquer pasta)
"""
import os, sys, json
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from dados_pesquisa import *  # noqa

RAIZ = os.path.abspath(os.path.join(AQUI, "..", ".."))
PESQ = os.path.join(RAIZ, "01_Pesquisa_de_Mercado")
TRENDS = os.path.join(RAIZ, "05_Fontes", "google_trends")
os.makedirs(os.path.join(PESQ, "graficos"), exist_ok=True)

# ------------------------------------------------------------ parâmetros (todos editáveis na planilha)
P = dict(
    ml_fixo_abaixo79=6.75,          # ESTIMADO: custo operacional ML por unidade < R$79 (faixa R$5,65–8,75 por peso desde 02/03/2026)
    ml_frete_05=20.00, ml_frete_1=22.00, ml_frete_2=25.00,   # ESTIMADO: frete grátis pago pelo vendedor >= R$79 (com reputação)
    sh_c1=0.20, sh_f1=4.50,         # Shopee CNPJ até R$79,99 (taxa fixa R$4,50 desde 01/10/2026)
    sh_c2=0.14, sh_f2=16.00,        # R$80–99,99
    sh_f3=20.00,                    # R$100–199,99
    sh_f4=26.00,                    # >= R$200
    amz_item=2.00,                  # plano Individual (R$2/item); Profissional = R$19/mês
    amz_frete_05=18.00, amz_frete_1=20.00, amz_frete_2=23.00,  # ESTIMADO DBA >= R$79
    mgl_com=0.16, mgl_fixo=5.00,    # Magalu (faixa 10–20% + R$5,00/item)
    mgl_frete_05=15.00, mgl_frete_1=18.00, mgl_frete_2=22.00,  # ESTIMADO coparticipação frete >= R$79
    imposto=0.04,                   # Simples Nacional Anexo I 1ª faixa (MEI: DAS fixo R$82,05/mês)
    ads=0.05,                       # verba de anúncios patrocinados
    frete_compra=0.06,              # frete do fornecedor até você (% do custo)
    emb_leve=1.20, emb_media=2.00, emb_pesada=3.50,
    cambio=CAMBIO_USD, frete_int_usd_kg=8.0, ii_rts=0.60, icms_sc=0.17,
)

CANAIS = ["Mercado Livre (Clássico)", "Shopee (CNPJ)", "Amazon (DBA)", "Magalu"]


def frete_peso(peso, a, b, c):
    return a if peso <= 0.5 else (b if peso <= 1 else c)


def embalagem(peso):
    return P["emb_leve"] if peso <= 0.15 else (P["emb_media"] if peso <= 0.5 else P["emb_pesada"])


def tarifas(prod, canal, preco, w=None):
    """retorna (comissão %, tarifa fixa R$, frete pago pelo vendedor R$)"""
    w = prod["peso_kg"] if w is None else w
    if canal.startswith("Mercado"):
        if preco < 79:
            return prod["categoria_ml"], P["ml_fixo_abaixo79"], 0.0
        return prod["categoria_ml"], 0.0, frete_peso(w, P["ml_frete_05"], P["ml_frete_1"], P["ml_frete_2"])
    if canal.startswith("Shopee"):
        if preco < 80:
            return P["sh_c1"], P["sh_f1"], 0.0
        fixo = P["sh_f2"] if preco < 100 else (P["sh_f3"] if preco < 200 else P["sh_f4"])
        return P["sh_c2"], fixo, 0.0
    if canal.startswith("Amazon"):
        fr = 0.0 if preco < 79 else frete_peso(w, P["amz_frete_05"], P["amz_frete_1"], P["amz_frete_2"])
        return prod["categoria_amz"], P["amz_item"], fr
    fr = 0.0 if preco < 79 else frete_peso(w, P["mgl_frete_05"], P["mgl_frete_1"], P["mgl_frete_2"])
    return P["mgl_com"], (P["mgl_fixo"] if preco >= 10 else 0.0), fr


def preco_canal(prod, canal):
    return {"Mercado": prod["preco_ml"], "Shopee": prod["preco_shopee"], "Amazon": prod["preco_amz"], "Magalu": prod["preco_mgl"]}[canal.split()[0]]


def custo_unit(custo, peso):
    return custo * (1 + P["frete_compra"]) + embalagem(peso)


def calc(prod, canal, kit=False):
    if kit:
        preco, custo0, peso = prod["kit_preco"], prod["kit_custo"], prod["kit_peso"]
    else:
        preco, custo0, peso = preco_canal(prod, canal), prod["custo_nac"], prod["peso_kg"]
    c, f, fr = tarifas(prod, canal, preco, peso)
    custo = custo_unit(custo0, peso)
    lucro = preco * (1 - c - P["imposto"] - P["ads"]) - f - fr - custo
    return dict(preco=preco, com=c, fixo=f, frete=fr, custo=custo, lucro=lucro, margem=lucro / preco, roi=lucro / custo)


def importacao(prod):
    cif_usd = prod["fob_usd"] + P["frete_int_usd_kg"] * prod["peso_kg"]
    cif = cif_usd * P["cambio"]
    rts = cif * (1 + P["ii_rts"]) / (1 - P["icms_sc"])
    return cif, rts

# ------------------------------------------------------------ cálculos
linhas = []
for pr in PRODUTOS:
    for cn in CANAIS:
        r = calc(pr, cn)
        linhas.append(dict(id=pr["id"], produto=pr["curto"], canal=cn, **r))
df = pd.DataFrame(linhas)
lk = []
for pr in PRODUTOS:
    if "kit_preco" in pr:
        for cn in CANAIS:
            lk.append(dict(id=pr["id"], produto=pr["kit_nome"], canal=cn, **calc(pr, cn, kit=True)))
dk = pd.DataFrame(lk)

# estatísticas de preço da amostra
an = pd.DataFrame(ANUNCIOS, columns=["id", "canal", "anuncio", "preco", "nota", "avaliacoes", "frete_gratis", "selo", "obs"])
est = {}
for pr in PRODUTOS:
    d = {}
    for cn in [ML, SH]:
        s = an[(an.id == pr["id"]) & (an.canal == cn)].preco.dropna()
        d[cn] = dict(n=int(an[(an.id == pr["id"]) & (an.canal == cn)].shape[0]), min=s.min() if len(s) else None,
                     med=s.median() if len(s) else None, max=s.max() if len(s) else None)
    rv = an[(an.id == pr["id"])].avaliacoes.dropna()
    d["max_aval"] = int(rv.max()) if len(rv) else 0
    est[pr["id"]] = d

# notas
ranking = []
for pr in PRODUTOS:
    n = dict(NOTAS[pr["id"]])
    mm = df[(df.id == pr["id"]) & (df.canal.isin([CANAIS[0], CANAIS[1]]))].margem.mean()
    n["margem"] = max(0, min(10, round(mm * 100 / 3, 1)))
    total = sum(n[k] * PESOS[k] for k in PESOS)
    ranking.append(dict(id=pr["id"], produto=pr["curto"], **n, margem_media=mm, total=round(total, 2)))
rk = pd.DataFrame(ranking).sort_values("total", ascending=False).reset_index(drop=True)
rk["posicao"] = rk.index + 1
ESCOLHIDOS = ["P07", "P05", "P10", "P06", "P04"]  # 5 primeiros do ranking; P04 x P01 (empate) decidido pela margem do kit

# ------------------------------------------------------------ gráficos
C = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"]
INK, INK2, GRID, BG = "#0b0b0b", "#52514e", "#e4e3df", "#fcfcfb"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GRID, "axes.labelcolor": INK2,
                     "xtick.color": INK2, "ytick.color": INK2, "figure.facecolor": BG, "axes.facecolor": BG})

t5 = pd.read_csv(os.path.join(TRENDS, "trends_5y_BR_normalizado.csv"), index_col=0, parse_dates=True)
termos = [("toalha microfibra", "Toalha microfibra"), ("carteira slim", "Carteira slim"), ("ALT necessaire", "Necessaire / organizador"),
          ("kit churrasco", "Kit churrasco"), ("ALT pincel de maquiagem", "Pincel de maquiagem")]
meses = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
fig, ax = plt.subplots(figsize=(10, 5.4), dpi=150)
ax.axvspan(10.5, 12.4, color="#f1efe8", zorder=0)
ax.text(11.45, 0.5, "Black Friday\ne Natal", ha="center", va="bottom", color=INK2, fontsize=9)
rot = []
for i, (t, nome) in enumerate(termos):
    s = t5[t]; f = s.groupby(s.index.month).mean() / s.mean()
    ax.plot(range(1, 13), f.values, color=C[i], lw=2, marker="o", ms=4, zorder=3, label=nome)
    rot.append([f.values[11], f"{nome} ({f.values[11]:.2f}×)"])
rot.sort(key=lambda r: -r[0])
for k in range(1, len(rot)):
    rot[k][0] = min(rot[k][0], rot[k - 1][0] - 0.06)
for y, txt in rot:
    ax.text(12.25, y, txt, va="center", color=INK, fontsize=9)
ax.axhline(1, color=INK2, lw=0.8, ls="--")
ax.set_xticks(range(1, 13)); ax.set_xticklabels(meses); ax.set_xlim(0.7, 12.3)
ax.set_ylabel("Interesse no mês ÷ média anual (tracejado = 1,0)"); ax.set_ylim(0.4, 1.95)
ax.grid(axis="y", color=GRID, lw=0.6); ax.spines[["top", "right"]].set_visible(False)
ax.set_title("Sazonalidade no Google (Brasil, média 2021–2026)", loc="left", color=INK, fontsize=13, pad=12)
ax.legend(loc="upper left", frameon=False, fontsize=8, ncol=3)
fig.text(0.01, 0.01, "Fonte: Google Trends, coleta em 13/11/2026. Cada curva dividida pela sua própria média.", color=INK2, fontsize=8)
plt.tight_layout(rect=(0, 0.03, 0.86, 1)); fig.savefig(os.path.join(PESQ, "graficos", "1_sazonalidade_google_trends.png")); plt.close()

# ranking
fig, ax = plt.subplots(figsize=(9, 5.6), dpi=150)
r2 = rk.iloc[::-1]
cores = [C[0] if i in ESCOLHIDOS else "#b9b7b0" for i in r2.id]
ax.barh(r2.produto, r2.total, color=cores, height=0.62)
for y, (v, i) in enumerate(zip(r2.total, r2.id)):
    ax.text(v + 0.08, y, f"{v:.1f}" + ("  ✔ escolhido" if i in ESCOLHIDOS else ""), va="center", color=INK, fontsize=9)
ax.set_xlim(0, 10); ax.set_xlabel("Nota ponderada (0–10)")
ax.grid(axis="x", color=GRID, lw=0.6); ax.spines[["top", "right", "left"]].set_visible(False)
ax.set_title("Ranking dos 12 candidatos", loc="left", color=INK, fontsize=13, pad=12)
fig.text(0.01, 0.01, "Pesos: demanda 25%, margem 25%, concorrência 15%, risco 15%, logística 10%, sazonalidade 10%.", color=INK2, fontsize=8)
plt.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(PESQ, "graficos", "2_ranking_candidatos.png")); plt.close()

# margem por canal (escolhidos)
fig, ax = plt.subplots(figsize=(10, 5.4), dpi=150)
w = 0.2
for j, cn in enumerate(CANAIS):
    vals = [dk[(dk.id == i) & (dk.canal == cn)].lucro.iloc[0] for i in ESCOLHIDOS]
    xs = [k + (j - 1.5) * (w + 0.02) for k in range(len(ESCOLHIDOS))]
    ax.bar(xs, vals, width=w, color=C[j], label=cn)
    for x, v in zip(xs, vals):
        ax.text(x, v + (0.4 if v >= 0 else -1.6), f"{v:.0f}", ha="center", fontsize=8, color=INK)
ax.axhline(0, color=INK2, lw=0.8)
ax.set_xticks(range(len(ESCOLHIDOS)))
ax.set_xticklabels([next(p["kit_nome"] for p in PRODUTOS if p["id"] == i).replace(" + ", "\n+ ").replace(" (", "\n(") for i in ESCOLHIDOS], fontsize=8)
ax.set_ylabel("Lucro líquido por unidade (R$)"); ax.grid(axis="y", color=GRID, lw=0.6)
ax.spines[["top", "right"]].set_visible(False); ax.set_ylim(0, 33)
ax.legend(frameon=False, ncol=4, loc="upper left", fontsize=8)
ax.set_title("Lucro por venda no formato recomendado (kit/presente), por marketplace", loc="left", color=INK, fontsize=13, pad=12)
fig.text(0.01, 0.01, "Após comissão, tarifa fixa, frete pago pelo vendedor, 4% de imposto, 5% de anúncios, custo, frete de compra e embalagem. Valores ESTIMADOS.", color=INK2, fontsize=8)
plt.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(PESQ, "graficos", "3_lucro_por_canal.png")); plt.close()

# ------------------------------------------------------------ planilha
wb = Workbook()
H = Font(bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="2A78D6")
EST = PatternFill("solid", fgColor="FFF4D6"); OK = PatternFill("solid", fgColor="E3F4EA"); BAD = PatternFill("solid", fgColor="FBE3E1")
thin = Side(style="thin", color="D9D9D9"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
BRL = 'R$ #,##0.00'; PCT = '0.0%'


def cab(ws, row, cols, widths=None):
    for j, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=j, value=c); cell.font = H; cell.fill = HF; cell.alignment = Alignment(wrap_text=True, vertical="center"); cell.border = BOX
    if widths:
        for j, w_ in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w_
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def titulo(ws, txt, sub=None):
    ws["A1"] = txt; ws["A1"].font = Font(bold=True, size=14)
    if sub:
        ws["A2"] = sub; ws["A2"].font = Font(italic=True, color="52514E")

# 1. Resumo / ranking
ws = wb.active; ws.title = "1_Ranking"
titulo(ws, "Ranking dos candidatos — Pesquisa de mercado " + DATA_COLETA,
       "Notas 0–10 (10 = melhor). Concorrência 10 = fácil de competir; Risco 10 = baixo risco. A nota de margem vem da aba 4.")
cols = ["Posição", "ID", "Produto", "Demanda", "Margem", "Concorrência", "Risco (legal/devolução)", "Logística", "Sazonalidade", "Nota final", "Decisão"]
cab(ws, 4, cols, [8, 6, 30, 10, 10, 13, 14, 10, 13, 11, 34])
pesos_row = ["", "", "Peso", PESOS["demanda"], PESOS["margem"], PESOS["concorrencia"], PESOS["risco"], PESOS["logistica"], PESOS["sazonalidade"], "", ""]
for j, v in enumerate(pesos_row, 1):
    c = ws.cell(row=5, column=j, value=v); c.font = Font(italic=True); c.number_format = '0%'
decisao = {"P07": "ESCOLHIDO — verão/praia, melhor margem", "P05": "ESCOLHIDO — presente masculino de Natal", "P10": "ESCOLHIDO — férias de verão",
           "P06": "ESCOLHIDO — Natal; kit com maleta de bambu", "P04": "ESCOLHIDO — só no kit com necessaire (Shopee avulso não paga)",
           "P01": "RESERVA — só se custo <= R$ 18 (frete pesa)",
           "P03": "RESERVA — sazonal, margem fina, confirmar INMETRO",
           "P11": "DESCARTADO — Anatel + devolução alta", "P12": "DESCARTADO — Anatel + marcas dominam",
           "P02": "DESCARTADO — demanda baixa", "P08": "DESCARTADO — ticket baixo, sem sazonalidade",
           "P09": "DESCARTADO — ticket baixo, guerra de preço"}
for i, r in rk.iterrows():
    row = 6 + i
    vals = [r.posicao, r.id, r.produto, r.demanda, None, r.concorrencia, r.risco, r.logistica, r.sazonalidade, None, decisao[r.id]]
    for j, v in enumerate(vals, 1):
        ws.cell(row=row, column=j, value=v).border = BOX
    # margem = média das margens ML e Shopee da aba 4, /3, limitada a 0–10
    ws.cell(row=row, column=5, value=f'=MAX(0,MIN(10,ROUND(AVERAGEIFS(\'4_Custos_e_Margens\'!$P:$P,\'4_Custos_e_Margens\'!$A:$A,B{row},\'4_Custos_e_Margens\'!$C:$C,"<>Amazon (DBA)",\'4_Custos_e_Margens\'!$C:$C,"<>Magalu")*100/3,1)))')
    ws.cell(row=row, column=10, value=f"=ROUND(SUMPRODUCT(D{row}:I{row},$D$5:$I$5),2)").font = Font(bold=True)
    ws.cell(row=row, column=11).fill = OK if r.id in ESCOLHIDOS else (BAD if "DESCART" in decisao[r.id] else EST)

# 2. Candidatos
ws = wb.create_sheet("2_Candidatos")
titulo(ws, "Ficha dos 12 candidatos", "Amarelo = valor ESTIMADO (confirmar com fornecedor).")
cols = ["ID", "Produto", "Demanda e sazonalidade", "ML: nº anúncios", "ML preço mín", "ML mediano", "ML máx",
        "Shopee: nº anúncios", "Shopee mín", "Shopee mediano", "Shopee máx", "Maior nº de avaliações visto",
        "Custo atacado nacional (ESTIMADO)", "Peso c/ embalagem kg (ESTIMADO)", "Caixa cm", "Fragilidade",
        "Regras legais / marketplace", "Risco de devolução", "Ideia de kit (ticket maior)", "Busca no ML"]
cab(ws, 4, cols, [6, 30, 48, 9, 10, 10, 10, 9, 10, 10, 10, 12, 13, 11, 11, 14, 60, 34, 34, 40])
for i, pr in enumerate(PRODUTOS):
    row = 5 + i; e = est[pr["id"]]
    vals = [pr["id"], pr["nome"], pr["notas_dem"], e[ML]["n"], e[ML]["min"], e[ML]["med"], e[ML]["max"],
            e[SH]["n"], e[SH]["min"], e[SH]["med"], e[SH]["max"], e["max_aval"], pr["custo_nac"], pr["peso_kg"], pr["dims"],
            pr["fragil"], pr["legal"], pr["devol"], pr["kit"], BUSCAS[pr["id"]]]
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=row, column=j, value=v); c.alignment = WRAP; c.border = BOX
        if j in (5, 6, 7, 9, 10, 11, 13): c.number_format = BRL
    ws.cell(row=row, column=13).fill = EST; ws.cell(row=row, column=14).fill = EST
    ws.row_dimensions[row].height = 120

# 3. Anúncios
ws = wb.create_sheet("3_Anuncios_coletados")
titulo(ws, f"Amostra de anúncios ativos coletada em {DATA_COLETA}",
       "Fonte: resultados indexados do Mercado Livre e agregadores de ofertas da Shopee. Campos vazios = não exibidos na fonte. Validar com a aba 8.")
cols = ["ID", "Produto", "Canal", "Anúncio / vendedor", "Preço R$", "Nota", "Nº avaliações", "Frete grátis", "Full / selo", "Observação"]
cab(ws, 4, cols, [6, 26, 14, 44, 11, 7, 12, 11, 16, 44])
nome = {p["id"]: p["curto"] for p in PRODUTOS}
for i, a in enumerate(ANUNCIOS):
    row = 5 + i
    vals = [a[0], nome[a[0]], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]]
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=row, column=j, value=v); c.border = BOX
    ws.cell(row=row, column=5).number_format = BRL
tab = Table(displayName="Anuncios", ref=f"A4:J{4 + len(ANUNCIOS)}"); tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
ws.add_table(tab)

# 5. Parâmetros (criada antes da 4 para as fórmulas; reordenada no fim)
wp = wb.create_sheet("5_Parametros_Tarifas")
titulo(wp, "Parâmetros de tarifas, impostos e custos (edite a coluna C e a planilha recalcula)",
       "Amarelo = ESTIMADO. Confirme no painel de cada marketplace antes de precificar.")
cab(wp, 4, ["Chave", "Descrição", "Valor", "Fonte / status"], [22, 70, 12, 70])
PARAM_DESC = [
    ("ml_fixo_abaixo79", "Mercado Livre: custo operacional por unidade em anúncios abaixo de R$ 79", "ESTIMADO. Desde 02/03/2026 varia por peso/medida (R$ 5,65 a R$ 8,75 até 2 kg). Ver painel Tarifas."),
    ("ml_frete_05", "ML: frete grátis pago pelo vendedor (>= R$ 79) até 0,5 kg", "ESTIMADO com desconto de reputação verde. Sem reputação, o custo é maior."),
    ("ml_frete_1", "ML: frete pago pelo vendedor 0,5–1 kg", "ESTIMADO"),
    ("ml_frete_2", "ML: frete pago pelo vendedor 1–2 kg", "ESTIMADO"),
    ("sh_c1", "Shopee CNPJ: comissão até R$ 79,99", "Tabela Shopee 2026 (programa de frete grátis já incluído)"),
    ("sh_f1", "Shopee: taxa fixa por item até R$ 79,99", "R$ 4,00 → R$ 4,50 desde 01/10/2026"),
    ("sh_c2", "Shopee CNPJ: comissão a partir de R$ 80", "Tabela Shopee 2026"),
    ("sh_f2", "Shopee: taxa fixa R$ 80–99,99", "Tabela Shopee 2026"),
    ("sh_f3", "Shopee: taxa fixa R$ 100–199,99", "Tabela Shopee 2026"),
    ("sh_f4", "Shopee: taxa fixa >= R$ 200", "Tabela Shopee 2026"),
    ("amz_item", "Amazon: tarifa por item no plano Individual", "Plano Profissional: R$ 19/mês (vale a partir de ~10 vendas/mês)"),
    ("amz_frete_05", "Amazon DBA: frete pago pelo vendedor (>= R$ 79) até 0,5 kg", "ESTIMADO"),
    ("amz_frete_1", "Amazon DBA: frete 0,5–1 kg", "ESTIMADO"),
    ("amz_frete_2", "Amazon DBA: frete 1–2 kg", "ESTIMADO"),
    ("mgl_com", "Magalu: comissão (varia 10–20% por categoria)", "ESTIMADO 16%. Há promoção de 9,9% para novos sellers"),
    ("mgl_fixo", "Magalu: tarifa fixa por item (itens >= R$ 10)", "R$ 5,00 (R$ 5,50 em moda/beleza/pet desde 10/09/2026)"),
    ("mgl_frete_05", "Magalu: coparticipação no frete (>= R$ 79) até 0,5 kg", "ESTIMADO"),
    ("mgl_frete_1", "Magalu: frete 0,5–1 kg", "ESTIMADO"),
    ("mgl_frete_2", "Magalu: frete 1–2 kg", "ESTIMADO"),
    ("imposto", "Imposto sobre a venda (Simples Nacional, Anexo I, 1ª faixa)", "Se for MEI: 0% por venda + DAS fixo R$ 82,05/mês (comércio, 2026). Teto MEI R$ 81 mil/ano"),
    ("ads", "Verba de anúncios patrocinados (% do preço)", "Premissa de gestão"),
    ("frete_compra", "Frete do fornecedor até você (% do custo)", "ESTIMADO"),
    ("emb_leve", "Embalagem até 150 g (envelope + etiqueta)", "ESTIMADO"),
    ("emb_media", "Embalagem até 500 g (caixa/envelope reforçado + plástico bolha)", "ESTIMADO"),
    ("emb_pesada", "Embalagem acima de 500 g (caixa + proteção)", "ESTIMADO"),
    ("cambio", "Dólar comercial (R$)", "Fechamento 10/11/2026 = R$ 5,58"),
    ("frete_int_usd_kg", "Frete internacional expresso US$/kg (comparação de importação)", "ESTIMADO"),
    ("ii_rts", "Imposto de Importação no Regime de Tributação Simplificada (fora do benefício PF/PRC)", "Receita Federal, página 'Quanto vou pagar de imposto' (atualizada 14/09/2026)"),
    ("icms_sc", "ICMS de SC sobre importação simplificada", "Comsefaz, tabela Convênio ICMS 81/2023 (SC = 17%)"),
]
REF = {}
for i, (k, d, f) in enumerate(PARAM_DESC):
    row = 5 + i
    wp.cell(row=row, column=1, value=k); wp.cell(row=row, column=2, value=d)
    c = wp.cell(row=row, column=3, value=P[k]); c.number_format = PCT if P[k] < 1 and k not in ("cambio",) else BRL
    if k in ("cambio", "frete_int_usd_kg"): c.number_format = '0.00'
    wp.cell(row=row, column=4, value=f)
    if "ESTIMADO" in f: c.fill = EST
    for j in range(1, 5): wp.cell(row=row, column=j).border = BOX; wp.cell(row=row, column=j).alignment = WRAP
    REF[k] = f"'5_Parametros_Tarifas'!$C${row}"

# 4. Custos e margens (com fórmulas)
def folha_margens(nome_aba, titulo_txt, kit):
    ws = wb.create_sheet(nome_aba)
    titulo(ws, titulo_txt,
           "Altere o preço (col. D) ou o custo (col. J) e veja o lucro. Tarifas vêm da aba 5. Verde = margem >= 20%; vermelho = < 10%.")
    cols = ["ID", "Produto", "Canal", "Preço de venda", "Peso kg", "Comissão %", "Tarifa fixa", "Frete pago pelo vendedor",
            "Imposto + anúncios", "Custo do produto (ESTIMADO)", "Frete de compra", "Embalagem", "Custo total unitário",
            "Total de tarifas", "Lucro líquido", "Margem %", "ROI sobre custo"]
    cab(ws, 4, cols, [6, 34, 22, 11, 8, 10, 10, 12, 11, 12, 10, 10, 12, 11, 11, 9, 9])
    row = 5
    for pr in PRODUTOS:
        if kit and "kit_preco" not in pr:
            continue
        for cn in CANAIS:
            r = calc(pr, cn, kit=kit)
            D, E = f"D{row}", f"E{row}"
            def fw(a, b, c):
                return f"IF({E}<=0.5,{REF[a]},IF({E}<=1,{REF[b]},{REF[c]}))"
            if cn.startswith("Mercado"):
                com = f"={pr['categoria_ml']}"; fixo = f"=IF({D}<79,{REF['ml_fixo_abaixo79']},0)"; fr = f"=IF({D}<79,0,{fw('ml_frete_05','ml_frete_1','ml_frete_2')})"
            elif cn.startswith("Shopee"):
                com = f"=IF({D}<80,{REF['sh_c1']},{REF['sh_c2']})"
                fixo = f"=IF({D}<80,{REF['sh_f1']},IF({D}<100,{REF['sh_f2']},IF({D}<200,{REF['sh_f3']},{REF['sh_f4']})))"; fr = "=0"
            elif cn.startswith("Amazon"):
                com = f"={pr['categoria_amz']}"; fixo = f"={REF['amz_item']}"; fr = f"=IF({D}<79,0,{fw('amz_frete_05','amz_frete_1','amz_frete_2')})"
            else:
                com = f"={REF['mgl_com']}"; fixo = f"=IF({D}>=10,{REF['mgl_fixo']},0)"; fr = f"=IF({D}<79,0,{fw('mgl_frete_05','mgl_frete_1','mgl_frete_2')})"
            vals = [pr["id"], pr["kit_nome"] if kit else pr["curto"], cn, r["preco"], pr["kit_peso"] if kit else pr["peso_kg"], com, fixo, fr,
                    f"={D}*({REF['imposto']}+{REF['ads']})", pr["kit_custo"] if kit else pr["custo_nac"], f"=J{row}*{REF['frete_compra']}",
                    f"=IF({E}<=0.15,{REF['emb_leve']},IF({E}<=0.5,{REF['emb_media']},{REF['emb_pesada']}))",
                    f"=J{row}+K{row}+L{row}", f"={D}*F{row}+G{row}+H{row}", f"={D}-N{row}-I{row}-M{row}", f"=O{row}/{D}", f"=O{row}/M{row}"]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=row, column=j, value=v); c.border = BOX
                if j in (4, 7, 8, 9, 10, 11, 12, 13, 14, 15): c.number_format = BRL
                if j in (6, 16, 17): c.number_format = PCT
            ws.cell(row=row, column=4).fill = EST; ws.cell(row=row, column=10).fill = EST
            if pr["id"] in ESCOLHIDOS: ws.cell(row=row, column=2).font = Font(bold=True)
            row += 1
    ws.conditional_formatting.add(f"P5:P{row-1}", CellIsRule(operator="greaterThanOrEqual", formula=["0.2"], fill=OK))
    ws.conditional_formatting.add(f"P5:P{row-1}", CellIsRule(operator="lessThan", formula=["0.1"], fill=BAD))

folha_margens("4_Custos_e_Margens", "Produto avulso: custo total por unidade e lucro por marketplace (fórmulas vivas)", False)
folha_margens("4b_Kits_Recomendados", "Formato recomendado (kit/presente): custo e lucro por marketplace (fórmulas vivas)", True)

# 6. Importação x nacional
wi = wb.create_sheet("6_Importacao_x_Nacional")
titulo(wi, "Comprar no Brasil ou importar? Custo de chegada por unidade",
       "Importação simplificada por CNPJ: II 60% + ICMS-SC 17% 'por dentro' (Receita Federal). A isenção de II até US$ 50 vale só para PESSOA FÍSICA em site do Remessa Conforme e NÃO serve para revenda.")
cab(wi, 4, ["ID", "Produto", "FOB Alibaba US$ (ESTIMADO)", "Peso kg", "Frete intl. US$", "CIF R$", "Custo importado c/ impostos (CNPJ, RTS)", "Custo atacado nacional (ESTIMADO)", "Mais barato"],
    [6, 28, 14, 9, 12, 11, 18, 16, 16])
for i, pr in enumerate(PRODUTOS):
    r_ = 5 + i
    vals = [pr["id"], pr["curto"], pr["fob_usd"], pr["peso_kg"], f"=D{r_}*{REF['frete_int_usd_kg']}", f"=(C{r_}+E{r_})*{REF['cambio']}",
            f"=F{r_}*(1+{REF['ii_rts']})/(1-{REF['icms_sc']})", pr["custo_nac"], f'=IF(G{r_}<H{r_},"Importar","Nacional")']
    for j, v in enumerate(vals, 1):
        c = wi.cell(row=r_, column=j, value=v); c.border = BOX
        if j in (6, 7, 8): c.number_format = BRL
        if j in (3, 5): c.number_format = '"US$" 0.00'
        if j == 4: c.number_format = '0.00'
    wi.cell(row=r_, column=3).fill = EST; wi.cell(row=r_, column=8).fill = EST
n0 = 5 + len(PRODUTOS) + 1
notas_imp = [
    "Leitura: com 60% de II + 17% de ICMS, o importado chega custando ~1,93× o valor CIF. Para quantidades pequenas, o atacado nacional ganha em quase todos os itens.",
    "Importação formal (DI, container/consolidado) tem II pela TEC do NCM (geralmente 16–20%) + IPI + PIS/COFINS-Importação + ICMS + despachante e armazenagem; só compensa a partir de alguns milhares de reais por pedido. Itajaí tem porto e muitas tradings: vale cotar na 2ª fase.",
    "Comprar como pessoa física via AliExpress/Shopee internacional (II 0% até US$ 50 + ICMS 17%) para REVENDER é irregular: revenda exige nota fiscal de entrada e pode gerar apreensão (pena de perdimento) se houver habitualidade.",
]
for k, t in enumerate(notas_imp):
    c = wi.cell(row=n0 + k, column=1, value=t); wi.merge_cells(start_row=n0 + k, start_column=1, end_row=n0 + k, end_column=9)
    c.alignment = WRAP; wi.row_dimensions[n0 + k].height = 34

# 7. Sazonalidade
wt = wb.create_sheet("7_Sazonalidade_Trends")
titulo(wt, "Google Trends — Brasil (5 anos e 12 meses)", "Índices normalizados entre si com 'kit churrasco' como âncora. Fator nov/dez = interesse do mês ÷ média do ano.")
sz = pd.read_csv(os.path.join(TRENDS, "resumo_sazonalidade.csv"))
cab(wt, 4, ["Termo pesquisado", "Índice médio 5 anos", "Fator novembro", "Fator dezembro", "Mês de pico", "Variação últimos 12 meses x 12 anteriores (%)"], [26, 16, 14, 14, 12, 22])
for i, r in sz.iterrows():
    for j, v in enumerate([r.termo.replace("ALT ", ""), r.indice_medio_5a, r.fator_nov, r.fator_dez, meses[int(r.mes_pico) - 1], r["var_12m_vs_12m_ant_%"]], 1):
        wt.cell(row=5 + i, column=j, value=v).border = BOX

# 8. Validação
wv = wb.create_sheet("8_Validar_10_anuncios")
titulo(wv, "Checklist de validação: abra a busca, ordene por 'Mais vendidos' e anote 10 anúncios por canal",
       "Faça isso nos 5 escolhidos antes de comprar estoque (≈10 min por produto). As fórmulas no topo calculam mín/mediana/máx.")
cab(wv, 4, ["ID", "Canal", "Nº", "Título do anúncio", "Preço R$", "Vendidos (se aparecer)", "Nota", "Reputação vendedor (cor/medalha)", "Frete grátis?", "Full / Shopee Entrega?", "Link"],
    [6, 14, 5, 40, 11, 12, 7, 18, 10, 14, 40])
r_ = 5
for pid in ESCOLHIDOS:
    for cn in ["Mercado Livre", "Shopee"]:
        for k in range(1, 11):
            for j, v in enumerate([pid, cn, k], 1):
                wv.cell(row=r_, column=j, value=v).border = BOX
            for j in range(4, 12): wv.cell(row=r_, column=j).border = BOX
            wv.cell(row=r_, column=5).number_format = BRL
            r_ += 1
wv["M4"] = "Resumo"; wv["M4"].font = Font(bold=True)
rr = 5
for pid in ESCOLHIDOS:
    for cn in ["Mercado Livre", "Shopee"]:
        wv.cell(row=rr, column=13, value=f"{pid} {cn}")
        wv.cell(row=rr, column=14, value=f'=IFERROR(MINIFS(E:E,A:A,"{pid}",B:B,"{cn}"),"")').number_format = BRL
        wv.cell(row=rr, column=15, value=f'=IFERROR(MEDIAN(IF((A5:A{r_-1}="{pid}")*(B5:B{r_-1}="{cn}")*(E5:E{r_-1}<>""),E5:E{r_-1})),"")').number_format = BRL
        wv.cell(row=rr, column=16, value=f'=IFERROR(MAXIFS(E:E,A:A,"{pid}",B:B,"{cn}"),"")').number_format = BRL
        rr += 1
wv["N4"], wv["O4"], wv["P4"] = "Mínimo", "Mediana", "Máximo"
for c in ("N4", "O4", "P4"): wv[c].font = Font(bold=True)
wv.column_dimensions["M"].width = 22

# 9. Pendências
wpd = wb.create_sheet("9_Confirmar")
titulo(wpd, "O que você precisa confirmar antes de comprar estoque")
pend = [
 ("Custos de atacado (todos)", "Pedir cotação a 3 fornecedores por produto (atacado SP/25 de Março, distribuidores de Itajaí/Brusque/Blumenau, Atacado no ML/Shopee).", "Alta"),
 ("Tarifas exatas de cada marketplace", "Painel do vendedor: ML 'Tarifas e faturamento' / calculadora; Shopee Central do Vendedor; Amazon Seller Central; Magalu Seller Portal.", "Alta"),
 ("Frete pago pelo vendedor (ML, Amazon, Magalu)", "Simular no anúncio com peso/medidas reais. Conta nova sem reputação paga frete cheio.", "Alta"),
 ("Enquadramento fiscal", "Abrir MEI (CNAE 4789-0/99 ou 4759-8/99 etc.) ou ME no Simples. Confirmar com contador se SC cobra ICMS-ST/antecipação nas compras de outro estado.", "Alta"),
 ("Mini ventilador — INMETRO", "Pedir ao fornecedor o registro INMETRO ou declaração de que o produto está fora do escopo; consultar o banco de produtos certificados do INMETRO.", "Alta"),
 ("Kit churrasco — material", "Confirmar o tipo de aço inox, a maleta de bambu e se a faca vem com protetor de lâmina para o envio.", "Média"),
 ("Copo térmico — laudos", "Inox 304, tampa livre de BPA, teste de vazamento em 3 unidades da amostra.", "Média"),
 ("Toalha / organizador — etiqueta têxtil", "Etiqueta com composição, CNPJ, país de origem e instruções de lavagem (Conmetro 02/2008).", "Média"),
 ("Validação de preços (aba 8)", "Registrar 10 anúncios no ML e 10 na Shopee de cada escolhido; também olhar Amazon e Magalu.", "Média"),
 ("Prazo do fornecedor", "O cronograma assume pedido em 16/11 e chegada em 18/11 (toalha, SC) e 24/11 (demais, SP; 20/11 é feriado). Se o fornecedor demorar mais, empurre as publicações no cronograma.", "Alta"),
]
cab(wpd, 3, ["Item", "Como confirmar", "Prioridade"], [36, 100, 10])
for i, (a, b, c) in enumerate(pend):
    for j, v in enumerate([a, b, c], 1):
        cell = wpd.cell(row=4 + i, column=j, value=v); cell.alignment = WRAP; cell.border = BOX


# 0. Escolhidos (resumo vivo, ligado à aba 4b)
from dados_anuncios import PRODUTOS as ANUN
we = wb.create_sheet("Escolhidos", 0)
titulo(we, "Produtos escolhidos — formato de venda, preço, margem por marketplace e onde comprar",
       "Valores de custo, lucro e margem são fórmulas ligadas à aba 4b (mude preço/custo lá). Amarelo = ESTIMADO. Atualizado em 14/11/2026.")
cols = ["Ordem", "ID", "Produto (formato de venda)", "Preço de venda", "Custo total por unidade (ESTIMADO)",
        "Lucro ML", "Margem ML", "Lucro Shopee", "Margem Shopee", "Lucro Amazon", "Margem Amazon", "Lucro Magalu", "Margem Magalu",
        "Onde comprar (fornecedor nacional)", "Prazo de entrega (dias úteis, ESTIMADO)", "Fotos de referência do fornecedor", "Por que foi escolhido"]
cab(we, 4, cols, [7, 6, 40, 11, 13, 11, 9, 11, 9, 11, 9, 11, 9, 46, 12, 50, 50])
motivo = {"P07": "1º no ranking: verão/praia (Itapema), busca +18% em 12 meses, leve e quase sem devolução.",
          "P05": "2º: presente masculino de Natal (dez. +38% no Google), 150 g, melhor margem percentual.",
          "P10": "3º: férias de dez./jan.; 'necessaire' tem o 3º maior volume de busca; risco quase nulo.",
          "P06": "4º: maior pico de Natal (dez. +64%) e maior lucro em reais por venda.",
          "P04": "5º: presente feminino (dez. +37%); kit com esponjas dá margem acima de 30%."}
K = "'4b_Kits_Recomendados'"
CAN = ["Mercado Livre (Clássico)", "Shopee (CNPJ)", "Amazon (DBA)", "Magalu"]
for i, a in enumerate(ANUN):
    r_ = 5 + i; pid = a["id"]
    we.cell(row=r_, column=1, value=i + 1); we.cell(row=r_, column=2, value=pid)
    we.cell(row=r_, column=3, value=next(p["kit_nome"] for p in PRODUTOS if p["id"] == pid))
    we.cell(row=r_, column=4, value=f'=SUMIFS({K}!$D:$D,{K}!$A:$A,B{r_},{K}!$C:$C,"{CAN[0]}")').number_format = BRL
    we.cell(row=r_, column=5, value=f'=SUMIFS({K}!$M:$M,{K}!$A:$A,B{r_},{K}!$C:$C,"{CAN[0]}")').number_format = BRL
    for j, cn in enumerate(CAN):
        we.cell(row=r_, column=6 + 2 * j, value=f'=SUMIFS({K}!$O:$O,{K}!$A:$A,B{r_},{K}!$C:$C,"{cn}")').number_format = BRL
        we.cell(row=r_, column=7 + 2 * j, value=f'=SUMIFS({K}!$P:$P,{K}!$A:$A,B{r_},{K}!$C:$C,"{cn}")').number_format = PCT
    we.cell(row=r_, column=14, value=a["fornecedor_nacional"]); we.cell(row=r_, column=15, value=a["prazo_dias_uteis"])
    we.cell(row=r_, column=16, value="Ver 05_Fontes\\Creditos_Imagens.txt (Alibaba)"); we.cell(row=r_, column=17, value=motivo[pid])
    for j in range(1, 18):
        c = we.cell(row=r_, column=j); c.border = BOX; c.alignment = WRAP
    we.cell(row=r_, column=5).fill = EST; we.cell(row=r_, column=15).fill = EST
    we.row_dimensions[r_].height = 48
we.conditional_formatting.add("F5:M9", CellIsRule(operator="lessThan", formula=["0"], fill=BAD))
n0 = 11
for k, t in enumerate(["Reservas: mini ventilador portátil (verão; confirmar INMETRO) e copo térmico 1,2 L (só se o custo for até R$ 18).",
                       "Descartados: fone Bluetooth e smartwatch (Anatel), garrafa motivacional, suporte veicular, removedor de pelos pet.",
                       "Formato de cada kit definido pelas fotos dos fornecedores (Sessão 2): o que aparece nas fotos é o que vai no anúncio."]):
    we.cell(row=n0 + k, column=1, value=t)

wb.move_sheet("5_Parametros_Tarifas", offset=2)
for s in wb.worksheets: s.sheet_view.zoomScale = 90
saida = os.path.join(PESQ, "Pesquisa_de_Mercado.xlsx")
wb.save(saida)

# ------------------------------------------------------------ números para o relatório
out = dict(ranking=rk.to_dict(orient="records"), margens=df.to_dict(orient="records"), kits=dk.to_dict(orient="records"), est=est,
           imp={p["id"]: importacao(p) for p in PRODUTOS})
json.dump(out, open(os.path.join(AQUI, "_resultados.json"), "w"), ensure_ascii=False, indent=1, default=float)
print(rk[["posicao", "id", "produto", "demanda", "margem", "concorrencia", "risco", "logistica", "sazonalidade", "total"]].to_string(index=False))
print(df[df.id.isin(ESCOLHIDOS)][["id", "canal", "preco", "lucro", "margem"]].round(2).to_string(index=False))
print(dk[["id", "produto", "canal", "preco", "custo", "lucro", "margem", "roi"]].round(2).to_string(index=False))
print("salvo:", saida)
