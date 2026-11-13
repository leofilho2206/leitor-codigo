# -*- coding: utf-8 -*-
"""Etapa 3b — cronograma de anúncios e ações (14/11/2026 a 10/01/2027).
Gera em 02_Cronograma: Cronograma_de_Anuncios.xlsx, Agenda_Anuncios.ics e Visualizar_Anuncios.html.
As margens de cada preço (normal e promocional) usam o MESMO modelo de custos da pesquisa (gerar_pesquisa.py)."""
import os, sys, glob, html, datetime as dt, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, AQUI)
from dados_anuncios import PRODUTOS as ANUNCIOS, LIMITES
import dados_pesquisa as DP

# --- importa o modelo de custos/tarifas da pesquisa (fonte única de verdade)
_src = open(os.path.join(AQUI, "gerar_pesquisa.py"), encoding="utf-8").read()
NS = {"__file__": os.path.join(AQUI, "gerar_pesquisa.py")}
exec(_src[:_src.index("# ------------------------------------------------------------ cálculos")], NS)
CANAL_PESQ = {"Mercado Livre": "Mercado Livre (Clássico)", "Shopee": "Shopee (CNPJ)", "Amazon": "Amazon (DBA)", "Magalu": "Magalu"}
PESQ = {p["id"]: p for p in DP.PRODUTOS}

RAIZ = os.path.abspath(os.path.join(AQUI, "..", ".."))
OUT = os.path.join(RAIZ, "02_Cronograma"); os.makedirs(OUT, exist_ok=True)
HOJE = dt.date(2026, 11, 13)
DIAS = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]
FERIADOS = {dt.date(2026, 11, 20): "Dia Nacional de Zumbi e da Consciência Negra (feriado nacional)",
            dt.date(2026, 12, 25): "Natal", dt.date(2027, 1, 1): "Confraternização Universal"}
MK = ["Mercado Livre", "Shopee", "Amazon", "Magalu"]
PASTA_MK = {"Mercado Livre": "mercado_livre", "Shopee": "shopee", "Amazon": "amazon", "Magalu": "magalu"}
ARQ_MK = {"Mercado Livre": "MercadoLivre", "Shopee": "Shopee", "Amazon": "Amazon", "Magalu": "Magalu"}
A = {p["id"]: p for p in ANUNCIOS}

FONTES = {
 "horario": "Pico de compras 18h–23h, máximo às 23h nos marketplaces (dados Elo citados por Central do Varejo, 2026): https://centraldovarejo.com.br/o-que-os-dados-de-remessa-da-melhor-envio-revelam-sobre-o-novo-calendario-do-e-commerce-datas-duplas/ ; 54–56% compram à noite ou de madrugada (Octadesk/Opinion Box): https://centraldovarejo.com.br/54-dos-brasileiros-compram-a-noite-e-67-ja-desistiram-por-experiencia-ruim/ — consulta 14/11/2026",
 "bf": "Campanhas de Black Friday 2026: Shopee 13–29/11, só vendedores selecionados, sem freezing; ML exige reputação verde/amarela na Central de promoções; Amazon: agendar ofertas com ~1 semana de antecedência — https://hunterhub.com.br/blog/quando-e-a-black-friday-2026-vendedor/ ; https://base.com/pt-BR/blog/freezing-black-friday/ — consulta 14/11/2026",
 "13": "13º salário: 1ª parcela até 30/11 e 2ª até 20/12 (Lei 4.749/1965) — https://www.planalto.gov.br/ccivil_03/leis/l4749.htm",
 "feriado": "20/11 é feriado nacional desde 2024 (Lei 14.759/2023) — https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2023/lei/l14759.htm",
}


def brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def lucro(pid, canal, preco):
    p = PESQ[pid]
    c, f, fr = NS["tarifas"](p, CANAL_PESQ[canal], preco, p["kit_peso"])
    custo = NS["custo_unit"](p["kit_custo"], p["kit_peso"])
    l = preco * (1 - c - NS["P"]["imposto"] - NS["P"]["ads"]) - f - fr - custo
    return l, l / preco


def dias_uteis(d0, n):
    d = d0
    while n:
        d += dt.timedelta(days=1)
        if d.weekday() < 5 and d not in FERIADOS:
            n -= 1
    return d


# ------------------------------------------------------------------ premissas de estoque
PEDIDO = dt.date(2026, 11, 16)  # segunda: fechar pedidos depois das cotações de sábado
CHEGADA = {pid: dias_uteis(PEDIDO, A[pid]["prazo_dias_uteis"]) for pid in A}
PUBLICA = {pid: CHEGADA[pid] + dt.timedelta(days=1) for pid in A}
for pid in A:
    assert PESQ[pid]["kit_preco"] == A[pid]["preco"], f"preço divergente {pid}"

ev = []


def add(d, h, mk, pid, acao, obs, preco=None, promo=None, fonte="", dur=30):
    hh, mm = map(int, h.split(":"))
    ev.append(dict(dt=dt.datetime(d.year, d.month, d.day, hh, mm), mk=mk, pid=pid, acao=acao, obs=obs, preco=preco, promo=promo, fonte=fonte, dur=dur))


def nome(pid):
    return A[pid]["nome"] if pid else "Todos os produtos"


# 1) preparação (sem estoque ainda)
add(dt.date(2026, 11, 14), "09:00", "Todos", None, "Preparação: abrir MEI (gov.br) com CNAE de comércio varejista",
    "Portal do Empreendedor (gratuito). Anote CNPJ. Depois peça a inscrição estadual na SEF/SC para emitir NF-e de produto.")
add(dt.date(2026, 11, 14), "10:00", "Todos", None, "Comprar estoque: pedir cotação a 3 fornecedores por produto (WhatsApp)",
    "Peça: preço por 10/30/50 un., prazo de entrega em Itajaí, NF, fotos e autorização de uso das fotos. Lista de fornecedores na aba Premissas.")
for i, mk in enumerate(MK):
    add(dt.date(2026, 11, 14), f"{14 + i}:00", mk, None, f"Preparação: criar conta de vendedor ({mk}) com o CNPJ",
        "Amazon e Magalu podem levar 1–3 dias para aprovar a conta: por isso começamos já.")
add(dt.date(2026, 11, 16), "09:00", "Todos", None, "Preparação: configurar emissor de NF-e e embalagens",
    "Emissor gratuito (Sebrae/ERP de marketplace). Comprar caixas, envelopes, plástico bolha e etiquetas.")
for k, pid in enumerate(A):
    add(PEDIDO, f"14:{k*10:02d}", "Todos", pid, "Comprar estoque: fechar pedido e pagar",
        f"Fornecedor: {A[pid]['fornecedor_nacional']}. Quantidade sugerida no lote 1: 20 un. (ESTIMADO). "
        f"Prazo {A[pid]['prazo_dias_uteis']} dias úteis → chegada prevista {CHEGADA[pid]:%d/%m/%Y} (ESTIMADO; 20/11 é feriado).",
        preco=A[pid]["preco"], fonte=FONTES["feriado"])
for k, pid in enumerate(A):
    add(CHEGADA[pid], f"{14 + k // 4}:{(k % 4) * 15:02d}", "Todos", pid, "Receber estoque, conferir e fotografar",
        "Conferir quantidade e qualidade, comparar com as fotos do fornecedor e refazer a foto principal se o produto for diferente (03_Imagens\\GUIA_DE_FOTOS.txt).")

# 2) publicação (só depois do estoque) — 1 anúncio a cada 15 min a partir das 09:00
cont = {}
for pid in A:
    d = PUBLICA[pid]
    for mk in MK:
        n = cont.get(d, 0); cont[d] = n + 1
        hora = (dt.datetime(2000, 1, 1, 9, 0) + dt.timedelta(minutes=15 * n)).time()
        p = A[pid]; l, m = lucro(pid, mk, p["preco"])
        add(d, f"{hora:%H:%M}", mk, pid, "Publicar anúncio",
            f"Estoque disponível desde {CHEGADA[pid]:%d/%m}. Publicar de manhã para o anúncio estar aprovado e indexado antes do pico das 18h–23h. "
            f"Texto: 04_Anuncios\\{pid}_{p['pasta']}_{ARQ_MK[mk]}.txt. Lucro estimado {brl(l)} ({m:.0%}).",
            preco=p["preco"], fonte=FONTES["horario"])

# 3) Black Friday / Cyber Monday (27–30/11)
BF_INI, BF_FIM = dt.date(2026, 11, 27), dt.date(2026, 11, 30)
DESC = {"Mercado Livre": 0.10, "Shopee": 0.10, "Amazon": 0.08, "Magalu": 0.08}
for k, pid in enumerate(A):
    d = PUBLICA[pid]
    # inscrição na campanha oficial da Shopee (seletiva) assim que publicar
    n = cont.get(("sh", d), 0); cont[("sh", d)] = n + 1
    add(d, f"14:{n*10:02d}", "Shopee", pid, "Ativar promoção: indicar o produto na campanha Black Friday da Shopee",
        "Central de Marketing › Campanhas › Black Friday (13–29/11). Aprovação NÃO garantida (só vendedores selecionados). Se não for aprovado, vale a promoção da loja abaixo.",
        preco=A[pid]["preco"], fonte=FONTES["bf"])
slot = 0
for mk in MK:
    for pid in A:
        pr = round(A[pid]["preco"] * (1 - DESC[mk]) - 0.001, 1) - 0.01 if False else round(A[pid]["preco"] * (1 - DESC[mk]), 2)
        pr = float(f"{int(pr)}.90") if pr - int(pr) >= 0.9 else float(f"{int(pr) - 1}.90") if pr - int(pr) < 0.9 else pr
        l, m = lucro(pid, mk, pr)
        assert l > 0, (pid, mk, pr, l)
        tipo = {"Mercado Livre": "Central de promoções › Desconto próprio (desconto do vendedor)",
                "Shopee": "Central de Marketing › Promoções da loja (desconto) — vale também se a campanha oficial não aprovar",
                "Amazon": "Seller Central › Anunciar › Promoções › Desconto percentual (ou cupom; confira a tarifa do cupom)",
                "Magalu": "Seller Portal › Preço promocional por período (ou campanha Magalu, se convidado)"}[mk]
        h = dt.time(15 + slot // 4, (slot % 4) * 15)
        add(dt.date(2026, 11, 26), f"{h:%H:%M}", mk, pid, f"Ativar promoção Black Friday + Cyber Monday ({DESC[mk]:.0%} OFF)",
            f"{tipo}. Período: 27/11 00:00 até 30/11 23:59 (BF, Cyber Monday e 1ª parcela do 13º). "
            f"Preço promocional {brl(pr)} → lucro estimado {brl(l)} ({m:.0%}). Use como 'preço de' o preço praticado desde a publicação.",
            preco=A[pid]["preco"], promo=pr, fonte=FONTES["bf"] + " | " + FONTES["13"])
        slot += 1
for j, mk in enumerate(MK):
    add(dt.date(2026, 11, 25), f"{16}:{j*15:02d}", mk, None, "Anúncio patrocinado: ativar campanha automática (todos os produtos)",
        "ML: Mercado Ads (Product Ads) · Shopee: Shopee Ads automático · Amazon: Sponsored Products automático · Magalu: Magalu Ads. "
        "Orçamento inicial R$ 10/dia por marketplace (≈5% do faturamento previsto, já descontado na margem).", fonte=FONTES["horario"])
for j, mk in enumerate(MK):
    add(BF_INI, f"{17}:{j*15:02d}", mk, None, "Anúncio patrocinado: dobrar orçamento para o pico da noite (27/11)",
        "Black Friday: suba para R$ 20/dia hoje e volte a R$ 10/dia em 01/12. O pico de compras é das 18h às 23h.", fonte=FONTES["horario"])
add(BF_INI, "12:00", "Todos", None, "Ajustar preço: conferir se as promoções estão ativas e o estoque de cada canal",
    "Abrir cada anúncio no celular, conferir preço com desconto, estoque e prazo. Responder perguntas em até 1 h (influencia a conversão).")
add(BF_FIM, "12:00", "Todos", None, "Ajustar preço: conferir Cyber Monday (último dia das promoções)",
    "As promoções terminam às 23:59. Pagamento da 1ª parcela do 13º até hoje.", fonte=FONTES["13"])
add(dt.date(2026, 12, 1), "09:00", "Todos", None, "Ajustar preço: promoções encerradas — conferir preço normal e analisar vendas",
    "Planilha 01_Pesquisa_de_Mercado: compare as vendas por produto/canal. Produto sem venda em nenhum canal: revisar foto principal e título.")
for j, mk in enumerate(MK):
    add(dt.date(2026, 12, 1), f"{10}:{j*15:02d}", mk, None, "Anúncio patrocinado: voltar orçamento para R$ 10/dia",
        "Pausar anúncios patrocinados de produtos com custo por venda acima de 5% do preço.")

# 4) reposição para o Natal
add(dt.date(2026, 12, 1), "14:00", "Todos", None, "Comprar estoque: pedido de reposição para o Natal",
    "Repor os produtos que venderam (lote 2). Prazo de 2 a 5 dias úteis (ESTIMADO) → chegada até 08/12.")
add(dt.date(2026, 12, 8), "15:00", "Mercado Livre", None, "Comprar estoque: enviar lote ao Full (se o ML tiver liberado o Full para sua conta)",
    "Full = entrega em 1–2 dias e selo 'Full', decisivo nas últimas semanas antes do Natal. Se não estiver liberado, mantenha o envio normal.")

# 5) Natal + 2ª parcela do 13º
NAT_INI, NAT_FIM = dt.date(2026, 12, 14), dt.date(2026, 12, 20)
slot = 0
for mk in MK:
    for pid in ("P06", "P05", "P04"):  # presentes
        pr = A[pid]["preco"]
        cup = 0.05
        l, m = lucro(pid, mk, round(pr * (1 - cup), 2))
        assert l > 0
        h = dt.time(20 + slot // 6, (slot % 6) * 10)
        add(dt.date(2026, 12, 13), f"{h:%H:%M}", mk, pid, "Cupom: 5% OFF de Natal (14/12 a 20/12)",
            f"Cupom da loja de 5% para a semana da 2ª parcela do 13º (até 20/12). Preço com cupom {brl(round(pr*(1-cup),2))} → lucro estimado {brl(l)} ({m:.0%}).",
            preco=pr, promo=round(pr * (1 - cup), 2), fonte=FONTES["13"])
        slot += 1
add(dt.date(2026, 12, 17), "09:00", "Todos", None, "Ajustar preço: último dia recomendado de postagem para chegar antes do Natal",
    "Depois de hoje, entrega até 24/12 só é provável no Full/Flex. Não prometa 'chega antes do Natal' em perguntas sem conferir o prazo do anúncio.")
add(dt.date(2026, 12, 21), "09:00", "Todos", None, "Ajustar preço: fim do cupom de Natal; preparar o verão",
    "A partir de agora o foco é a toalha de praia e o organizador de mala (férias).")

# 6) verão (toalha e organizador)
slot = 0
for mk in MK:
    for pid in ("P07", "P10"):
        h = dt.time(19, slot * 7)
        add(dt.date(2026, 12, 26), f"{h:%H:%M}", mk, pid, "Anúncio patrocinado: reforçar verão (R$ 15/dia até 10/01)",
            "Pico de busca de toalha e mini ventilador em janeiro (Google Trends). Organizador: férias de janeiro.", preco=A[pid]["preco"])
        slot += 1

# 7) revisões semanais de preço
for d in (dt.date(2026, 11, 23), dt.date(2026, 12, 7), dt.date(2026, 12, 28), dt.date(2027, 1, 4)):
    add(d, "08:30", "Todos", None, "Ajustar preço: revisão semanal (preço dos concorrentes, estoque, perguntas)",
        "Comparar o seu preço com os 10 primeiros anúncios de cada canal. Só baixar preço se o lucro estimado continuar positivo (planilha da pesquisa, aba 4b).")
add(dt.date(2027, 1, 10), "10:00", "Todos", None, "Ajustar preço: balanço da temporada e decisão do próximo lote",
    "Ver vendas por produto/canal. Manter os 2–3 melhores, testar 1 reserva (mini ventilador, se o INMETRO estiver ok).")

# ------------------------------------------------------------------ validações
ev.sort(key=lambda e: e["dt"])
vistos = set()
for e in ev:
    assert e["dt"].date() > HOJE, ("data no passado", e)
    assert e["dt"] not in vistos, ("conflito de horário", e["dt"], e["acao"])
    vistos.add(e["dt"])
    if e["acao"].startswith(("Publicar", "Ativar promoção", "Cupom")) and e["pid"]:
        assert e["dt"].date() >= PUBLICA[e["pid"]] or e["acao"].startswith("Ativar promoção Black") and e["dt"].date() >= max(PUBLICA.values()), ("ação antes do estoque", e)
    if e["dt"].date() in FERIADOS:
        e["obs"] += f" (Atenção: {FERIADOS[e['dt'].date()]}.)"


def imgs(pid, mk):
    if not pid or mk == "Todos":
        return ""
    d = os.path.join(RAIZ, "03_Imagens", f"{pid}_{A[pid]['pasta']}", "anuncio", PASTA_MK[mk])
    fs = sorted(os.path.basename(f) for f in glob.glob(os.path.join(d, "*.jpg")))
    assert fs, d
    return os.path.relpath(d, RAIZ).replace("/", "\\") + " → " + ", ".join(fs)


def titulo(pid, mk):
    if not pid:
        return ""
    if mk == "Todos":
        return A[pid]["nome"]
    return A[pid]["titulos"][mk]


# ------------------------------------------------------------------ XLSX
wb = Workbook()
H = Font(bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="2A78D6")
thin = Side(style="thin", color="D9D9D9"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WR = Alignment(wrap_text=True, vertical="top")
COR_ACAO = {"Comprar estoque": "FFF4D6", "Publicar anúncio": "E3F4EA", "Ativar promoção": "FBE3E1", "Cupom": "FBE3E1",
            "Ajustar preço": "EAF1FB", "Anúncio patrocinado": "F1E9FB", "Preparação": "F2F2F2", "Receber estoque": "FFF4D6"}
COLS = ["Data", "Dia da semana", "Horário", "Marketplace", "Produto", "Ação", "Título do anúncio", "Preço (R$)", "Preço promocional (R$)",
        "Lucro estimado no preço praticado (R$)", "Margem", "Imagens a usar", "Observações", "Fonte"]
LARG = [12, 14, 9, 14, 30, 42, 50, 11, 12, 13, 9, 60, 70, 40]


def aba(ws, linhas, titulo_aba):
    ws["A1"] = titulo_aba; ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"Gerado em 14/11/2026. Horário de Brasília. Valores ESTIMADOS de lucro usam o modelo de custos da pesquisa (aba 4b). "
                f"Estoque: pedido em {PEDIDO:%d/%m/%Y}; chegada prevista " + ", ".join(f"{A[p]['nome'].split(' ')[0]} {CHEGADA[p]:%d/%m}" for p in A) + " (ESTIMADO).")
    ws["A2"].font = Font(italic=True, color="52514E")
    for j, c in enumerate(COLS, 1):
        x = ws.cell(row=4, column=j, value=c); x.font = H; x.fill = HF; x.alignment = WR; x.border = BOX
        ws.column_dimensions[get_column_letter(j)].width = LARG[j - 1]
    r = 5
    for e in linhas:
        pr = e["preco"]; pp = e["promo"]
        if e["pid"] and e["mk"] in MK:
            l, m = lucro(e["pid"], e["mk"], pp or pr or A[e["pid"]]["preco"])
        else:
            l = m = None
        vals = [e["dt"].date(), DIAS[e["dt"].weekday()], e["dt"].strftime("%H:%M"), e["mk"], nome(e["pid"]), e["acao"], titulo(e["pid"], e["mk"]),
                pr, pp, l, m, imgs(e["pid"], e["mk"]), e["obs"], e["fonte"]]
        for j, v in enumerate(vals, 1):
            x = ws.cell(row=r, column=j, value=v); x.border = BOX; x.alignment = WR
        ws.cell(row=r, column=1).number_format = "DD/MM/YYYY"
        for j in (8, 9, 10):
            ws.cell(row=r, column=j).number_format = 'R$ #,##0.00'
        ws.cell(row=r, column=11).number_format = '0%'
        cor = next((v for k, v in COR_ACAO.items() if e["acao"].startswith(k)), None)
        if cor:
            ws.cell(row=r, column=6).fill = PatternFill("solid", fgColor=cor)
        r += 1
    ws.freeze_panes = "A5"; ws.auto_filter.ref = f"A4:{get_column_letter(len(COLS))}{max(r - 1, 4)}"


ws = wb.active; ws.title = "Geral"
aba(ws, ev, "Cronograma de anúncios e ações — todos os marketplaces")
for mk in MK:
    aba(wb.create_sheet(mk.replace(" ", "_")), [e for e in ev if e["mk"] in (mk, "Todos")], f"Cronograma — {mk} (inclui ações gerais 'Todos')")
wp = wb.create_sheet("Premissas_e_Fontes")
lin = [("Premissa / fonte", "Detalhe"),
       ("Hoje", "13/11/2026 (sexta). Cronograma começa em 14/11/2026 e vai até 10/01/2027."),
       ("Prazo do fornecedor (ESTIMADO)", "; ".join(f"{A[p]['nome']}: {A[p]['prazo_dias_uteis']} dias úteis ({A[p]['fornecedor_nacional']}) → chega {CHEGADA[p]:%d/%m/%Y}, publica {PUBLICA[p]:%d/%m/%Y}" for p in A)),
       ("Regra de publicação", "Só publicar no dia seguinte à chegada do estoque (conferência e fotos no dia da chegada)."),
       ("Feriado", FONTES["feriado"]),
       ("Horários", "Publicação de manhã (09h–11h) para o anúncio estar aprovado antes do pico; reforço de patrocinado às 17h–19h; configurações de promoção na véspera à noite. " + FONTES["horario"]),
       ("Campanhas Black Friday", FONTES["bf"] + ". Conta nova (sem reputação) dificilmente entra nas campanhas oficiais: por isso o cronograma usa as promoções da própria loja, que não dependem de convite."),
       ("13º salário", FONTES["13"]),
       ("Margens", "Calculadas com as tarifas e custos da aba 4b da Pesquisa_de_Mercado.xlsx (comissão, tarifa fixa, frete pago pelo vendedor, 4% imposto, 5% anúncios). Todas as ações têm lucro positivo (validado no script)."),
       ("Desconto de Black Friday", "ML e Shopee 10%; Amazon e Magalu 8% (tarifas maiores). Preços arredondados para ,90."),
       ("Preço 'de'", "Use como preço de referência o preço praticado desde a publicação. Desconto sobre preço inflado é infração ao Código de Defesa do Consumidor.")]
for i, (a_, b_) in enumerate(lin, 1):
    wp.cell(row=i, column=1, value=a_).alignment = WR; wp.cell(row=i, column=2, value=b_).alignment = WR
    if i == 1:
        wp.cell(row=1, column=1).font = H; wp.cell(row=1, column=1).fill = HF; wp.cell(row=1, column=2).font = H; wp.cell(row=1, column=2).fill = HF
wp.column_dimensions["A"].width = 30; wp.column_dimensions["B"].width = 150; wp.freeze_panes = "A2"
wb.save(os.path.join(OUT, "Cronograma_de_Anuncios.xlsx"))

# ------------------------------------------------------------------ ICS
def esc(t):
    return t.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def fold(l):
    b = l.encode("utf-8"); out = []
    while len(b) > 74:
        cut = 74
        while (b[cut] & 0xC0) == 0x80:
            cut -= 1
        out.append(b[:cut].decode()); b = b" " + b[cut:]
    out.append(b.decode()); return "\r\n".join(out)


ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Projeto Vendas//Cronograma//PT-BR", "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
       "X-WR-CALNAME:Projeto Vendas — Anúncios", "X-WR-TIMEZONE:America/Sao_Paulo",
       "BEGIN:VTIMEZONE", "TZID:America/Sao_Paulo", "BEGIN:STANDARD", "DTSTART:19700101T000000", "TZOFFSETFROM:-0300", "TZOFFSETTO:-0300",
       "TZNAME:-03", "END:STANDARD", "END:VTIMEZONE"]
stamp = "20261114T000000Z"
for i, e in enumerate(ev, 1):
    fim = e["dt"] + dt.timedelta(minutes=e["dur"])
    resumo = f"[{e['mk']}] {e['acao']}" + (f" — {A[e['pid']]['nome']}" if e["pid"] else "")
    desc = e["obs"]
    if e["pid"] and e["mk"] in MK:
        desc = f"Título: {titulo(e['pid'], e['mk'])}\nPreço: {brl(e['preco'] or A[e['pid']]['preco'])}" + (f" → promocional {brl(e['promo'])}" if e["promo"] else "") + f"\nImagens: {imgs(e['pid'], e['mk'])}\n" + desc
    ics += ["BEGIN:VEVENT", f"UID:pv-{i:03d}-{e['dt']:%Y%m%dT%H%M}@projeto-vendas", f"DTSTAMP:{stamp}",
            f"DTSTART;TZID=America/Sao_Paulo:{e['dt']:%Y%m%dT%H%M%S}", f"DTEND;TZID=America/Sao_Paulo:{fim:%Y%m%dT%H%M%S}",
            fold("SUMMARY:" + esc(resumo)), fold("DESCRIPTION:" + esc(desc)),
            "BEGIN:VALARM", "ACTION:DISPLAY", "TRIGGER:-PT30M", fold("DESCRIPTION:" + esc(resumo)), "END:VALARM", "END:VEVENT"]
ics.append("END:VCALENDAR")
open(os.path.join(OUT, "Agenda_Anuncios.ics"), "w", encoding="utf-8", newline="").write("\r\n".join(ics) + "\r\n")

# ------------------------------------------------------------------ HTML
def desc_txt(pid, mk):
    t = open(os.path.join(RAIZ, "04_Anuncios", f"{pid}_{A[pid]['pasta']}_{ARQ_MK[mk]}.txt"), encoding="utf-8").read()
    a = t.index("DESCRIÇÃO (copiar e colar)"); b = t.index("FICHA TÉCNICA")
    return t[a:b].split("-" * 78)[1].strip()


COR_MK = {"Mercado Livre": "#2d3277", "Shopee": "#ee4d2d", "Amazon": "#232f3e", "Magalu": "#0086ff"}
cards = []
for pid in A:
    p = A[pid]
    for mk in MK:
        d = os.path.join(RAIZ, "03_Imagens", f"{pid}_{p['pasta']}", "anuncio", PASTA_MK[mk])
        fs = sorted(glob.glob(os.path.join(d, "*.jpg")))
        rel = [os.path.relpath(f, OUT).replace(os.sep, "/") for f in fs]
        pub = next(e for e in ev if e["pid"] == pid and e["mk"] == mk and e["acao"] == "Publicar anúncio")
        promo = next((e for e in ev if e["pid"] == pid and e["mk"] == mk and e["acao"].startswith("Ativar promoção Black")), None)
        acoes = [e for e in ev if e["pid"] == pid and e["mk"] in (mk, "Todos")]
        l, m = lucro(pid, mk, p["preco"])
        thumbs = "".join(f'<img src="{html.escape(r)}" alt="foto {i+1}" loading="lazy" onclick="this.closest(\'.card\').querySelector(\'.main\').src=this.src">' for i, r in enumerate(rel))
        lis = "".join(f"<li><b>{e['dt']:%d/%m/%Y} ({DIAS[e['dt'].weekday()][:3]}) {e['dt']:%H:%M}</b> — {html.escape(e['acao'])}</li>" for e in acoes)
        cards.append(f"""<article class="card" data-mk="{mk}" data-prod="{pid}">
<div class="bar" style="background:{COR_MK[mk]}">{mk}</div>
<div class="body"><div class="gal"><img class="main" src="{html.escape(rel[0])}" alt="foto principal"><div class="thumbs">{thumbs}</div></div>
<div class="info"><h3>{html.escape(p['titulos'][mk])}</h3><div class="muted">{len(p['titulos'][mk])}/{LIMITES[mk]} caracteres · {html.escape(p['categorias'][mk])}</div>
<div class="price">{brl(p['preco'])}</div><div class="muted">à vista · ou 3x de {brl(p['preco']/3)} (parcelamento conforme o marketplace)</div>
{f'<div class="promo">Black Friday (27–30/11): <b>{brl(promo["promo"])}</b></div>' if promo else ''}
<div class="when">Publicar em <b>{pub['dt']:%d/%m/%Y}</b> ({DIAS[pub['dt'].weekday()]}) às <b>{pub['dt']:%H:%M}</b></div>
<div class="muted">Lucro estimado por venda: {brl(l)} ({m:.0%}) · ESTIMADO</div>
<details><summary>Descrição</summary><pre>{html.escape(desc_txt(pid, mk))}</pre></details>
<details><summary>Ações deste anúncio ({len(acoes)})</summary><ul>{lis}</ul></details></div></div></article>""")
filtros = "".join(f'<button data-f="{mk}">{mk}</button>' for mk in MK)
pagina = f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Visualizar Anúncios</title>
<style>:root{{--bg:#f4f5f7;--card:#fff;--ink:#1b1b1b;--muted:#5d5d5d}}*{{box-sizing:border-box}}body{{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--ink)}}
header{{padding:20px 16px;max-width:1200px;margin:auto}}h1{{margin:0 0 6px;font-size:26px}}.f{{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}}.f button{{border:1px solid #ccc;background:#fff;border-radius:999px;padding:8px 14px;cursor:pointer;font-size:14px}}.f button.on{{background:#1b1b1b;color:#fff}}
main{{max-width:1200px;margin:auto;padding:0 16px 40px;display:grid;gap:18px}}.card{{background:var(--card);border-radius:14px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.06)}}
.bar{{color:#fff;font-weight:700;padding:8px 16px;font-size:14px}}.body{{display:grid;grid-template-columns:minmax(0,420px) 1fr;gap:20px;padding:16px}}
.main{{width:100%;aspect-ratio:1;object-fit:contain;border:1px solid #eee;border-radius:10px;background:#fff}}.thumbs{{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}}.thumbs img{{width:56px;height:56px;object-fit:cover;border:1px solid #ddd;border-radius:6px;cursor:pointer}}
h3{{margin:0 0 4px;font-size:20px;line-height:1.25}}.muted{{color:var(--muted);font-size:13px;margin:4px 0}}.price{{font-size:30px;font-weight:700;margin-top:10px}}.promo{{background:#fde8e4;color:#9b2a14;padding:6px 10px;border-radius:8px;display:inline-block;margin:6px 0}}
.when{{margin:10px 0;padding:8px 10px;background:#eaf3ea;border-radius:8px}}pre{{white-space:pre-wrap;font-family:inherit;font-size:14px;background:#fafafa;padding:10px;border-radius:8px}}summary{{cursor:pointer;font-weight:600;margin-top:10px}}
@media (max-width:760px){{.body{{grid-template-columns:1fr}}}}</style></head><body>
<header><h1>Visualizar anúncios — Projeto Vendas</h1><div class="muted">5 produtos × 4 marketplaces = 20 anúncios. Clique nas miniaturas para trocar a foto. Datas e horários de Brasília. Gerado em 14/11/2026.</div>
<div class="f"><button data-f="todos" class="on">Todos</button>{filtros}</div></header><main>{''.join(cards)}</main>
<script>document.querySelectorAll('.f button').forEach(b=>b.onclick=()=>{{document.querySelectorAll('.f button').forEach(x=>x.classList.remove('on'));b.classList.add('on');
document.querySelectorAll('.card').forEach(c=>c.style.display=(b.dataset.f==='todos'||c.dataset.mk===b.dataset.f)?'':'none');}});</script></body></html>"""
open(os.path.join(OUT, "Visualizar_Anuncios.html"), "w", encoding="utf-8").write(pagina)
json.dump([dict(dt=e["dt"].isoformat(), mk=e["mk"], pid=e["pid"], acao=e["acao"]) for e in ev], open(os.path.join(AQUI, "_eventos.json"), "w"), ensure_ascii=False, indent=0)
print("ok:", len(ev), "eventos; chegada", {p: f"{CHEGADA[p]:%d/%m}" for p in A}, "publica", {p: f"{PUBLICA[p]:%d/%m}" for p in A})
