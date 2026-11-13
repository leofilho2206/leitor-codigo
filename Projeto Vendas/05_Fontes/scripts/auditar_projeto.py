# -*- coding: utf-8 -*-
"""Etapa 4 — auditoria independente do projeto (refaz as contas do zero e confere todos os arquivos).
Uso: python auditar_projeto.py <pasta_com_planilhas_recalculadas_pelo_LibreOffice>"""
import os, sys, re, glob, json, datetime as dt
import openpyxl
from PIL import Image
from icalendar import Calendar

AQUI = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, AQUI)
from dados_anuncios import PRODUTOS, LIMITES
import dados_pesquisa as DP

RAIZ = os.path.abspath(os.path.join(AQUI, "..", ".."))
RC = sys.argv[1]
erros, avisos = [], []
ok = lambda c, m: None if c else erros.append(m)
DIAS = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]
HOJE = dt.date(2026, 11, 13)

# ---------- 1) contas refeitas do zero (tabelas digitadas aqui de novo, a partir das fontes do relatório)
def tarifa(canal, preco, peso, com_ml, com_amz):
    fr3 = lambda a, b, c: a if peso <= 0.5 else b if peso <= 1 else c
    if canal == "Mercado Livre":
        return (com_ml, 6.75, 0) if preco < 79 else (com_ml, 0, fr3(20, 22, 25))
    if canal == "Shopee":
        if preco < 80: return (0.20, 4.50, 0)
        return (0.14, 16 if preco < 100 else 20 if preco < 200 else 26, 0)
    if canal == "Amazon":
        return (com_amz, 2.00, 0 if preco < 79 else fr3(18, 20, 23))
    return (0.16, 5.00 if preco >= 10 else 0, 0 if preco < 79 else fr3(15, 18, 22))


def lucro(pp, canal, preco):
    c, f, fr = tarifa(canal, preco, pp["kit_peso"], pp["categoria_ml"], pp["categoria_amz"])
    emb = 1.20 if pp["kit_peso"] <= 0.15 else 2.00 if pp["kit_peso"] <= 0.5 else 3.50
    custo = pp["kit_custo"] * 1.06 + emb
    return preco - preco * c - f - fr - preco * 0.09 - custo


PESQ = {p["id"]: p for p in DP.PRODUTOS}
MK = ["Mercado Livre", "Shopee", "Amazon", "Magalu"]
CAN = {"Mercado Livre": "Mercado Livre (Clássico)", "Shopee": "Shopee (CNPJ)", "Amazon": "Amazon (DBA)", "Magalu": "Magalu"}
wbp = openpyxl.load_workbook(os.path.join(RC, "Pesquisa_de_Mercado.xlsx"), data_only=True)
ok(not [1 for s in wbp for r in s.iter_rows() for c in r if isinstance(c.value, str) and c.value.startswith("#") and c.value[1:4] in ("REF", "VAL", "DIV", "NAM", "N/A")], "Pesquisa: erro de fórmula")
k4 = wbp["4b_Kits_Recomendados"]; linhas4 = {(r[0], r[2]): r for r in k4.iter_rows(min_row=5, values_only=True) if r[0]}
esc = {r[1]: r for r in wbp["Escolhidos"].iter_rows(min_row=5, max_row=9, values_only=True)}
tab = []
for a in PRODUTOS:
    pp = PESQ[a["id"]]
    ok(abs(pp["kit_preco"] - a["preco"]) < 0.001, f"{a['id']}: preço pesquisa {pp['kit_preco']} x anúncio {a['preco']}")
    ok(abs(esc[a["id"]][3] - a["preco"]) < 0.001, f"{a['id']}: preço aba Escolhidos {esc[a['id']][3]}")
    for mk in MK:
        l = lucro(pp, mk, a["preco"]); x = linhas4[(a["id"], CAN[mk])]
        ok(abs(x[14] - l) < 0.02, f"{a['id']} {mk}: lucro planilha {x[14]:.2f} x refeito {l:.2f}")
        ok(l > 0, f"{a['id']} {mk}: margem negativa")
        tab.append((a["id"], mk, a["preco"], round(l, 2), round(l / a["preco"] * 100, 1)))

# ---------- 2) cronograma
wbc = openpyxl.load_workbook(os.path.join(RC, "Cronograma_de_Anuncios.xlsx"), data_only=True)
ok(wbc.sheetnames[:5] == ["Geral", "Mercado_Livre", "Shopee", "Amazon", "Magalu"], f"abas do cronograma: {wbc.sheetnames}")
g = wbc["Geral"]; ok(g.freeze_panes == "A5" and g.auto_filter.ref, "cronograma sem filtro/congelamento")
rows = [r for r in g.iter_rows(min_row=5, values_only=True) if r[0]]
vistos = set(); nome2id = {a["nome"]: a["id"] for a in PRODUTOS}
pub = {}
for r in rows:
    d = r[0].date() if isinstance(r[0], dt.datetime) else r[0]
    ok(d > HOJE, f"data no passado: {r[:6]}")
    ok(DIAS[d.weekday()] == r[1], f"dia da semana errado: {d} {r[1]}")
    chave = (d, r[2]); ok(chave not in vistos, f"conflito de horário: {chave} {r[5]}"); vistos.add(chave)
    pid = nome2id.get(r[4])
    if pid and r[3] in MK:
        ok(abs(r[7] - next(a["preco"] for a in PRODUTOS if a["id"] == pid)) < 0.001, f"preço divergente no cronograma {pid} {r[3]} {r[7]}")
        if r[6]:
            ok(r[6] == next(a["titulos"][r[3]] for a in PRODUTOS if a["id"] == pid), f"título divergente no cronograma {pid} {r[3]}")
        preco_pr = r[8] or r[7]
        ok(lucro(PESQ[pid], r[3], preco_pr) > 0, f"lucro negativo no cronograma {pid} {r[3]} {preco_pr}")
        if r[5] == "Publicar anúncio":
            pub[(pid, r[3])] = d
for (pid, mk), d in pub.items():
    rec = [r for r in rows if nome2id.get(r[4]) == pid and r[5].startswith("Receber estoque")]
    ok(rec and (rec[0][0].date() if isinstance(rec[0][0], dt.datetime) else rec[0][0]) < d, f"publicação antes do estoque {pid} {mk}")
    for r in rows:
        if nome2id.get(r[4]) == pid and r[3] == mk and r[5].startswith(("Ativar promoção", "Cupom")):
            dd = r[0].date() if isinstance(r[0], dt.datetime) else r[0]
            ok(dd >= d, f"promoção antes da publicação {pid} {mk} {dd}")
ok(len(pub) == 20, f"publicações: {len(pub)} (esperado 20)")
for mk in MK:
    ok(len([1 for r in wbc[mk.replace(' ', '_')].iter_rows(min_row=5, values_only=True) if r[0]]) == len([r for r in rows if r[3] in (mk, "Todos")]), f"aba {mk} incompleta")
cal = Calendar.from_ical(open(os.path.join(RAIZ, "02_Cronograma", "Agenda_Anuncios.ics"), "rb").read())
evs = list(cal.walk("VEVENT"))
ok(len(evs) == len(rows), f"ICS {len(evs)} eventos x planilha {len(rows)}")
ok(all(len(list(e.walk("VALARM"))) == 1 and str(list(e.walk("VALARM"))[0]["TRIGGER"].to_ical(), "ascii") == "-PT30M" for e in evs), "ICS sem lembrete de 30 min")
ics_t = sorted((e["DTSTART"].dt.replace(tzinfo=None)) for e in evs)
xls_t = sorted(dt.datetime.combine(r[0].date() if isinstance(r[0], dt.datetime) else r[0], dt.time(*map(int, r[2].split(":")))) for r in rows)
ok(ics_t == xls_t, "horários do ICS diferentes da planilha")

# ---------- 3) anúncios .txt
for a in PRODUTOS:
    for mk, arq in zip(MK, ["MercadoLivre", "Shopee", "Amazon", "Magalu"]):
        f = os.path.join(RAIZ, "04_Anuncios", f"{a['id']}_{a['pasta']}_{arq}.txt"); ok(os.path.exists(f), f"falta {f}")
        t = open(f, encoding="utf-8").read()
        tit = t.split("TÍTULO (copiar e colar)")[1].strip().splitlines()[0]
        ok(len(tit) <= LIMITES[mk], f"título longo {a['id']} {mk}")
        ok(f"À vista: R$ {a['preco']:.2f}".replace(".", ",") in t, f"preço ausente/diferente no txt {a['id']} {mk}")
        for w in ("stanley", "tramontina", "zoreya", "tilonstar", "goldensun", "avental", "estanque", "necessaire + esponja", "12 pincéis"):
            ok(w not in t.lower() or (w == "12 pincéis" and "anuncia 12" in t), f"termo indevido '{w}' em {os.path.basename(f)}")
        for linha in re.findall(r"^\s+\d+\. (\S+\.jpg)", t, re.M):
            pass
        pasta = re.search(r"Pasta: (.+)", t).group(1).strip().replace("\\", os.sep)
        for im in re.findall(r"^\s+\d+\. (\S+\.jpg)", t, re.M):
            ok(os.path.exists(os.path.join(RAIZ, pasta, im)), f"imagem citada não existe: {pasta}/{im}")

# ---------- 4) imagens
for a in PRODUTOS:
    base = os.path.join(RAIZ, "03_Imagens", f"{a['id']}_{a['pasta']}")
    orig = sorted(glob.glob(os.path.join(base, "originais", "*.jpg")))
    ok(len(orig) >= 5, f"{a['id']}: só {len(orig)} originais")
    for f in orig:
        im = Image.open(f); im.verify()
        if min(Image.open(f).size) < 750:
            avisos.append(f"{a['id']}: original de baixa resolução {os.path.basename(f)} {Image.open(f).size}")
    for mk in ("mercado_livre", "shopee", "amazon", "magalu"):
        fs = sorted(glob.glob(os.path.join(base, "anuncio", mk, "*.jpg")))
        ok(len(fs) >= 3, f"{a['id']} {mk}: poucas imagens")
        for f in fs:
            im = Image.open(f); ok(im.size == (1200, 1200), f"tamanho {f} {im.size}"); ok(os.path.getsize(f) < 1_000_000, f"arquivo > 1 MB {f}")
        p = Image.open(os.path.join(base, "anuncio", mk, "01_principal.jpg")).convert("RGB")
        cantos = [p.getpixel(c) for c in [(2, 2), (1197, 2), (2, 1197), (1197, 1197), (600, 2), (2, 600)]]
        ok(all(min(c) >= 250 for c in cantos), f"{a['id']} {mk}: fundo da principal não é branco {cantos}")
        if mk == "mercado_livre":
            ok(not any(re.search(r"beneficios|o_que_vem|medidas", f) for f in fs), f"{a['id']}: arte com texto no conjunto do ML")
cred = open(os.path.join(RAIZ, "05_Fontes", "Creditos_Imagens.txt"), encoding="utf-8").read()
for a in PRODUTOS:
    ok(a["id"] in cred and "alibaba.com" in cred, f"créditos sem {a['id']}")

# ---------- 5) relatório e fontes
rel = open(os.path.join(RAIZ, "01_Pesquisa_de_Mercado", "Relatorio_Pesquisa_de_Mercado.md"), encoding="utf-8").read()
for a in PRODUTOS:
    ok(f"R$ {a['preco']:.2f}".replace(".", ",") in rel, f"preço de {a['id']} ausente no relatório")
for velho in ("R$ 64,90", "R$ 129,90", "bolsa estanque**", "+ avental**"):
    ok(velho not in rel, f"relatório ainda cita '{velho}'")
ok(len(re.findall(r"https?://", rel)) >= 20, "relatório com poucas fontes")
ok(os.path.exists(os.path.join(RAIZ, "01_Pesquisa_de_Mercado", "Pesquisa_de_Mercado.pdf")), "falta PDF")

print("MARGENS REFEITAS (preço normal):")
for t_ in tab:
    print("  ", t_)
print(f"\nERROS: {len(erros)}"); print("\n".join(erros[:40]))
print(f"AVISOS: {len(avisos)}"); print("\n".join(avisos))
json.dump(dict(erros=erros, avisos=avisos, margens=tab), open(os.path.join(AQUI, "_auditoria.json"), "w"), ensure_ascii=False, indent=1)
