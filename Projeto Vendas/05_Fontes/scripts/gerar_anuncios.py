# -*- coding: utf-8 -*-
"""Etapa 3a — anúncios prontos para copiar e colar.
Gera 04_Anuncios/<Produto>_<Marketplace>.txt (20 arquivos) + Anuncios_todos_os_canais.xlsx
e valida títulos (limite de caracteres e termos proibidos)."""
import os, sys, glob
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

AQUI = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, AQUI)
from dados_anuncios import PRODUTOS, LIMITES

RAIZ = os.path.abspath(os.path.join(AQUI, "..", ".."))
OUT = os.path.join(RAIZ, "04_Anuncios"); os.makedirs(OUT, exist_ok=True)
IMG = os.path.join(RAIZ, "03_Imagens")
PROIBIDAS = ["stanley", "tramontina", "macrilan", "secrid", "ridge", "zoreya", "tilonstar", "goldensun", "original", "réplica",
             "replica", "frete grátis", "promoção", "whatsapp", "@", "http", "melhor do brasil", "garantido", "100%", "antiareia"]
CANAIS = list(LIMITES)
PASTA_MK = {"Mercado Livre": "mercado_livre", "Shopee": "shopee", "Amazon": "amazon", "Magalu": "magalu"}
ARQ_MK = {"Mercado Livre": "MercadoLivre", "Shopee": "Shopee", "Amazon": "Amazon", "Magalu": "Magalu"}
NOTA_TIPO = {
 "Mercado Livre": "Tipo de anúncio: CLÁSSICO (comissão menor). Envio: Mercado Envios (coleta/agência); quando tiver giro, avalie o Full.",
 "Shopee": "Envio: Shopee Envios (postagem em agência/coleta). Prazo de postagem: 1 dia útil. Programa de frete grátis já incluso na comissão.",
 "Amazon": "Plano: Individual (R$ 2/item) até passar de ~10 vendas/mês; depois Profissional (R$ 19/mês). Envio: DBA (você envia).",
 "Magalu": "Envio: Magalu Entregas (coleta/postagem). Confira no Seller Portal a comissão da categoria (estimamos 16% + R$ 5,00).",
}


def brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parcelado(p, canal):
    v = p["preco"]
    base = f"À vista: {brl(v)} (Pix, boleto ou cartão em 1x)."
    if canal == "Mercado Livre":
        par = (f"Parcelado: no anúncio Clássico o comprador parcela em até 12x no cartão com juros do Mercado Pago (você recebe o valor à vista). "
               f"Sem juros só no Premium (comissão maior) — se ativar, 3x de {brl(v/3)} ou 6x de {brl(v/6)} sem juros.")
    elif canal == "Shopee":
        par = (f"Parcelado: o comprador parcela no cartão pelas condições da Shopee (até 12x, com ou sem juros conforme campanha); você recebe o valor à vista. "
               f"Referência: 3x de {brl(v/3)}.")
    elif canal == "Amazon":
        par = (f"Parcelado: a Amazon oferece o parcelamento no cartão conforme o valor e o emissor; você recebe à vista. Referência: 3x de {brl(v/3)}.")
    else:
        par = (f"Parcelado: o Magalu oferece parcelamento no cartão conforme a campanha; você recebe à vista. Referência: 3x de {brl(v/3)}.")
    return base + "\n" + par


def checar(p):
    erros = []
    for c, t in p["titulos"].items():
        if len(t) > LIMITES[c]:
            erros.append(f"{p['id']} {c}: título com {len(t)} caracteres (limite {LIMITES[c]})")
        for w in PROIBIDAS:
            if w in t.lower():
                erros.append(f"{p['id']} {c}: título contém termo proibido/arriscado '{w}'")
    return erros


def descricao(p, canal):
    L = [p["nome"].upper(), "", p["chamada"], "", p["descricao_extra"], "", "POR QUE ESCOLHER"]
    L += [f"✔ {a}: {b}" for a, b in p["beneficios"]]
    L += ["", "O QUE VEM NA EMBALAGEM"] + [f"• {k}" for k in p["kit"]]
    L += ["", "ESPECIFICAÇÕES"] + [f"• {k}: {v}" for k, v in p["ficha"].items() if k != "Marca"]
    L += ["", "MEDIDAS (aproximadas)"] + [f"• {a}: {b}" for a, b in p["medidas"]]
    L += ["", "PERGUNTAS FREQUENTES"]
    for q, r in p["faq"]:
        L += [f"› {q}", f"  {r}"]
    L.append("")
    fim = {"Mercado Livre": "Produto novo, com nota fiscal. Envio rápido após a confirmação do pagamento. Dúvidas? Use o campo de perguntas.",
           "Shopee": "Produto novo, com nota fiscal. Postagem em até 1 dia útil. Dúvidas? Chame no chat da loja.",
           "Amazon": "Produto novo, com nota fiscal.",
           "Magalu": "Produto novo, com nota fiscal. Postagem em até 1 dia útil."}[canal]
    L.append(fim)
    return "\n".join(L)


def bullets(p):
    b = [f"{a.upper()}: {d[0].upper() + d[1:]}." for a, d in p["beneficios"]]
    b.append("O QUE VEM: " + "; ".join(p["kit"]) + ".")
    return b[:5]


def imagens(p, canal):
    d = os.path.join(IMG, f"{p['id']}_{p['pasta']}", "anuncio", PASTA_MK[canal])
    fs = sorted(os.path.basename(f) for f in glob.glob(os.path.join(d, "*.jpg")))
    return d, fs


todos_erros = []
for f in glob.glob(os.path.join(OUT, "*.md")) + glob.glob(os.path.join(OUT, "*.txt")):
    if not os.path.basename(f).startswith("LEIA"):
        os.remove(f)

wb = Workbook(); ws = wb.active; ws.title = "Resumo"
H = Font(bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="2A78D6")
thin = Side(style="thin", color="D9D9D9"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WR = Alignment(wrap_text=True, vertical="top")
cab = ["ID", "Produto", "Marketplace", "Título", "Caracteres", "Limite", "Dentro do limite?", "Categoria sugerida", "Preço à vista", "Arquivo .txt", "Pasta de imagens"]
for j, c in enumerate(cab, 1):
    x = ws.cell(row=1, column=j, value=c); x.font = H; x.fill = HF; x.alignment = WR
for j, w in enumerate([6, 34, 14, 70, 10, 8, 10, 50, 12, 36, 48], 1):
    ws.column_dimensions[chr(64 + j)].width = w
wd = wb.create_sheet("Descricoes")
for j, c in enumerate(["ID", "Marketplace", "Descrição (copiar e colar)"], 1):
    x = wd.cell(row=1, column=j, value=c); x.font = H; x.fill = HF
wd.column_dimensions["A"].width = 6; wd.column_dimensions["B"].width = 14; wd.column_dimensions["C"].width = 110
r = rd = 2
for p in PRODUTOS:
    todos_erros += checar(p)
    for c in CANAIS:
        t = p["titulos"][c]; d, fs = imagens(p, c)
        if not fs:
            todos_erros.append(f"{p['id']} {c}: sem imagens em {d} (rode gerar_imagens.py antes)")
        nome_arq = f"{p['id']}_{p['pasta']}_{ARQ_MK[c]}.txt"
        rel_img = os.path.relpath(d, RAIZ).replace("/", "\\")
        txt = [f"ANÚNCIO — {p['nome']} — {c.upper()}", f"Arquivo gerado em 14/11/2026. Preço e dados conferidos com 01_Pesquisa_de_Mercado\\Pesquisa_de_Mercado.xlsx (aba Escolhidos).", "=" * 78, "",
               "TÍTULO (copiar e colar)", t, f"({len(t)} de {LIMITES[c]} caracteres permitidos)", "",
               "CATEGORIA SUGERIDA", p["categorias"][c], "(o marketplace sugere a categoria pelo título; confirme a que ele propor)", "",
               "PREÇO", parcelado(p, c), "", "TIPO DE ANÚNCIO / ENVIO", NOTA_TIPO[c], ""]
        if c in ("Amazon", "Magalu"):
            txt += ["DESTAQUES (bullet points)"] + [f"- {b}" for b in bullets(p)] + [""]
        txt += ["DESCRIÇÃO (copiar e colar)", "-" * 78, descricao(p, c), "-" * 78, "",
                "FICHA TÉCNICA / ATRIBUTOS (preencha todos — aumenta a exposição)"]
        txt += [f"  {k}: {v}" for k, v in p["ficha"].items()]
        txt += ["  Condição: Novo", "  Garantia: 90 dias (garantia legal — CDC)", "",
                "VARIAÇÕES", "  " + p["variacoes"], "",
                "PALAVRAS-CHAVE / TAGS" + (" (campo 'Termos de pesquisa')" if c == "Amazon" else ""), "  " + ", ".join(p["palavras"]), "",
                "RESPOSTAS PRONTAS PARA PERGUNTAS DOS COMPRADORES"]
        txt += [f"  P: {q}\n  R: {a_} Qualquer dúvida, estamos à disposição!" for q, a_ in p["faq"]]
        txt += ["", "IMAGENS (enviar nesta ordem)", f"  Pasta: {rel_img}"] + [f"  {i}. {f}" for i, f in enumerate(fs, 1)]
        if c == "Mercado Livre":
            txt.append("  (No Mercado Livre usamos só fotos sem texto: o ML penaliza textos e logos nas imagens.)")
        txt += ["", "ANTES DE PUBLICAR, CONFIRA NA AMOSTRA DO FORNECEDOR:"] + [f"  [ ] {s}" for s in p["spec_confirmar"]]
        txt += ["  [ ] Peso e medidas da EMBALAGEM de envio (definem o frete)", "  [ ] Foto principal igual ao produto que você recebeu", ""]
        open(os.path.join(OUT, nome_arq), "w", encoding="utf-8").write("\r\n".join(txt) + "\r\n")
        vals = [p["id"], p["nome"], c, t, len(t), LIMITES[c], "SIM" if len(t) <= LIMITES[c] else "NÃO", p["categorias"][c], p["preco"], nome_arq, rel_img]
        for j, v in enumerate(vals, 1):
            x = ws.cell(row=r, column=j, value=v); x.border = BOX; x.alignment = WR
        ws.cell(row=r, column=9).number_format = 'R$ #,##0.00'
        r += 1
        wd.cell(row=rd, column=1, value=p["id"]); wd.cell(row=rd, column=2, value=c)
        x = wd.cell(row=rd, column=3, value=descricao(p, c)); x.alignment = WR; wd.row_dimensions[rd].height = 409
        rd += 1
ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:K{r-1}"; wd.freeze_panes = "A2"
wb.save(os.path.join(OUT, "Anuncios_todos_os_canais.xlsx"))
if todos_erros:
    print("PROBLEMAS:"); print("\n".join(todos_erros)); sys.exit(1)
print("ok:", len(glob.glob(os.path.join(OUT, "*.txt"))), "arquivos .txt")
