# -*- coding: utf-8 -*-
"""Etapa 2 — imagens.

1) Copia as fotos baixadas dos fornecedores para 03_Imagens/<produto>/originais (produto_01.jpg ...)
2) Gera 05_Fontes/Creditos_Imagens.txt
3) Monta o conjunto de anúncio em 03_Imagens/<produto>/anuncio/<marketplace>/ (1200 x 1200 px)
   - foto principal: fundo branco puro, sem texto, produto ocupando ~88% do quadro
   - Mercado Livre: só fotos sem texto
   - Shopee / Amazon / Magalu: principal + artes (benefícios, o que vem, medidas, uso) + fotos

Uso: python gerar_imagens.py <pasta_com_downloads>
     (a pasta de downloads tem uma subpasta por fornecedor, como baixado em 14/11/2026)
"""
import os, sys, json, shutil, subprocess, glob, html, datetime
import numpy as np
from PIL import Image, ImageFilter

AQUI = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, AQUI)
from dados_anuncios import PRODUTOS

RAIZ = os.path.abspath(os.path.join(AQUI, "..", ".."))
IMG = os.path.join(RAIZ, "03_Imagens")
DL = sys.argv[1] if len(sys.argv) > 1 else os.path.join(RAIZ, "05_Fontes", "_downloads_fornecedores")
TMP = os.path.join(os.environ.get("TMPDIR", "/tmp"), "pv_render"); os.makedirs(TMP, exist_ok=True)
S = 1200
CONSULTA = "14/11/2026"

# ------------------------------------------------------------------ fontes das fotos
FONTES = {
 "P07": dict(dir="toalha", url="https://www.alibaba.com/product-introduction/Customized-Eco-Friendly-Sand-Free-Solid_1601739159751.html",
             fornecedor="Fornecedor chinês no Alibaba (anúncio 1601739159751 — toalha microfibra cor sólida, sand free)",
             sel=[("01", "kit de toalhas enroladas com bolsas de tela, fundo branco"), ("02", "pilha de toalhas dobradas, várias cores"),
                  ("03", "toalhas enroladas com elástico, várias cores"), ("04", "teste de absorção (arte do fornecedor com texto 'Absorbs 4X')"),
                  ("06", "modelo segurando a toalha (arte do fornecedor com texto promocional nas bordas)")]),
 "P05": dict(dir="cart_c", url="https://www.alibaba.com/product-introduction/Bottom-Slide-Button-Ffid-Aluminum-Credit_1600162930638.html",
             fornecedor="Fornecedor chinês no Alibaba (anúncio 1600162930638 — porta-cartões de alumínio pop-up RFID)",
             sel=[("01", "carteira preta com cartões, fundo branco"), ("02", "carteira vermelha com medidas 95 x 60 x 9 mm"),
                  ("04", "carteiras prata (aberta e fechada)"), ("05", "detalhes do alumínio (arte com texto 'DETAILS')"),
                  ("06", "mão ejetando os cartões (arte com texto 'RFID BLOCKING')")]),
 "P10": dict(dir="organizador", url="https://www.alibaba.com/product-introduction/Large-Capacity-Waterproof-Polyester-Packing-Cube_1600258181886.html",
             fornecedor="Fornecedor chinês no Alibaba (anúncio 1600258181886 — kit 7 organizadores de mala, poliéster)",
             sel=[("01", "kit azul-marinho dentro da mala, fundo branco"), ("02", "kit cinza dentro da mala"),
                  ("03", "as 7 peças do kit cinza expostas"), ("05", "bolsa para sapatos"), ("06", "kit bege dentro da mala")]),
 "P06": dict(dir="bbq_g", url="https://www.alibaba.com/product-introduction/x_1600334347228.html",
             fornecedor="Fornecedor chinês no Alibaba (anúncio 1600334347228 — kit churrasco 5 peças com maleta de bambu)",
             sel=[("01", "maleta de bambu aberta com as 5 peças, fundo branco"), ("02", "maleta fechada com medidas 43 x 17,5 x 7 cm"),
                  ("03", "garfo trinchante"), ("04", "pegador"), ("05", "espátula multifuncional"), ("06", "faca")]),
 "P04": dict(dir="pinc_b", url="https://www.alibaba.com/product-introduction/Beautetools-Vegan-Luxury-High-End-Quality_1601422451195.html",
             fornecedor="Fornecedor chinês no Alibaba (anúncio 1601422451195 — kit de pincéis veganos com esponjas)",
             sel=[("01", "pincéis em pé com 3 esponjas (cena)"), ("02", "pincéis lado a lado, fundo pêssego"),
                  ("04", "pincéis com esponjas, formato retrato"), ("05", "variação rosa com esponjas, formato retrato"),
                  ("03", "detalhe dos pincéis grandes (baixa resolução, 310 px — só referência)")]),
}
# recortes (frações x0,y0,x1,y1) para tirar texto do fornecedor
RECORTE = {("P07", "06"): (0.11, 0.10, 0.85, 0.84), ("P04", "01"): (0.05, 0.0, 0.95, 0.9), ("P05", "05"): (0.0, 0.0, 0.75, 0.75), ("P05", "06"): (0.2, 0.1, 1.0, 0.9),
           ("P10", "03"): (0.0, 0.22, 1.0, 0.88)}
# papel de cada original no anúncio
PAPEL = {
 "P07": dict(principal="01", fotos=["02", "03"], uso="06", arte="01"),
 "P05": dict(principal="01", fotos=["04", "05"], uso="06", arte="01"),
 "P10": dict(principal="03", fotos=["01", "05"], uso="02", arte="03"),
 "P06": dict(principal="01", fotos=["05", "02"], uso="01", arte="01"),
 "P04": dict(principal="02", fotos=["01"], uso="01", arte="02"),
}


def achar(d, n):
    fs = glob.glob(os.path.join(DL, d, n + "_*"))
    if not fs:
        raise SystemExit(f"faltou o arquivo {d}/{n}_*")
    return fs[0]


def abrir(p):
    im = Image.open(p)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA"); bg = Image.new("RGBA", im.size, "white"); bg.alpha_composite(im); im = bg
    return im.convert("RGB")


def recorta(im, box):
    w, h = im.size
    return im.crop((int(box[0] * w), int(box[1] * h), int(box[2] * w), int(box[3] * h)))


def quadrado_cheio(im):
    """recorte central quadrado + redimensiona para 1200 (fotos de cena)."""
    w, h = im.size; m = min(w, h)
    im = im.crop(((w - m) // 2, (h - m) // 2, (w - m) // 2 + m, (h - m) // 2 + m))
    return im.resize((S, S), Image.LANCZOS).filter(ImageFilter.UnsharpMask(radius=1.2, percent=60, threshold=2))


def fundo_branco(im, tol=None, ocupacao=0.88):
    """Foto principal: fundo branco puro (255), produto centralizado ocupando ~88% do quadro."""
    a = np.asarray(im).astype(np.int32)
    borda = np.concatenate([a[0], a[-1], a[:, 0], a[:, -1]])
    bg = np.median(borda, axis=0)
    dist = np.sqrt(((a - bg) ** 2).sum(axis=2))
    claro = bg.min() > 200
    tol = tol or (28 if claro else 42)
    mask = (dist > tol).astype(np.uint8) * 255
    m = Image.fromarray(mask).filter(ImageFilter.MedianFilter(5)).filter(ImageFilter.MaxFilter(5))
    bbox = m.getbbox()
    alpha = m.filter(ImageFilter.GaussianBlur(1.5))
    branco = Image.new("RGB", im.size, (255, 255, 255))
    limpo = Image.composite(im, branco, alpha)
    obj = limpo.crop(bbox)
    ow, oh = obj.size; esc = (S * ocupacao) / max(ow, oh)
    obj = obj.resize((max(1, int(ow * esc)), max(1, int(oh * esc))), Image.LANCZOS)
    tela = Image.new("RGB", (S, S), (255, 255, 255))
    tela.paste(obj, ((S - obj.size[0]) // 2, (S - obj.size[1]) // 2))
    # garante branco puro onde já era quase branco
    t = np.asarray(tela).copy(); quase = (t > 246).all(axis=2); t[quase] = 255
    return Image.fromarray(t).filter(ImageFilter.UnsharpMask(radius=1.2, percent=50, threshold=2))


def salvar(im, p):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    im.save(p, "JPEG", quality=92, optimize=True, progressive=True)


# ------------------------------------------------------------------ artes (HTML -> PNG via Chromium)
CSS = """*{box-sizing:border-box;margin:0;padding:0}body{width:1200px;height:1200px;font-family:Inter,Arial,sans-serif;background:%(fundo)s;color:#1b1b1b;overflow:hidden}
.h1{font-family:'Inter Display',Inter,sans-serif;font-weight:800;letter-spacing:-1.5px;color:%(cor)s}
.card{background:#fff;border-radius:28px;box-shadow:0 10px 30px rgba(0,0,0,.08);overflow:hidden}
.tag{display:inline-block;background:%(cor)s;color:#fff;font-weight:700;border-radius:999px;padding:10px 22px;font-size:26px}
"""


def arte_beneficios(p, foto):
    itens = "".join(f"""<div class="card" style="padding:26px 28px;display:flex;gap:20px;align-items:flex-start;height:200px">
      <svg width="52" height="52" viewBox="0 0 24 24" style="flex:none"><circle cx="12" cy="12" r="12" fill="{p['cor']}"/><path d="M7 12.5l3.2 3.2L17.5 8.5" stroke="#fff" stroke-width="2.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>
      <div><div style="font-size:34px;font-weight:800;line-height:1.1">{html.escape(a)}</div><div style="font-size:26px;color:#4a4a4a;margin-top:10px;line-height:1.25">{html.escape(b)}</div></div></div>"""
                    for a, b in p["beneficios"])
    return f"""<div style="padding:56px 64px 0"><div class="h1" style="font-size:70px;line-height:1.02">{html.escape(p['chamada'])}</div></div>
    <div class="card" style="margin:34px 64px 0;height:520px;display:flex;align-items:center;justify-content:center"><img src="{foto}" style="width:1040px;height:500px;object-fit:contain"></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:22px;margin:24px 64px 0">{itens}</div>"""


def arte_kit(p, foto):
    lis = "".join(f"""<div style="display:flex;gap:20px;align-items:center;padding:14px 0;border-bottom:2px solid #eee">
       <div style="flex:none;width:58px;height:58px;border-radius:50%;background:{p['cor']};color:#fff;font-weight:800;font-size:28px;display:flex;align-items:center;justify-content:center">{i+1}</div>
       <div style="font-size:31px;line-height:1.2">{html.escape(k)}</div></div>""" for i, k in enumerate(p["kit"]))
    return f"""<div style="padding:64px 64px 0;display:flex;justify-content:space-between;align-items:center"><div class="h1" style="font-size:76px">O que vem no kit</div></div>
    <div style="display:flex;gap:36px;padding:40px 64px">
      <div class="card" style="flex:0 0 560px;height:560px"><img src="{foto}" style="width:560px;height:560px;object-fit:cover"></div>
      <div class="card" style="flex:1;padding:24px 34px;display:flex;flex-direction:column;justify-content:center">{lis}</div></div>
    <div style="margin:0 64px;padding:30px 36px;border-radius:24px;background:{p['cor']};color:#fff;font-size:34px;font-weight:700">{html.escape(p['nome'])}</div>
    <div style="margin:22px 64px 0;font-size:24px;color:#666">Imagens ilustrativas do produto. Cores conforme a variação escolhida.</div>"""


def arte_medidas(p, foto):
    linhas = "".join(f"""<tr><td style="padding:22px 26px;font-size:32px;color:#444;border-bottom:2px solid #eee">{html.escape(a)}</td>
      <td style="padding:22px 26px;font-size:36px;font-weight:800;text-align:right;border-bottom:2px solid #eee;color:{p['cor']}">{html.escape(b)}</td></tr>"""
                     for a, b in p["medidas"])
    return f"""<div style="padding:64px 64px 0"><div class="h1" style="font-size:80px">Medidas</div></div>
    <div class="card" style="margin:36px 64px 0;height:520px;display:flex;align-items:center;justify-content:center"><img src="{foto}" style="height:500px;width:1040px;object-fit:contain"></div>
    <div class="card" style="margin:30px 64px 0"><table style="width:100%;border-collapse:collapse">{linhas}</table></div>
    <div style="margin:22px 64px 0;font-size:24px;color:#666">Medidas aproximadas, podem variar cerca de 1 cm.</div>"""


def arte_uso(p, foto):
    return f"""<img src="{foto}" style="position:absolute;inset:0;width:1200px;height:1200px;object-fit:cover">
    <div style="position:absolute;left:0;right:0;bottom:0;height:420px;background:linear-gradient(180deg,rgba(0,0,0,0),rgba(0,0,0,.72))"></div>
    <div style="position:absolute;left:64px;right:64px;bottom:70px;color:#fff">
      <div class="tag" style="background:{p['cor2']};color:#1b1b1b">Na prática</div>
      <div style="font-family:'Inter Display',Inter;font-weight:800;font-size:68px;line-height:1.05;margin-top:22px;letter-spacing:-1px">{html.escape(p['fotos_uso'][0])}</div>
      <div style="font-size:32px;margin-top:14px;opacity:.92">{html.escape(p['chamada'])}</div></div>"""


def render(jobs):
    js = os.path.join(TMP, "render.js")
    open(js, "w").write("""const { chromium } = require('playwright');const fs=require('fs');
(async()=>{const jobs=JSON.parse(fs.readFileSync(process.argv[2]));const b=await chromium.launch();
const p=await b.newPage({viewport:{width:1200,height:1200},deviceScaleFactor:1});
for(const j of jobs){await p.goto('file://'+j.html);await p.waitForTimeout(150);await p.screenshot({path:j.out,type:'png'});}
await b.close();})();""")
    jf = os.path.join(TMP, "jobs.json"); json.dump(jobs, open(jf, "w"))
    r = subprocess.run(["node", js, jf], capture_output=True, text=True)
    if r.returncode:
        raise SystemExit(r.stderr)


# ------------------------------------------------------------------ execução
creditos = ["CRÉDITOS E ORIGEM DAS IMAGENS — Projeto Vendas",
            f"Download em {CONSULTA}. Todas as fotos são de FORNECEDORES (fabricantes/atacadistas) no Alibaba.com,",
            "obtidas das páginas públicas de produto ('product-introduction'). Nenhuma foto é de outro vendedor de marketplace.",
            "",
            "IMPORTANTE: são fotos de REFERÊNCIA do fornecedor. Antes de publicar, (1) peça ao fornecedor que você comprar",
            "autorização para usar as fotos ou as fotos dele; (2) confira se o produto recebido é igual ao das fotos.",
            "Se for diferente, fotografe o seu produto (veja 03_Imagens/GUIA_DE_FOTOS.txt) e substitua.",
            "As artes (benefícios, o que vem, medidas, uso) foram montadas sobre essas fotos; o texto é nosso.", ""]
jobs = []
for p in PRODUTOS:
    f = FONTES[p["id"]]; base = os.path.join(IMG, f"{p['id']}_{p['pasta']}")
    if os.path.isdir(base):
        shutil.rmtree(base)
    orig = os.path.join(base, "originais"); os.makedirs(orig)
    creditos += [f"=== {p['id']} — {p['nome']} ===", f"Fornecedor: {f['fornecedor']}", f"Página: {f['url']}"]
    mapa = {}
    for i, (n, desc) in enumerate(f["sel"], 1):
        src = achar(f["dir"], n); im = abrir(src)
        nome = f"{p['slug']}_{i:02d}.jpg"
        im.save(os.path.join(orig, nome), "JPEG", quality=95)
        mapa[n] = im
        creditos.append(f"  originais/{nome}  ({im.size[0]}x{im.size[1]} px)  ←  https://sc04.alicdn.com/kf/{os.path.basename(src).split('_',1)[1]}  — {desc}")
    creditos.append("")
    pap = PAPEL[p["id"]]; tmpd = os.path.join(TMP, p["id"]); os.makedirs(tmpd, exist_ok=True)

    def foto(n):
        im = mapa[n]
        if (p["id"], n) in RECORTE:
            im = recorta(im, RECORTE[(p["id"], n)])
        return im

    principal = fundo_branco(foto(pap["principal"]))
    fotos = [quadrado_cheio(foto(n)) if (p["id"], n) not in RECORTE or True else None for n in pap["fotos"]]
    uso = quadrado_cheio(foto(pap["uso"]))
    arte_src = os.path.join(tmpd, "arte_src.jpg"); _a = fundo_branco(foto(pap["arte"]), ocupacao=0.94)
    _bb = Image.fromarray(((np.asarray(_a) < 245).any(axis=2) * 255).astype(np.uint8)).getbbox()
    _a.crop((max(0, _bb[0] - 20), max(0, _bb[1] - 20), min(S, _bb[2] + 20), min(S, _bb[3] + 20))).save(arte_src, quality=95)
    uso_src = os.path.join(tmpd, "uso_src.jpg"); uso.save(uso_src, quality=95)

    # artes
    artes = {"02_beneficios": arte_beneficios(p, arte_src), "03_o_que_vem": arte_kit(p, arte_src),
             "04_medidas": arte_medidas(p, arte_src), "05_uso": arte_uso(p, uso_src)}
    for k, corpo in artes.items():
        h = os.path.join(tmpd, k + ".html")
        open(h, "w", encoding="utf-8").write(f"<!doctype html><html><head><meta charset='utf-8'><style>{CSS % p}</style></head><body style='position:relative'>{corpo}</body></html>")
        jobs.append(dict(html=h, out=os.path.join(tmpd, k + ".png"), pid=p["id"]))
    # guarda o que é fixo
    an = os.path.join(base, "anuncio"); os.makedirs(an)
    p["_tmp"] = tmpd; p["_an"] = an; p["_principal"] = principal; p["_fotos"] = fotos; p["_uso"] = uso

render(jobs)

for p in PRODUTOS:
    an, tmpd = p["_an"], p["_tmp"]
    ml = os.path.join(an, "mercado_livre")
    salvar(p["_principal"], os.path.join(ml, "01_principal.jpg"))
    for i, im in enumerate(p["_fotos"], 2):
        salvar(im, os.path.join(ml, f"0{i}_foto.jpg"))
    salvar(p["_uso"], os.path.join(ml, f"0{len(p['_fotos'])+2}_uso.jpg"))
    for mk in ("shopee", "amazon", "magalu"):
        d = os.path.join(an, mk)
        salvar(p["_principal"], os.path.join(d, "01_principal.jpg"))
        for k in ("02_beneficios", "03_o_que_vem", "04_medidas", "05_uso"):
            salvar(abrir(os.path.join(tmpd, k + ".png")), os.path.join(d, k + ".jpg"))
        for i, im in enumerate(p["_fotos"], 6):
            salvar(im, os.path.join(d, f"0{i}_foto.jpg"))

open(os.path.join(RAIZ, "05_Fontes", "Creditos_Imagens.txt"), "w", encoding="utf-8").write("\r\n".join(creditos) + "\r\n")
print("ok")
