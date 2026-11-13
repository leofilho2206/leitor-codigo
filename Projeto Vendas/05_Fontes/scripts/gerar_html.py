# -*- coding: utf-8 -*-
"""Converte o relatório .md em .html (abre no navegador, imagens embutidas)."""
import os, re, base64, markdown
D = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "01_Pesquisa_de_Mercado")
md = open(os.path.join(D, "Relatorio_Pesquisa_de_Mercado.md"), encoding="utf-8").read()
html = markdown.markdown(md, extensions=["tables"])
def emb(m):
    p = os.path.join(D, m.group(1)); b = base64.b64encode(open(p, "rb").read()).decode()
    return f'src="data:image/png;base64,{b}"'
html = re.sub(r'src="(graficos/[^"]+)"', emb, html)
html = re.sub(r'(https?://[^\s<]+)', lambda m: f'<a href="{m.group(1)}">{m.group(1)}</a>' if 'href' not in m.string[max(0,m.start()-6):m.start()] else m.group(1), html)
css = """body{font-family:Segoe UI,Arial,sans-serif;max-width:1100px;margin:24px auto;padding:0 16px;color:#1b1b1b;line-height:1.5;background:#fcfcfb}
table{border-collapse:collapse;width:100%;margin:12px 0;font-size:14px}th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
th{background:#2a78d6;color:#fff;text-align:left}tr:nth-child(even) td{background:#f4f6f9}img{max-width:100%;border:1px solid #eee;margin:8px 0}
blockquote{border-left:4px solid #eb6834;background:#fff4ec;margin:12px 0;padding:8px 14px}h1{color:#2a78d6}h2{border-bottom:2px solid #e4e3df;padding-bottom:4px;margin-top:32px}
code{background:#eee;padding:1px 4px;border-radius:3px}a{word-break:break-all}"""
out = f'<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Pesquisa de Mercado</title><style>{css}</style></head><body>{html}</body></html>'
open(os.path.join(D, "Relatorio_Pesquisa_de_Mercado.html"), "w", encoding="utf-8").write(out)
print("ok")
