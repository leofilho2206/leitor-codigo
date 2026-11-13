# PROGRESSO — Projeto Vendas

Última atualização: **13/11/2026** (fim da Sessão 1)

## Sessão 1 — Etapas 0 e 1 ✅ CONCLUÍDA

### Etapa 0 — Pasta do projeto ✅
- [x] Estrutura criada: `00_LEIA-ME.txt`, `00_PROGRESSO.md`, `01_Pesquisa_de_Mercado`, `02_Cronograma`, `03_Imagens`, `04_Anuncios`, `05_Fontes`
- [x] Script `Copiar_para_Area_de_Trabalho.bat` / `.ps1`: localiza a Área de Trabalho real com `[Environment]::GetFolderPath('Desktop')` (inclusive dentro do OneDrive) e copia a pasta para lá
- ⚠️ A sessão rodou na nuvem (Linux), sem acesso ao seu PC Windows. A pasta foi gerada no repositório `leofilho2206/leitor-codigo`, branch `claude/business-10-days-eo2w5h`. **Para levar ao PC:** baixe o branch (GitHub → Code → Download ZIP), extraia e dê dois cliques em `Projeto Vendas\Copiar_para_Area_de_Trabalho.bat`

### Etapa 1 — Pesquisa de mercado ✅
- [x] 12 candidatos levantados (tendências 2026 e rankings de mais vendidos)
- [x] Google Trends Brasil, 5 anos e 12 meses, normalizado (`05_Fontes/google_trends/`)
- [x] 102 anúncios registrados (ML e Shopee) com preço, nota, avaliações, frete grátis, Full e selos
- [x] Tarifas dos 4 marketplaces (ML, Shopee, Amazon, Magalu) em planilha editável
- [x] Impostos de importação: Receita Federal (página oficial salva) e ICMS-SC 17% (PDF oficial do Comsefaz salvo)
- [x] Custo total por unidade e lucro por canal, para produto avulso e em kit (fórmulas conferidas no LibreOffice: 0 erros)
- [x] Comparação atacado nacional x importação
- [x] Riscos legais (INMETRO, Anatel, Anvisa, etiqueta têxtil, marca) e logística/devolução
- [x] Notas de 0 a 10 e ranking ponderado
- [x] Relatório final (`.md` + `.html` com gráficos) e 3 gráficos
- [x] Revisão final feita

**Produtos escolhidos (formato kit/presente):** toalha de praia microfibra; carteira slim RFID; kit organizador de mala; kit churrasco com maleta; kit pincéis de maquiagem.
**Reservas:** mini ventilador; copo térmico.

### Limitações registradas (ver relatório, seção 2)
- ML e Shopee bloqueiam a busca automática sem login, e Amazon e Magalu ficaram inacessíveis no ambiente. Os anúncios vieram de resultados indexados e agregadores. Foram 5 a 11 anúncios por produto no ML e 0 a 7 na Shopee: **a meta de 10 + 10 não foi atingida**. Para completar, use a aba `8_Validar_10_anuncios`.
- Custos de atacado, frete pago pelo vendedor e algumas tarifas estão marcados como **ESTIMADO** (células amarelas).

## Sessão 2 — Etapas 2, 3 e 4 ✅ CONCLUÍDA (14/11/2026)
- [x] Sessão 1 completada: aba **Escolhidos** na planilha e **Pesquisa_de_Mercado.pdf**
- [x] Etapa 2: fotos de fornecedores (Alibaba) em `03_Imagens\<produto>\originais` + créditos em `05_Fontes\Creditos_Imagens.txt`
- [x] Etapa 2: conjuntos de anúncio por marketplace (1200 px; ML sem texto; Shopee/Amazon/Magalu com artes)
- [x] Etapa 3: 20 anúncios .txt + planilha consolidada
- [x] Etapa 3: cronograma (.xlsx, .ics, .html) — 106 ações de 14/11/2026 a 10/01/2027
- [x] Etapa 4: revisão final em duas passadas (0 erros) + 00_LEIA-ME.txt — detalhes em `05_Fontes\Relatorio_de_Revisao.txt`

### Decisões tomadas sem perguntar (anotadas)
- Kits ajustados às fotos dos fornecedores: toalha com bolsa de tela (R$ 59,90); churrasco com maleta de bambu e sem avental (R$ 119,90); pincéis 13 + 3 esponjas; carteira de alumínio. A pesquisa foi recalculada com isso.
- Fotos: só de fornecedores (Alibaba). A busca do Alibaba pede verificação humana; usamos as páginas públicas de produto. Fotos com logo/marca d'água foram descartadas.
- Prazo de fornecedor ESTIMADO: 2 dias úteis (toalha, SC) e 5 dias úteis (demais, SP), com o feriado de 20/11.
- Conta nova não entra nas campanhas oficiais de Black Friday (ML exige reputação; Shopee é seletiva): o cronograma usa promoções da própria loja (10% ML/Shopee, 8% Amazon/Magalu).
- Ambiente: a sessão roda na nuvem, sem acesso ao seu PC. Use `Copiar_para_Area_de_Trabalho.bat` depois de baixar o branch.

## O que falta (com você)
- [ ] Cotar os 5 produtos (sábado 14/11) e fechar o pedido na segunda 16/11 — ver cronograma
- [ ] Itens da lista "O QUE VOCÊ PRECISA CONFIRMAR" no `00_LEIA-ME.txt`
- [ ] Depois de cotar, atualizar custos em `05_Fontes\scripts\dados_pesquisa.py` (kit_custo) e rodar os scripts de novo (ou editar a aba 4b da planilha)
