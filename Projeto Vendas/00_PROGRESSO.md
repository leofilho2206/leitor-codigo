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

## Para você fazer antes da Sessão 2
- [ ] Cotar os 5 produtos com 3 fornecedores cada (aba `9_Confirmar`)
- [ ] Completar a validação de 10 anúncios ML + 10 Shopee (aba 8)
- [ ] Decidir MEI ou ME e conferir as tarifas no painel de vendedor
- [ ] (Opcional) Na próxima sessão, rodar no seu PC (Claude Code no Windows ou Remote Control) para gravar direto na Área de Trabalho e acessar os marketplaces logado

## Sessão 2 — A FAZER
- [ ] Etapa 2: imagens dos produtos (`03_Imagens`)
- [ ] Etapa 3: anúncios — títulos, descrições e ficha técnica por marketplace (`04_Anuncios`)
- [ ] Etapa 4: cronograma até o Natal (Black Friday 27/11, Cyber Monday 30/11, 13º em 30/11 e 20/12) (`02_Cronograma`)
- [ ] Atualizar a planilha com os custos reais cotados
