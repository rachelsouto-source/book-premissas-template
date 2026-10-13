# Book de Premissas de Fachada e Interiores — Template

Template oficial dos **Books de Premissas** dos Spots Seazone (um book por empreendimento), para o time de Obras saber o que a imagem vendeu, o que o executivo especifica, onde divergem e o que validar antes da compra e da execução.

**Template atual:** [`template/TEMPLATE_v3_CAMPECHE_SPOT_Book_de_Premissas_rev13-10-2026.pptx`](template/) — v3 do Campeche Spot (Raquel Tavares, 13/10/2026).
Original no Google Slides: https://docs.google.com/presentation/d/1ErbFk7-BxDhf6TdUgVyL-WXBD-ilB3_ojECEsw0nflU/edit

O template anterior (Novo Campeche Spot III) fica em [`template/legacy/`](template/legacy/) só como histórico.

## Regras de conteúdo

**O book é um documento validado.** Ele vai para o administrador de obras assinar, então nenhum item sai como A VALIDAR, PENDENTE ou A COTAR: todo item é **DEFINIDO**, e cada divergência entre as fontes aparece já como decisão registrada.

- **Prevalecem o Briefing de Projeto e a imagem de lançamento.** O executivo deve replicar as cores e texturas do lançamento (mesma lógica do Cachoeira Spot). Onde o executivo diverge, o book indica "compatibilizar o executivo" ou "ajustar antes da compra".
- As páginas de pendências viram **"Pontos de atenção na obra"**, com o que a obra precisa observar, tudo como DEFINIDO.
- Canasvieiras: **portas e rodapés das áreas comuns PRETOS**, como na imagem de lançamento (decisão de 13/10/2026).

**Entra:** imagens externas de lançamento, materiais, cores, texturas, imagens e o comparativo **lançamento × executivo**, além dos interiores das **áreas comuns** (recepção, corredores, rooftop).

**Não entra:**
- **Interiores de unidades** (nenhuma unidade — seção removida na v3), nem em imagens soltas (capa, mosaicos, aberturas de seção)
- Índices urbanísticos
- Áreas das unidades (m²), dimensões ou reduções ("hall menor", "piscina reduzida")
- Custos / orçamento
- Atas e comunicações comerciais
- Listas de números de unidades

Nenhuma informação inventada: tudo vem das fontes do Spot (imagens de lançamento, executivo, memorial descritivo, projeto de interiores, `00 - Premissas de Projeto` — Briefing e LOG de considerações — e conferência de projetos, quando existir).

## Estrutura (38 slides na v3)

| # | Seção | Conteúdo |
|---|---|---|
| 1 | Capa | Nome do Spot · "Moodboard de premissas" · Lançamento × Executivo · Fachada · Interiores. Mosaico só com **imagens externas e de áreas comuns** — nunca imagem interna de unidade |
| 2 | Contextualização | Objetivo ("premissas de fachada e de **interiores das áreas comuns**") · Veja / Entenda / Confira · 7 perguntas que Obras deve conseguir responder (nº 5: conceito de interiores **das áreas comuns**) · regra do material e fontes (rodapé: "Interiores das unidades, índices urbanísticos, áreas e custos não fazem parte deste Book") |
| 3 | Como usar | Números na imagem · amostra circular · código do executivo · status **DEFINIDO / A VALIDAR / PENDENTE** · índice: 01 Fachada · 02 **Interiores das áreas comuns** · 03 **Materiais de interiores das áreas comuns** |
| 4–13 | **01 Fachada** | Imagens de lançamento (moodboard) · paleta de materiais · leitura geral com legenda numerada · mapa de texturas **T01–T12** (4 por slide: Onde / Especificação / Atenção na obra) · vistas com marcação e especificação (frontal, lateral, acesso) |
| 14–32 | **02 Interiores · áreas comuns** | Conceito Seazone × situação registrada · decisões de projeto (antes × agora) · revisões (ex.: nova recepção) com moodboard, itens com QTD e status (A VALIDAR / A COTAR / DEFINIDO / PENDENTE), pendências · referência de lançamento · fichas de premissas (recepção, corredores, rooftop interno e externo) |
| 33–37 | **03 Materiais de interiores** | Mapa de materialidade **T13–T28** para compra e aprovação de amostras, incluindo slide de **MDFs e cores de metal** (T25–T28) |
| 38 | Encerramento | |

## Identidade visual

Navy `#00143D` · coral `#FC6058` · Arial · barra de cabeçalho com logo · rodapé `SPOT · BOOK DE PREMISSAS DE FACHADA E INTERIORES · NN` · cards T01… · marcadores circulares numerados · amostras circulares recortadas das imagens.

## Como usar o template para um novo Spot

1. Copie o `.pptx` do template e renomeie `NOME DO SPOT - Book de Premissas de Fachada e Interiores.pptx`.
2. Troque nome do Spot, cidade, imagens e fontes; mantenha a ordem das seções.
3. Preencha texturas, vistas e fichas só com o que está nas fontes do Spot.
4. **Não** crie slides de unidades. Confira o checklist de "Não entra" antes de entregar.

## Books já migrados para o template v3

| Spot | Arquivo | Slides | Data |
|---|---|---|---|
| Campeche Spot | Google Slides (link acima) | 38 | 13/10/2026 |
| Canasvieiras Spot | [`books/CANASVIEIRAS_SPOT_Book_de_Premissas_v3.pptx`](books/) | 27 | 13/10/2026 |

## Gerador (`scripts/`)

- `tplkit.py`: clona slides do template v3 e troca textos e imagens mantendo a formatação. Recorta as imagens no tamanho exato do quadro, reposiciona os marcadores numerados e aplica os selos de status (DEFINIDO / A VALIDAR / PENDENTE / A COTAR).
- `build_canasvieiras.py`: conteúdo do Canasvieiras (base para os próximos Spots: copiar, trocar caminhos, textos e recortes).
- `render.ps1`: exporta os slides em PNG via PowerPoint para conferência visual.

```
py scripts/build_canasvieiras.py "saida.pptx"
```

Requisitos: Python com `python-pptx` e `Pillow`; as imagens de origem ficam nas pastas locais de cada Spot (não versionadas aqui).

Histórico de mudanças: [CHANGELOG.md](CHANGELOG.md).
