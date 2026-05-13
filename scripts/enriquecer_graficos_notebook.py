"""Insere markdown de interpretação e atualiza código comentado nos gráficos."""
from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path


def to_src(s: str) -> list[str]:
    return [line + "\n" for line in s.splitlines()]


def md_cell(text: str) -> dict:
    lines = text.strip().split("\n")
    src = [ln + "\n" for ln in lines[:-1]]
    if lines:
        src.append(lines[-1])
    return {"cell_type": "markdown", "metadata": {}, "source": src}


def patch_code_cell(c: dict, new_source: str) -> dict:
    o = deepcopy(c)
    o["source"] = to_src(new_source)
    o["outputs"] = []
    o["execution_count"] = None
    return o


def src(c: dict) -> str:
    return "".join(c.get("source", []))


MD_CORR = r"""
#### Gráfico: correlação entre regiões (heatmap)

**O que é:** cada célula colorida é o coeficiente de **Pearson** entre duas regiões, calculado sobre a **série de totais anuais** (um ponto por ano) de pagamentos GND 4+5, em **R$ milhões**. A tabela numérica acima do gráfico é a mesma matriz.

**Como ler:** tons **avermelhados** (positivos) indicam que, ao longo de 2021–2025, quando o total de uma região sobe em relação à média dos anos, o da outra tende a subir junto; tons **azulados** (negativos) indicam movimento **contrário** entre as séries.

**Interpretação (sem causalidade):** isso mede apenas **co-movimento no tempo** — por exemplo, ciclos macro do orçamento federal, RP ou choques em um ano que afetam várias regiões na mesma direção. **Não** significa que uma região “cause” o fluxo da outra.

**Por que excluímos algumas categorias:** rótulos como `SEM INFORMACAO`, `NACIONAL` ou `CODIGO INVALIDO` geram séries estranhas e correlações espúrias (às vezes ±1). O gráfico usa só **macro-regiões** com geografia mais consistente.

**Limite estatístico:** são **poucos anos**; um exercício atípico muda muito o coeficiente. Use como apoio exploratório no relatório, não como prova definitiva.
"""

MD_YOY = r"""
#### Gráficos 1 e 2: total pago por ano e variação % (YoY)

**Gráfico à esquerda (barras azuis):** para cada ano, a altura é a **soma de todo o caixa** pago naquele exercício, em **bilhões de reais**, incluindo GND 4 e 5. Os **rótulos** em cima das barras servem para citar valores exatos no PDF.

**Gráfico à direita (barras vermelhas):** mostra a **variação percentual** do total em relação ao **ano imediatamente anterior** (YoY). A linha **zero** separa anos em que o agregado **cresceu** ou **diminuiu** em relação ao ano anterior.

**Interpretação:** quedas ou saltos fortes podem refletir **calendário de pagamentos**, **restos a pagar**, mudança na **composição** (mais grupo 4 ou 5) ou efeitos de **reclassificação** — sempre vale cruzar com notas metodológicas e contexto orçamentário, não só com “prioridade política”.

**O que não concluir daqui:** isso **não** mede qualidade de gasto, entrega de obra nem impacto social — apenas **volume de pagamentos** neste recorte.
"""

MD_STACK = r"""
#### Gráfico: composição % — grupo 4 (investimentos) vs grupo 5 (inversões financeiras) por ano

**O que é:** gráfico de barras **empilhadas a 100%**. Em cada ano, as duas fatias somam **100%** do total pago naquele exercício. Mostra a **receita relativa** de cada natureza **dentro** do ano, não o volume em bilhões (para absoluto, use o gráfico de totais anuais).

**Como interpretar:** se a fatia do **grupo 5** cresce em determinado ano, aquele exercício ficou **proporcionalmente** mais carregado em pagamentos classificados como **inversão financeira**; se a fatia do **grupo 4** domina, o ano foi mais puxado por **investimento** no sentido orçamentário (bens, obras, equipamentos etc., conforme cadastro).

**Cuidado:** “investimento orçamentário” **não** é o mesmo que “investimento em desenvolvimento econômico” do jornal; os nomes vêm do **GND** e do manual do planejamento.

**Legenda:** textos longos vêm da descrição oficial da despesa; no relatório você pode encurtar para “GND 4” e “GND 5” se definir no texto.
"""

MD_ORG = r"""
#### Gráficos seguintes: órgãos, sazonalidade mensal e mapa região × ano

**Primeira figura — painel de cima (12 barras horizontais):** cada barra é a **soma 2021–2025** dos pagamentos para aquele **órgao máximo**. Mede **concentração de fluxo** no período: quem “pesa” mais no caixa agregado. Órgãos com grandes programas de infraestrutura ou equipamentos tendem a aparecer no topo — isso é compatível com a **missão** do órgão, mas **não prova** por si só eficiência ou favoritismo.

**Primeira figura — painel de baixo (12 barras verticais por mês):** soma o caixa de **todos os anos** em cada **mês civil** (janeiro a dezembro). **Picos em dezembro** são comuns no setor público (fechamento de exercício, liquidações). Use para descrever **ritmo de execução**, não “preferência por mês” isolada.

**Segunda figura (mapa de calor):** cada célula é o total **naquele ano** e **naquela região**, em **bilhões**; quanto mais **escuro o azul**, maior o valor. Regiões problemáticas no cadastro foram **excluídas** para leitura mais limpa. Compare **ao longo da linha** (mesma região, anos diferentes) para ver mudança do caixa **neste indicador**.

**Limitações:** totais regionais ainda podem sofrer com **lacunas geográficas** nas linhas micro; evite conclusões territoriais fortes sem cruzar com população, empenho ou outras fontes.
"""

CORR = r'''# =============================================================================
# GRÁFICO: correlação entre regiões (tabela + heatmap)
# - Cada valor Pearson compara a série de TOTAIS ANUAIS (2021–2025) entre duas regiões.
# - Interpretação: co-movimento no tempo; NÃO implica causalidade nem "fluxo entre UFs".
# =============================================================================
# Correlação entre regiões (totais anuais em milhões) — só macro-regiões válidas
# (exclui SEM INFORMACAO, NACIONAL etc., que geram correlações espúrias de ±1)
import matplotlib.pyplot as plt
import seaborn as sns

# Regiões excluídas: cadastro "NACIONAL"/"SEM INFO" não representa território homogêneo
reg_excl = ["SEM INFORMACAO", "CODIGO INVALIDO", "NACIONAL", "EXTERIOR"]
reg_df = df[~df["regiao"].isin(reg_excl)].copy()

# Linhas = anos, colunas = região; valores em milhões de reais
pivot_reg = reg_df.pivot_table(
    index="ano", columns="regiao", values="valor_reais", aggfunc="sum"
) / 1e6

# Matriz simétrica: diagonal = 1 (cada região com ela mesma)
corr_reg = pivot_reg.corr().round(3)
display(corr_reg)

# Heatmap: divergindo em zero; anotações mostram o coeficiente em cada célula
fig, ax = plt.subplots(figsize=(6.5, 5.2))
sns.heatmap(
    corr_reg,
    annot=True,
    fmt=".2f",
    cmap="RdBu_r",
    center=0,
    vmin=-1,
    vmax=1,
    square=True,
    linewidths=0.35,
    ax=ax,
    cbar_kws={"shrink": 0.75, "label": "Pearson"},
)
ax.set_title(
    "Co-movimento entre regiões (totais anuais, R$ milhões)\n"
    "Correlação no tempo ≠ relação causal entre regiões"
)
plt.tight_layout()
plt.show()
'''

YOY = r'''# =============================================================================
# GRÁFICOS 1 e 2: total anual (bilhões) e variação % ano a ano (YoY)
# - Esquerda: NÍVEL do caixa agregado por exercício.
# - Direita: CRESCIMENTO ou QUEDA percentual vs. ano anterior (primeiro ano sem YoY).
# =============================================================================
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

sns.set_theme(style="whitegrid", context="notebook", font_scale=1.05)

# Série anual em bilhões de reais (soma de todas as linhas daquele ano)
by_year = df.groupby("ano")["valor_reais"].sum() / 1e9
# Variação percentual; primeiro ano vira NaN e é removido no gráfico da direita
yoy = (by_year.pct_change() * 100).dropna()

fig, axes = plt.subplots(1, 2, figsize=(12, 4.6))

# --- Painel esquerdo: barras com rótulo do valor exato (bilhões) ---
ax0 = axes[0]
x = by_year.index.astype(str)
ax0.bar(x, by_year.values, color="#1f4e79", edgecolor="white", linewidth=0.7)
ax0.set_title("Total pago por ano (GND 4 e 5, governo central)")
ax0.set_xlabel("Ano")
ax0.set_ylabel("R$ bilhões")
ymax = by_year.max()
for xi, yi in zip(x, by_year.values):
    ax0.text(xi, yi + ymax * 0.02, f"{yi:.1f}", ha="center", va="bottom", fontsize=9, color="#222")

# --- Painel direito: YoY; linha em 0 separa crescimento de queda ---
ax1 = axes[1]
x1 = yoy.index.astype(str)
ax1.bar(x1, yoy.values, color="#c0392b", edgecolor="white", linewidth=0.7)
ax1.axhline(0, color="#333", linewidth=0.9)
ax1.set_title("Variação % em relação ao ano anterior")
ax1.set_xlabel("Ano")
ax1.set_ylabel("Variação (%)")

plt.tight_layout()
plt.show()
'''

STACK = r'''# =============================================================================
# GRÁFICO: composição % dentro de cada ano (GND 4 vs GND 5)
# - Cada coluna empilhada soma 100%: mostra a MISTURA relativa, não o volume absoluto.
# =============================================================================
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Composição grupo 4 vs 5 dentro de cada ano (empilhado 100%)
_pivot_g = df.pivot_table(
    index="ano", columns="grupo_despesa_desc", values="valor_reais", aggfunc="sum"
)
# Percentual do total de cada ano pertencente a cada coluna de grupo_despesa_desc
pct = (_pivot_g.div(_pivot_g.sum(axis=1), axis=0) * 100).sort_index()
cols = list(pct.columns)

fig, ax = plt.subplots(figsize=(8.5, 4.8))
bottom = np.zeros(len(pct))
x = pct.index.astype(str)
pal = sns.color_palette("Set2", n_colors=max(len(cols), 3))
for i, c in enumerate(cols):
    lab = str(c)[:48] + ("…" if len(str(c)) > 48 else "")
    vals = pct[c].values
    ax.bar(x, vals, bottom=bottom, label=lab, color=pal[i % len(pal)], edgecolor="white", linewidth=0.5)
    bottom = bottom + vals
ax.set_ylim(0, 100)
ax.set_ylabel("Participação no ano (%)")
ax.set_xlabel("Ano")
ax.set_title("Investimentos vs inversões financeiras — share dentro de cada exercício")
ax.legend(title="Grupo de despesa", bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=8)
plt.tight_layout()
plt.show()
'''

ORG = r'''# =============================================================================
# GRÁFICOS: (A) top 12 órgãos + sazonalidade mensal  (B) heatmap região × ano
# - (A) mede concentração institucional e padrão de mês (ex.: dezembro).
# - (B) intensidade = bilhões; mesmas exclusões geográficas que na correlação.
# =============================================================================
# Maiores órgãos + sazonalidade mensal + mapa região × ano (dados filtrados)
import matplotlib.pyplot as plt
import seaborn as sns

MES_ORD = [
    "JANEIRO", "FEVEREIRO", "MARCO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
]
# Ordena meses corretamente (ordem alfabética quebraria jan/dez)
mcat = pd.Categorical(df["mes"], categories=MES_ORD, ordered=True)
mensal = df.assign(_mes=mcat).groupby("_mes", observed=True)["valor_reais"].sum() / 1e9

# Soma 2021–2025 por órgão máximo; depois os 12 maiores
top12 = df.groupby("orgao_maximo_desc")["valor_reais"].sum().nlargest(12) / 1e9

reg_excl = ["SEM INFORMACAO", "CODIGO INVALIDO", "NACIONAL", "EXTERIOR"]
reg_ok = df[~df["regiao"].isin(reg_excl)].copy()
# Matriz para heatmap: linha = ano, coluna = região; valores em bilhões
reg_heat = reg_ok.pivot_table(index="ano", columns="regiao", values="valor_reais", aggfunc="sum") / 1e9

# --- Figura A: órgãos (horizontal) + perfil mensal (vertical) ---
fig, axes = plt.subplots(2, 1, figsize=(11, 8.5))

axt = axes[0]
top12.sort_values().plot(kind="barh", ax=axt, color="#1b4332", width=0.78)
axt.set_xlabel("R$ bilhões (soma 2021–2025)")
axt.set_title("12 maiores órgãos (campo órgão máximo)")

axm = axes[1]
axm.bar(range(len(mensal)), mensal.values, color="#457b9d", edgecolor="white", linewidth=0.4)
axm.set_xticks(range(len(mensal)))
axm.set_xticklabels([str(i)[:3].title() for i in mensal.index], rotation=35, ha="right")
axm.set_ylabel("R$ bilhões")
axm.set_title("Sazonalidade: soma de caixa por mês (todos os anos)")

plt.tight_layout()
plt.show()

# --- Figura B: região no eixo Y (transposta) para leitura tipo mapa ---
fig2, axh = plt.subplots(figsize=(7.5, 4.2))
sns.heatmap(
    reg_heat.T,
    annot=True,
    fmt=".1f",
    cmap="Blues",
    linewidths=0.4,
    ax=axh,
    cbar_kws={"label": "R$ bilhões"},
)
axh.set_title("Pagamentos por região e ano (excl. SEM INFO / NACIONAL / EXTERIOR)")
axh.set_xlabel("Ano")
axh.set_ylabel("Região")
plt.tight_layout()
plt.show()
'''


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    p = root / "notebooks" / "exploracao_investimentos.ipynb"
    nb = json.loads(p.read_text(encoding="utf-8"))
    cells = nb["cells"]

    idx_corr = idx_story = idx_yoy = idx_stack = idx_org = None
    for i, c in enumerate(cells):
        s = src(c)
        if c["cell_type"] == "code" and "Correlação entre regiões" in s and "pivot_reg" in s:
            idx_corr = i
        if c["cell_type"] == "markdown" and "## 4) Storytelling" in s:
            idx_story = i
        if c["cell_type"] == "code" and "sns.set_theme(style=\"whitegrid\"" in s and "by_year = df.groupby" in s:
            idx_yoy = i
        if c["cell_type"] == "code" and s.startswith("# Composição grupo 4 vs 5"):
            idx_stack = i
        if c["cell_type"] == "code" and s.startswith("# Maiores órgãos + sazonalidade"):
            idx_org = i

    if None in (idx_corr, idx_story, idx_yoy, idx_stack, idx_org):
        raise SystemExit(f"Índices não encontrados: {idx_corr=} {idx_story=} {idx_yoy=} {idx_stack=} {idx_org=}")

    out: list[dict] = []
    for i, c in enumerate(cells):
        if i == idx_corr:
            out.append(md_cell(MD_CORR))
            out.append(patch_code_cell(c, CORR))
            continue
        if i == idx_story:
            out.append(c)
            out.append(md_cell(MD_YOY))
            continue
        if i == idx_yoy:
            out.append(patch_code_cell(c, YOY))
            continue
        if i == idx_stack:
            out.append(md_cell(MD_STACK))
            out.append(patch_code_cell(c, STACK))
            continue
        if i == idx_org:
            out.append(md_cell(MD_ORG))
            out.append(patch_code_cell(c, ORG))
            continue
        out.append(c)

    nb["cells"] = out
    p.write_text(json.dumps(nb, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Notebook atualizado:", p)


if __name__ == "__main__":
    main()
