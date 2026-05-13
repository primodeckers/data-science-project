# -*- coding: utf-8 -*-
"""
Gera os PNG referenciados em `relatorio_investimentos.md` (pasta `figures/`).

Os gráficos reproduzem a lógica do notebook `exploracao_investimentos.ipynb`
(mesmos agregados). Requer `data/investimentos_2021_2025.parquet` — corra antes
`python scripts/carregar_investimentos.py` se necessário.

Uso (na raiz do projeto):
    python scripts/gerar_figuras_relatorio.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

ROOT = Path(__file__).resolve().parents[1]
PARQUET = ROOT / "data" / "investimentos_2021_2025.parquet"
OUT_DIR = ROOT / "figures"


def _fig01_evolucao(df: pd.DataFrame, path: Path) -> None:
    sns.set_theme(style="whitegrid", context="notebook", font_scale=1.05)
    by_year = df.groupby("ano")["valor_reais"].sum() / 1e9
    yoy = (by_year.pct_change() * 100).dropna()
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.6))
    ax0 = axes[0]
    x = by_year.index.astype(str)
    ax0.bar(x, by_year.values, color="#1f4e79", edgecolor="white", linewidth=0.7)
    ax0.set_title("Total pago por ano (GND 4 e 5, governo central)")
    ax0.set_xlabel("Ano")
    ax0.set_ylabel("R$ bilhões")
    ymax = by_year.max()
    for xi, yi in zip(x, by_year.values):
        ax0.text(xi, yi + ymax * 0.02, f"{yi:.1f}", ha="center", va="bottom", fontsize=9, color="#222")
    ax1 = axes[1]
    x1 = yoy.index.astype(str)
    ax1.bar(x1, yoy.values, color="#c0392b", edgecolor="white", linewidth=0.7)
    ax1.axhline(0, color="#333", linewidth=0.9)
    ax1.set_title("Variação % em relação ao ano anterior")
    ax1.set_xlabel("Ano")
    ax1.set_ylabel("Variação (%)")
    plt.tight_layout()
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _fig02_composicao(df: pd.DataFrame, path: Path) -> None:
    _pivot_g = df.pivot_table(
        index="ano", columns="grupo_despesa_desc", values="valor_reais", aggfunc="sum"
    )
    pct = (_pivot_g.div(_pivot_g.sum(axis=1), axis=0) * 100).sort_index()
    cols = list(pct.columns)
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    bottom = np.zeros(len(pct))
    x = pct.index.astype(str)
    pal = sns.color_palette("Set2", n_colors=max(len(cols), 3))
    for i, c in enumerate(cols):
        lab = str(c)[:48] + ("…" if len(str(c)) > 48 else "")
        vals = pct[c].values
        ax.bar(
            x,
            vals,
            bottom=bottom,
            label=lab,
            color=pal[i % len(pal)],
            edgecolor="white",
            linewidth=0.5,
        )
        bottom = bottom + vals
    ax.set_ylim(0, 100)
    ax.set_ylabel("Participação no ano (%)")
    ax.set_xlabel("Ano")
    ax.set_title("Investimentos vs inversões financeiras — share dentro de cada exercício")
    ax.legend(title="Grupo de despesa", bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=8)
    plt.tight_layout()
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _fig03_orgaos_sazonalidade(df: pd.DataFrame, path: Path) -> None:
    MES_ORD = [
        "JANEIRO", "FEVEREIRO", "MARCO", "ABRIL", "MAIO", "JUNHO",
        "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
    ]
    mcat = pd.Categorical(df["mes"], categories=MES_ORD, ordered=True)
    mensal = df.assign(_mes=mcat).groupby("_mes", observed=True)["valor_reais"].sum() / 1e9
    top12 = df.groupby("orgao_maximo_desc")["valor_reais"].sum().nlargest(12) / 1e9
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
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _fig04_regiao_ano(df: pd.DataFrame, path: Path) -> None:
    reg_excl = ["SEM INFORMACAO", "CODIGO INVALIDO", "NACIONAL", "EXTERIOR"]
    reg_ok = df[~df["regiao"].isin(reg_excl)].copy()
    reg_heat = reg_ok.pivot_table(index="ano", columns="regiao", values="valor_reais", aggfunc="sum") / 1e9
    fig, axh = plt.subplots(figsize=(7.5, 4.2))
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
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _fig05_correlacao(df: pd.DataFrame, path: Path) -> None:
    reg_excl = ["SEM INFORMACAO", "CODIGO INVALIDO", "NACIONAL", "EXTERIOR"]
    reg_df = df[~df["regiao"].isin(reg_excl)].copy()
    pivot_reg = reg_df.pivot_table(
        index="ano", columns="regiao", values="valor_reais", aggfunc="sum"
    ) / 1e6
    corr_reg = pivot_reg.corr().round(3)
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
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    if not PARQUET.is_file():
        print("Falta o Parquet:", PARQUET, file=sys.stderr)
        print("Execute: python scripts/carregar_investimentos.py", file=sys.stderr)
        sys.exit(1)
    df = pd.read_parquet(PARQUET)
    df["ano"] = df["ano"].astype(str)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    jobs = [
        ("fig01_evolucao.png", _fig01_evolucao),
        ("fig02_composicao.png", _fig02_composicao),
        ("fig03_orgaos_sazonalidade.png", _fig03_orgaos_sazonalidade),
        ("fig04_regiao_ano.png", _fig04_regiao_ano),
        ("fig05_correlacao_regioes.png", _fig05_correlacao),
    ]
    for name, fn in jobs:
        out = OUT_DIR / name
        fn(df, out)
        print("Gravado:", out.resolve())


if __name__ == "__main__":
    main()
