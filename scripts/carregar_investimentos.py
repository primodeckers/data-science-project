"""
Carrega e unifica investimentos federais (2021–2025): CSVs + XLSX 2025.
Executa validação inicial para iniciar a análise exploratória.

Uso (na raiz do projeto):
    python scripts/carregar_investimentos.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]

# Mesma ordem dos arquivos CSV metadados / export SIAFI (27 colunas).
COLUNAS_CANONICAS = [
    "ano",
    "mes",
    "esfera_orcamentaria",
    "esfera_orcamentaria_desc",
    "orgao_maximo",
    "orgao_maximo_desc",
    "uo",
    "uo_desc",
    "grupo_despesa",
    "grupo_despesa_desc",
    "aplicacao",
    "aplicacao_desc",
    "resultado",
    "resultado_desc",
    "funcao",
    "funcao_desc",
    "subfuncao",
    "subfuncao_desc",
    "programa",
    "programa_desc",
    "acao",
    "acao_desc",
    "regiao",
    "uf",
    "uf_desc",
    "municipio",
    "movimento_liquido_reais",
]


def parse_movimento_br(raw: object) -> float:
    """Converte movimento_liquido_reais (formato BR, negativos entre parênteses)."""
    if pd.isna(raw):
        return float("nan")
    # Excel já entrega float/int — não aplicar regra BR (remover '.' quebra decimais).
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return float(raw)
    s = str(raw).strip().strip('"').strip("'")
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        neg = True
        s = s[1:].strip()
    if not s or s.upper() in {"NAN", "NONE", ""}:
        return float("nan")
    # Remove separador de milhar (.), decimal (,)
    s = s.replace(".", "").replace(",", ".")
    try:
        v = float(s)
    except ValueError:
        return float("nan")
    return -v if neg else v


def load_csv(path: Path) -> pd.DataFrame:
    for enc in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            df = pd.read_csv(path, sep=";", encoding=enc, low_memory=False)
            break
        except UnicodeDecodeError:
            continue
    else:  # pragma: no cover
        raise RuntimeError(f"Encoding desconhecido: {path}")
    if "movimento_liquido_reais" not in df.columns and len(df.columns) == len(COLUNAS_CANONICAS):
        df.columns = COLUNAS_CANONICAS
    df["valor_reais"] = df["movimento_liquido_reais"].map(parse_movimento_br)
    src = path.stem
    df["_fonte_arquivo"] = src
    return df


def load_xlsx_2025(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    if "movimento_liquido_reais" not in df.columns and len(df.columns) == len(COLUNAS_CANONICAS):
        df.columns = COLUNAS_CANONICAS
    if "movimento_liquido_reais" not in df.columns:
        raise ValueError(f"Coluna movimento_liquido_reais não encontrada em {path}")
    df["valor_reais"] = df["movimento_liquido_reais"].map(parse_movimento_br)
    df["_fonte_arquivo"] = path.stem
    return df


def carregar_unificado(base: Path | None = None, *, verbose: bool = True) -> pd.DataFrame:
    """Une todos os CSV + XLSX 2025 e devolve um DataFrame com `valor_reais`."""
    raiz = base or ROOT
    csv_paths = sorted(raiz.glob("investimentos_20*.csv"))
    xlsx_2025 = raiz / "investimentos-2025.xlsx"

    frames: list[pd.DataFrame] = []
    for p in csv_paths:
        frames.append(load_csv(p))
        if verbose:
            print(f"OK CSV  {p.name}: {len(frames[-1]):,} linhas")

    if xlsx_2025.exists():
        frames.append(load_xlsx_2025(xlsx_2025))
        if verbose:
            print(f"OK XLSX {xlsx_2025.name}: {len(frames[-1]):,} linhas")
    elif verbose:
        print(f"Aviso: {xlsx_2025.name} não encontrado — série sem 2025.", file=sys.stderr)

    if not frames:
        raise FileNotFoundError(
            f"Nenhum investimentos_*.csv em {raiz}. Coloque os arquivos na raiz do projeto."
        )

    return pd.concat(frames, ignore_index=True)


def salvar_parquet(df: pd.DataFrame, base: Path | None = None, *, verbose: bool = True) -> Path | None:
    raiz = base or ROOT
    out_parquet = raiz / "data" / "investimentos_2021_2025.parquet"
    out_parquet.parent.mkdir(parents=True, exist_ok=True)
    export = df.copy()
    export["movimento_liquido_reais"] = export["movimento_liquido_reais"].map(
        lambda x: "" if pd.isna(x) else str(x)
    )
    try:
        export.to_parquet(out_parquet, index=False)
        if verbose:
            print(f"\nSalvo: {out_parquet}")
        return out_parquet
    except Exception as e:
        if verbose:
            print(f"\nParquet não gravado ({e}). pip install pyarrow")
        return None


def main() -> None:
    df = carregar_unificado(ROOT, verbose=True)

    # Resumo
    print("\n--- Base unificada ---")
    print(f"Linhas totais: {len(df):,}")
    print(f"Colunas: {len(df.columns)}")
    if "ano" in df.columns:
        anos = sorted(df["ano"].astype(str).unique())
        print(f"Anos: {anos}")

    total = df["valor_reais"].sum()
    print(f"Soma valor_reais (R$): {total:,.2f}")

    n_missing = df["valor_reais"].isna().sum()
    if n_missing:
        print(f"Atenção: {n_missing:,} linhas com valor não numérico após parse.")

    # Por ano
    if "ano" in df.columns:
        print("\n--- Total por ano (R$) ---")
        by_year = df.groupby(df["ano"].astype(str), dropna=False)["valor_reais"].sum()
        for a, v in by_year.items():
            print(f"  {a}: {v:,.2f}")

    salvar_parquet(df, ROOT, verbose=True)


if __name__ == "__main__":
    main()
