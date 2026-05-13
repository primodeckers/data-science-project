# Projeto: investimentos federais (GND 4 e 5)

Análise exploratória de pagamentos liquidados (governo central, 2021–2025), com pipeline em Python, notebook, figuras para relatório e exportação para Word.

## Estrutura do repositório

| Caminho | Descrição |
|---------|-----------|
| `data/investimentos_2021_2025.parquet` | Base unificada gerada pelos scripts (versionada). |
| `figures/` | PNG usados no `relatorio_investimentos.md` (gerados por script). |
| `notebooks/exploracao_investimentos.ipynb` | Exploração interativa sobre o Parquet. |
| `scripts/carregar_investimentos.py` | Lê CSV + XLSX na raiz e grava o Parquet. |
| `scripts/gerar_figuras_relatorio.py` | Recria as figuras a partir do Parquet. |
| `scripts/export_relatorio_docx.py` | Exporta o Markdown para `relatorio_investimentos.docx` (capa opcional via JSON). |
| `scripts/enriquecer_graficos_notebook.py` | Apoio ao notebook (gráficos enriquecidos). |
| `relatorio_investimentos.md` | Fonte do relatório em Markdown. |
| `relatorio_investimentos.docx` | Relatório em Word (pode voltar a gerar com o script). |
| `relatorio_cabecalho.json` | Campos da capa IDP para o DOCX. |
| `metadados-investimentos.pdf` | Metadados / dicionário de campos dos ficheiros de origem. |
| `dicionario-de-dados.odt` | Dicionário de dados (documento de apoio). |
| `requirements.txt` | Dependências Python. |
| `trabalho-final-projeto-integrado.md` | Enunciado / texto de apoio ao projeto. |

Outros CSV na raiz (`conjunto-dados.csv`, `temas-compras-contratos.csv`) são auxiliares conforme o uso do curso.

## O que **não** vai para o Git

Definido em `.gitignore`:

- **Dados brutos** `investimentos_2021.csv` … `investimentos_2024.csv` e `investimentos-2025.xlsx` (dezenas de MB; obtidos pelo portal / disciplina e colocados localmente na raiz).
- Ambientes virtuais, cache Python, checkpoints de Jupyter, ficheiros temporários e pastas comuns de IDE.

O Parquet em `data/` permanece no repositório para quem clonar poder correr figuras e notebook sem repor os CSV.

## Pré-requisitos

- Python 3.11+ (recomendado).
- `pip` atualizado.

## Configuração

Na raiz do projeto:

```bash
python -m venv .venv
```

Ative o ambiente (Windows PowerShell: `.venv\Scripts\Activate.ps1`; Git Bash: `source .venv/Scripts/activate`) e instale dependências:

```bash
pip install -r requirements.txt
```

## Fluxo de trabalho típico

1. **Dados brutos (opcional se já tiver o Parquet)**  
   Coloque na raiz: `investimentos_2021.csv` … `investimentos_2024.csv` e `investimentos-2025.xlsx`. Depois:

   ```bash
   python scripts/carregar_investimentos.py
   ```

   Isto cria ou sobrescreve `data/investimentos_2021_2025.parquet`.

2. **Figuras do relatório**

   ```bash
   python scripts/gerar_figuras_relatorio.py
   ```

3. **Notebook**  
   Abra `notebooks/exploracao_investimentos.ipynb` com o kernel do `.venv`.

4. **Word**  
   Com as figuras presentes:

   ```bash
   python scripts/export_relatorio_docx.py
   ```

   Opções úteis: `--gerar-figuras` (corre o gerador de figuras antes), `--no-cabecalho`, `-i` / `-o` para caminhos alternativos. Ver docstring em `scripts/export_relatorio_docx.py`.

## Documentação adicional

- Metadados e significado das colunas: `metadados-investimentos.pdf` e `dicionario-de-dados.odt`.
- Comentários de uso no topo de cada script em `scripts/`.

## Licença

Ver ficheiro `LICENSE` na raiz.
