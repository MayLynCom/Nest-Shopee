import io
import pandas as pd


def _parse_brl(value) -> float:
    """Converte string no formato BRL (ex: '4.890,02') para float."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    # Remove R$, espaços e pontos de milhar, troca vírgula decimal por ponto
    s = s.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_numero(value) -> float:
    """Converte string numérica com possível separador de milhar para float."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _classificar_curva(pct_acumulado: float) -> str:
    if pct_acumulado <= 0.80:
        return "A"
    elif pct_acumulado <= 0.95:
        return "B"
    else:
        return "C"


def processar_produtos(file) -> dict:
    """
    Lê o xlsx de métricas de produtos (aba 'Produtos com Melhor Desempenho').
    Retorna dict com:
      - df: DataFrame apenas com produtos ativos (Excluídos removidos)
      - gmv_bruto: soma de faturamento de TODOS os produtos (incluindo Excluídos)
    """
    df = pd.read_excel(
        file,
        sheet_name="Produtos com Melhor Desempenho",
        engine="openpyxl",
    )

    # Calcular GMV bruto ANTES de remover excluídos
    col_fat = "Vendas (Pedido pago) (BRL)"
    gmv_bruto = df[col_fat].apply(_parse_brl).sum()

    # Remover produtos excluídos da tabela de análise
    df = df[df["Status Atual do Item"] != "Excluído"].copy()

    # Selecionar e renomear colunas relevantes
    df = df[["ID do Item", "Produto", col_fat, "Unidades (Pedido pago)"]].copy()
    df.rename(
        columns={
            "ID do Item": "ID do Produto",
            "Produto": "Nome",
            col_fat: "Faturamento",
            "Unidades (Pedido pago)": "Unidades Vendidas",
        },
        inplace=True,
    )

    # Converter valores
    df["Faturamento"] = df["Faturamento"].apply(_parse_brl)
    df["Unidades Vendidas"] = df["Unidades Vendidas"].apply(_parse_numero).astype(int)
    df["ID do Produto"] = df["ID do Produto"].astype(str).str.strip()

    # Agrupar por ID (caso haja variações duplicadas do mesmo produto)
    df = (
        df.groupby(["ID do Produto", "Nome"], as_index=False)
        .agg({"Faturamento": "sum", "Unidades Vendidas": "sum"})
    )

    # Ticket Médio
    df["Ticket Médio"] = df.apply(
        lambda r: r["Faturamento"] / r["Unidades Vendidas"] if r["Unidades Vendidas"] > 0 else 0.0,
        axis=1,
    )

    # Ordenar por faturamento decrescente para curva ABC
    df = df.sort_values("Faturamento", ascending=False).reset_index(drop=True)

    total = df["Faturamento"].sum()
    if total > 0:
        df["% Acumulado"] = df["Faturamento"].cumsum() / total
    else:
        df["% Acumulado"] = 0.0

    df["Curva"] = df["% Acumulado"].apply(_classificar_curva)

    return {"df": df, "gmv_bruto": gmv_bruto}


def processar_ads_principal(file) -> dict:
    """
    Lê o CSV principal de anúncios Shopee (skiprows=6).
    Retorna dict com:
      - receita_ads: float
      - investimento_ads: float
      - ids_em_ads: set de strings com IDs de produtos em ADS
    """
    content = file.read()
    df = pd.read_csv(
        io.BytesIO(content),
        skiprows=6,
        encoding="utf-8-sig",
        dtype=str,
    )

    # Normalizar nomes de colunas (remover espaços extras)
    df.columns = df.columns.str.strip()

    # Filtrar apenas anúncios em andamento
    if "Status" in df.columns:
        df = df[df["Status"].str.strip() == "Em Andamento"].copy()

    # Converter GMV e Despesas para float (CSV já usa ponto como decimal)
    df["GMV"] = pd.to_numeric(df["GMV"], errors="coerce").fillna(0.0)
    df["Despesas"] = pd.to_numeric(df["Despesas"], errors="coerce").fillna(0.0)

    receita_ads = df["GMV"].sum()
    investimento_ads = df["Despesas"].sum()

    # IDs de produtos em ADS (excluir linhas de grupo marcadas com "-")
    ids_coluna = "ID do produto" if "ID do produto" in df.columns else None
    ids_em_ads: set = set()
    if ids_coluna:
        ids_em_ads = set(
            df[df[ids_coluna].str.strip() != "-"][ids_coluna].str.strip().tolist()
        )

    return {
        "receita_ads": receita_ads,
        "investimento_ads": investimento_ads,
        "ids_em_ads": ids_em_ads,
    }


def processar_grupos_ads(files: list) -> set:
    """
    Lê um ou mais CSVs de grupo de anúncios e retorna o set de IDs de produtos
    encontrados (excluindo grupos marcados com '-').
    """
    ids_grupos: set = set()
    for file in files:
        content = file.read()
        try:
            df = pd.read_csv(
                io.BytesIO(content),
                skiprows=6,
                encoding="utf-8-sig",
                dtype=str,
            )
            df.columns = df.columns.str.strip()

            ids_coluna = "ID do produto" if "ID do produto" in df.columns else None
            if ids_coluna:
                ids_validos = df[df[ids_coluna].str.strip() != "-"][ids_coluna].str.strip()
                ids_grupos.update(ids_validos.tolist())
        except Exception:
            continue

    return ids_grupos


def processar_tudo(
    file_produtos,
    file_ads,
    files_grupos: list | None = None,
) -> dict:
    """
    Processa todos os arquivos e retorna um dict com:
      - df: DataFrame final com todos os produtos
      - gmv: float
      - receita_ads: float
      - investimento_ads: float
      - tacos: float (%)
    """
    # Produtos
    resultado_produtos = processar_produtos(file_produtos)
    df = resultado_produtos["df"]
    gmv = resultado_produtos["gmv_bruto"]  # inclui produtos Excluídos

    # ADS principal
    ads = processar_ads_principal(file_ads)
    receita_ads = ads["receita_ads"]
    investimento_ads = ads["investimento_ads"]
    ids_em_ads = ads["ids_em_ads"]

    # Grupos opcionais
    if files_grupos:
        ids_grupos = processar_grupos_ads(files_grupos)
        ids_em_ads = ids_em_ads | ids_grupos

    # Marcar ADS no DataFrame de produtos
    df["ADS"] = df["ID do Produto"].apply(lambda x: "Sim" if x in ids_em_ads else "Não")

    # TACOS
    tacos = (investimento_ads / gmv * 100) if gmv > 0 else 0.0

    return {
        "df": df,
        "gmv": gmv,
        "receita_ads": receita_ads,
        "investimento_ads": investimento_ads,
        "tacos": tacos,
    }
