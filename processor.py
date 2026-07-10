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


def _detectar_skiprows(content: bytes, coluna_referencia: str = "GMV") -> int:
    """
    Detecta automaticamente quantas linhas pular para encontrar o cabeçalho
    que contém a coluna de referência. Tenta utf-8-sig e latin-1, separadores ; e ,
    """
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            texto = content.decode(encoding)
            for i, linha in enumerate(texto.splitlines()):
                for sep in (";", ","):
                    partes = [p.strip().strip('"') for p in linha.split(sep)]
                    if coluna_referencia in partes:
                        return i
        except Exception:
            continue
    return 0


def _ler_csv_shopee(content: bytes, skiprows: int) -> pd.DataFrame:
    """Lê CSVs da Shopee detectando automaticamente vírgula ou ponto e vírgula."""
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(
                io.BytesIO(content),
                skiprows=skiprows,
                sep=None,
                engine="python",
                encoding=encoding,
                dtype=str,
            )
        except UnicodeDecodeError:
            continue
    return pd.read_csv(
        io.BytesIO(content),
        skiprows=skiprows,
        sep=None,
        engine="python",
        encoding="utf-8-sig",
        dtype=str,
    )


def processar_produtos(file) -> dict:
    """
    Lê o xlsx parentskudetail, aba 'Produtos com Melhor Desempenho'.
    Colunas esperadas: ID do Item, Produto, Vendas (Pedido pago) (BRL), Unidades (Pedido pago).
    Se a coluna 'ID da Variação' existir, mantém apenas linhas onde ela é diferente de '-'.
    Retorna dict com:
      - df: DataFrame processado (curva ABC, ticket médio)
      - gmv_bruto: soma de Vendas (Pedido pago) (BRL) de todas as linhas da aba
    """
    df = pd.read_excel(
        file,
        sheet_name="Produtos com Melhor Desempenho",
        engine="openpyxl",
    )
    df.columns = df.columns.str.strip()

    if "ID da Variação" in df.columns:
        df = df[df["ID da Variação"].astype(str).str.strip() == "-"].copy()

    col_fat = "Vendas (Pedido pago) (BRL)"
    col_unid = "Unidades (Pedido pago)"
    gmv_bruto = pd.to_numeric(df[col_fat].apply(_parse_brl), errors="coerce").fillna(0.0).sum()

    df = df[["ID do Item", "Produto", col_fat, col_unid]].copy()
    df.rename(
        columns={
            "ID do Item": "ID do Produto",
            "Produto": "Nome",
            col_fat: "Faturamento",
            col_unid: "Unidades Vendidas",
        },
        inplace=True,
    )

    # Converter valores — cast explícito para float/int para evitar object dtype
    df["Faturamento"] = pd.to_numeric(
        df["Faturamento"].apply(_parse_brl), errors="coerce"
    ).fillna(0.0)
    df["Unidades Vendidas"] = pd.to_numeric(
        df["Unidades Vendidas"].apply(_parse_numero), errors="coerce"
    ).fillna(0).astype(int)
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
    Lê o CSV principal de anúncios Shopee.
    Detecta automaticamente a linha do cabeçalho procurando pela coluna 'GMV'.
    Soma GMV e Despesas sobre todas as linhas (receita e investimento ADS).
    Os IDs marcados como em ADS vêm apenas das linhas com Status «Em Andamento».
    Retorna dict com:
      - receita_ads: float
      - investimento_ads: float
      - ids_em_ads: set de strings com IDs de produtos em ADS em andamento
      - despesas_por_produto: dict {id_produto: total_despesas}
    """
    content = file.read()
    skiprows = _detectar_skiprows(content, "GMV")

    df = _ler_csv_shopee(content, skiprows)

    # Normalizar nomes de colunas (remover espaços extras)
    df.columns = df.columns.str.strip()

    if "GMV" not in df.columns:
        raise KeyError(
            f"Coluna 'GMV' não encontrada. Colunas disponíveis: {df.columns.tolist()}"
        )
    if "Despesas" not in df.columns:
        raise KeyError(
            f"Coluna 'Despesas' não encontrada. Colunas disponíveis: {df.columns.tolist()}"
        )

    # Converter GMV e Despesas para float (CSV já usa ponto como decimal)
    df["GMV"] = pd.to_numeric(df["GMV"], errors="coerce").fillna(0.0)
    df["Despesas"] = pd.to_numeric(df["Despesas"], errors="coerce").fillna(0.0)

    receita_ads = df["GMV"].sum()
    investimento_ads = df["Despesas"].sum()

    # IDs só de anúncios em andamento (após já ter totals do arquivo inteiro)
    if "Status" in df.columns:
        df_para_ids = df[df["Status"].str.strip() == "Em Andamento"].copy()
    else:
        df_para_ids = df

    ids_coluna = "ID do produto" if "ID do produto" in df_para_ids.columns else None
    ids_em_ads: set = set()
    despesas_por_produto: dict = {}

    if ids_coluna:
        df_validos = df_para_ids[df_para_ids[ids_coluna].str.strip() != "-"].copy()
        df_validos["_id"] = df_validos[ids_coluna].str.strip()
        ids_em_ads = set(df_validos["_id"].tolist())

        # Agrupa despesas por produto
        agrupado = df_validos.groupby("_id")["Despesas"].sum()
        despesas_por_produto = agrupado.to_dict()

    return {
        "receita_ads": receita_ads,
        "investimento_ads": investimento_ads,
        "ids_em_ads": ids_em_ads,
        "despesas_por_produto": despesas_por_produto,
    }


def processar_grupos_ads(files: list) -> dict:
    """
    Lê um ou mais CSVs de grupo de anúncios e retorna dict com:
      - ids: set de IDs de produtos encontrados (excluindo linhas de grupo com '-')
      - despesas_por_produto: dict {id: float} com soma de Despesas por produto
    """
    ids_grupos: set = set()
    despesas_grupos: dict = {}
    for file in files:
        content = file.read()
        try:
            skiprows = _detectar_skiprows(content, "GMV")
            df = _ler_csv_shopee(content, skiprows)
            df.columns = df.columns.str.strip()

            ids_coluna = "ID do produto" if "ID do produto" in df.columns else None
            if ids_coluna:
                df_produtos = df[df[ids_coluna].str.strip() != "-"].copy()
                ids_validos = df_produtos[ids_coluna].str.strip()
                ids_grupos.update(ids_validos.tolist())

                if "Despesas" in df.columns:
                    df_produtos["_id"] = ids_validos.values
                    df_produtos["_despesa"] = pd.to_numeric(
                        df_produtos["Despesas"], errors="coerce"
                    ).fillna(0.0)
                    for _, row in df_produtos[["_id", "_despesa"]].iterrows():
                        pid = row["_id"]
                        despesas_grupos[pid] = despesas_grupos.get(pid, 0.0) + row["_despesa"]
        except Exception:
            continue

    return {"ids": ids_grupos, "despesas_por_produto": despesas_grupos}


def processar_tudo(
    file_produtos,
    file_ads=None,
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
    gmv = resultado_produtos["gmv_bruto"]

    # ADS principal opcional
    receita_ads = 0.0
    investimento_ads = 0.0
    ids_em_ads: set = set()
    despesas_por_produto: dict = {}
    if file_ads:
        ads = processar_ads_principal(file_ads)
        receita_ads = ads["receita_ads"]
        investimento_ads = ads["investimento_ads"]
        ids_em_ads = ads["ids_em_ads"]
        despesas_por_produto = ads["despesas_por_produto"]

    # Grupos opcionais
    if files_grupos:
        resultado_grupos = processar_grupos_ads(files_grupos)
        ids_em_ads = ids_em_ads | resultado_grupos["ids"]
        for pid, gasto in resultado_grupos["despesas_por_produto"].items():
            despesas_por_produto[pid] = despesas_por_produto.get(pid, 0.0) + gasto

    # Marcar ADS no DataFrame de produtos
    df["ADS"] = df["ID do Produto"].apply(lambda x: "Sim" if x in ids_em_ads else "Não")

    # Gasto ADS por produto e TACOS por produto
    df["Gasto ADS"] = df["ID do Produto"].apply(
        lambda x: despesas_por_produto.get(x, 0.0)
    )
    df["TACOS Produto"] = df.apply(
        lambda r: (r["Gasto ADS"] / r["Faturamento"] * 100) if r["Faturamento"] > 0 else 0.0,
        axis=1,
    )

    # TACOS global
    tacos = (investimento_ads / gmv * 100) if gmv > 0 else 0.0

    return {
        "df": df,
        "gmv": gmv,
        "receita_ads": receita_ads,
        "investimento_ads": investimento_ads,
        "tacos": tacos,
    }
