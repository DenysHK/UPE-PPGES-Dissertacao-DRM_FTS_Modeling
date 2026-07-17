# ============================================================
# INSTALAÇÃO DE PACOTES (executar fora do script, no terminal)
# python -m pip install pandas matplotlib openpyxl requests
# ============================================================

# -------------------- IMPORTAÇÕES --------------------
# Biblioteca para expressões regulares (busca de padrões em texto)
import re

# Permite tratar strings como arquivos (útil para ler CSV vindo da web)
from io import StringIO

# Manipulação de dados em formato de tabela (DataFrame)
import pandas as pd

# Geração de gráficos
import matplotlib.pyplot as plt

# Requisições HTTP (download de dados)
import requests


# -------------------- FONTES DE DADOS --------------------
# Fonte principal: Copernicus (C3S/ECMWF)
# Contém séries harmonizadas de temperatura global anual
PRIMARY_URL = (
    "https://climate.copernicus.eu/sites/default/files/2026-01/"
    "GCH2025_PR_Fig1_timeseries_annual_global_temperature_anomalies_preindustrial_data_0.csv"
)

# Fonte alternativa (fallback): Met Office
# Utilizada caso a fonte principal falhe
FALLBACK_URL = "https://www.metoffice.gov.uk/hadobs/monitoring/global-temperature.html"


# Lista de colunas esperadas no dataset final
EXPECTED = [
    "Year",
    "ERA5",
    "JRA-3Q",
    "GISTEMPv4",
    "NOAAGlobalTempv6",
    "Berkeley Earth",
    "HadCRUT5",
]


# ============================================================
# FUNÇÃO: DOWNLOAD DE TEXTO
# ============================================================
def download_text(url: str) -> str:
    """
    Faz download do conteúdo de uma URL e retorna como string.

    Parâmetros:
        url (str): endereço web do recurso

    Retorno:
        str: conteúdo da resposta (texto bruto)
    """
    r = requests.get(url, timeout=60)  # timeout evita travamentos
    r.raise_for_status()  # gera erro se a requisição falhar
    return r.text


# ============================================================
# FUNÇÃO: IDENTIFICAR CABEÇALHO DO CSV
# ============================================================
def find_header_and_read_csv(txt: str) -> pd.DataFrame:
    """
    Identifica automaticamente a linha de cabeçalho dentro de um CSV
    que pode conter metadados antes da tabela real.

    Estratégia:
    - Procura linha contendo 'Year' e 'ERA5'
    - Lê a partir dessa linha

    Retorno:
        DataFrame com os dados estruturados
    """
    lines = txt.splitlines()
    header_idx = None

    # Busca da linha que contém o cabeçalho real
    for i, line in enumerate(lines):
        line_clean = line.strip().replace('"', "")
        if "Year" in line_clean and "ERA5" in line_clean:
            header_idx = i
            break

    # Caso não encontre cabeçalho válido
    if header_idx is None:
        raise ValueError("Cabeçalho com 'Year' e 'ERA5' não encontrado.")

    # Reconstrói o CSV a partir do cabeçalho correto
    csv_text = "\n".join(lines[header_idx:])

    # Primeira tentativa: separador por vírgula
    df = pd.read_csv(StringIO(csv_text))

    # Se falhar (ex: CSV europeu com ;)
    if "Year" not in [str(c).strip() for c in df.columns]:
        df = pd.read_csv(StringIO(csv_text), sep=";")

    # Limpeza dos nomes das colunas
    df.columns = [str(c).strip() for c in df.columns]

    return df


# ============================================================
# FUNÇÃO: TENTATIVA DE DOWNLOAD DA FONTE PRINCIPAL
# ============================================================
def try_copernicus() -> pd.DataFrame:
    """
    Tenta carregar dados do Copernicus (fonte primária).
    """
    txt = download_text(PRIMARY_URL)
    df = find_header_and_read_csv(txt)
    return df


# ============================================================
# FUNÇÃO: FALLBACK (MET OFFICE)
# ============================================================
def try_metoffice_fallback() -> pd.DataFrame:
    """
    Caso a fonte principal falhe, busca automaticamente um CSV
    válido dentro da página do Met Office.

    Estratégia:
    - Busca links .csv na página HTML
    - Testa cada um até encontrar estrutura compatível
    """
    html = download_text(FALLBACK_URL)

    # Busca URLs completas de CSV
    matches = re.findall(r'https?://[^"\']+\.csv', html)

    # Se não encontrar, tenta links relativos
    if not matches:
        rel = re.findall(r'["\']([^"\']+\.csv)["\']', html)
        matches = [
            "https://www.metoffice.gov.uk" + m if m.startswith("/") else m
            for m in rel
        ]

    # Testa cada CSV encontrado
    for url in matches:
        try:
            txt = download_text(url)
            df = pd.read_csv(StringIO(txt))
            df.columns = [str(c).strip() for c in df.columns]

            # Verifica se contém estrutura esperada
            if "Year" in df.columns and ("ERA5" in df.columns or "HadCRUT5" in df.columns):
                return df
        except Exception:
            continue

    raise ValueError("Fallback do Met Office não encontrou um CSV utilizável.")


# ============================================================
# FUNÇÃO PRINCIPAL DE CARREGAMENTO E TRATAMENTO
# ============================================================
def load_harmonized_dataframe() -> pd.DataFrame:
    """
    Carrega os dados climáticos harmonizados e realiza:
    - seleção de colunas relevantes
    - conversão numérica
    - limpeza de dados
    """

    try:
        df = try_copernicus()
        source = "Copernicus"
    except Exception as e1:
        print(f"Falha na fonte principal ({e1}). Tentando fallback...")
        df = try_metoffice_fallback()
        source = "Met Office fallback"

    print(f"Fonte usada: {source}")
    print("Colunas originais:", df.columns.tolist())

    # Mantém apenas colunas esperadas
    keep = [c for c in EXPECTED if c in df.columns]

    if "Year" not in keep:
        raise ValueError("A coluna 'Year' não foi encontrada.")

    df = df[keep].copy()

    # Conversão para tipo numérico
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")

    for c in df.columns:
        if c != "Year":
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Remove linhas inválidas
    df = df.dropna(subset=["Year"]).copy()
    df["Year"] = df["Year"].astype(int)

    return df.sort_values("Year").reset_index(drop=True)


# ============================================================
# FUNÇÃO: SALVAR RESULTADOS
# ============================================================
def save_outputs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Salva os dados em dois formatos:
    - formato largo (wide)
    - formato longo (long, ideal para gráficos)
    """

    # Formato original (wide)
    df.to_excel("ecmwf_like_temperature_wide.xlsx", index=False)

    # Conversão para formato longo
    long_df = df.melt(
        id_vars="Year",
        var_name="Dataset",
        value_name="Temp_C_above_preindustrial",
    ).dropna()

    long_df.to_excel("ecmwf_like_temperature_long.xlsx", index=False)

    return long_df


# ============================================================
# FUNÇÃO: PLOTAGEM (ESTILO ECMWF)
# ============================================================
def plot_ecmwf_style(long_df: pd.DataFrame) -> None:
    """
    Gera gráfico inspirado no padrão ECMWF:
    - ERA5 em barras coloridas
    - demais datasets em pontos cinza
    """

    fig, ax = plt.subplots(figsize=(11.5, 7.2))

    # Plot das outras fontes (cinza)
    other = long_df[long_df["Dataset"] != "ERA5"]

    for _, g in other.groupby("Dataset"):
        ax.scatter(
            g["Year"],
            g["Temp_C_above_preindustrial"],
            s=10,
            facecolors="white",
            edgecolors="gray",
            linewidths=0.9,
            zorder=2,
        )

    # ERA5 destacado
    era5 = long_df[long_df["Dataset"] == "ERA5"].sort_values("Year")

    if not era5.empty:
        vals = era5["Temp_C_above_preindustrial"].clip(lower=0)

        # Normalização para colormap
        vmin, vmax = vals.min(), vals.max()
        norm = (vals - vmin) / (vmax - vmin) if vmax != vmin else vals * 0

        colors = plt.cm.YlOrRd(norm)

        # Barras
        ax.bar(
            era5["Year"],
            era5["Temp_C_above_preindustrial"],
            width=0.9,
            color=colors,
            edgecolor="none",
            zorder=1,
        )

        # Linha sobreposta
        ax.plot(
            era5["Year"],
            era5["Temp_C_above_preindustrial"],
            linewidth=1.2,
            zorder=3,
        )

    # Linha de referência climática (1,5 °C)
    ax.axhline(1.5, color="red", linestyle="--", linewidth=2)
    x_center = sum(ax.get_xlim()) / 2

    ax.text(
        x_center,
        1.5,
        "Limite de 1.5 °C",
        color="red",
        fontsize=11,
        ha="center",
        va="bottom",  # fica logo acima da linha
        fontweight="bold"
    )

    # Configuração dos eixos
    ax.set_xlim(long_df["Year"].min() - 3, long_df["Year"].max() + 5)
    ax.set_ylim(0, 1.7)
    #ax.set_yticks(sorted(set(list(ax.get_yticks()) + [1.5])))

    ax.set_ylabel("Aumento médio anual da Temperatura (°C)", fontsize=12)
    ax.set_xlabel("Ano", loc="right", fontsize=12)

    # Estética
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", color="0.85", linewidth=0.8)

    # Salvamento das figuras
    plt.tight_layout(rect=[0, 0.06, 1, 1])
    plt.savefig("ecmwf_like_temperature_figure.png", dpi=300, bbox_inches="tight")
    plt.savefig("ecmwf_like_temperature_figure.svg", bbox_inches="tight")

    plt.show()


# ============================================================
# FUNÇÃO PRINCIPAL
# ============================================================
def main():
    """
    Pipeline completo:
    1. Carrega dados
    2. Trata e limpa
    3. Salva outputs
    4. Gera gráfico
    """

    df = load_harmonized_dataframe()

    print(df.head())
    print(df.tail())
    print(df.columns.tolist())

    long_df = save_outputs(df)
    plot_ecmwf_style(long_df)


# Execução do script
if __name__ == "__main__":
    main()