# -*- coding: utf-8 -*-
"""
Gera a memória final de cálculo para o acoplamento conceitual RSB-SFT.

Entrada esperada: dois arquivos .xlsx no mesmo diretório deste script:
  1) RSB_Results.xlsx
  2) RESULTADOS_Gray2024_FeKSiO2_HTFT_H2CO_1.80a2.20_WF_P5a40.xlsx

Saída padrão:
  Selecao_acoplamento_RSB_SFT.xlsx

Observação metodológica:
  O script não faz simulação nova nem define uma planta industrial integrada.
  Ele lê as planilhas de resultados já geradas, expande os resultados em formato
  longo, calcula a aderência do syngas da RSB ao alvo definido pela SFT e documenta
  a memória de cálculo da seleção de cenários destacados.

Dependência externa apenas para escrever o arquivo .xlsx de saída:
  pip install openpyxl

Autor: gerado para apoiar a dissertação de Denys Haluli Kabbaz.
"""

from __future__ import annotations

import argparse
import math
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as exc:
    raise SystemExit(
        "Este script precisa do pacote openpyxl para gerar a planilha de saída.\n"
        "Instale com: pip install openpyxl"
    ) from exc


# -----------------------------------------------------------------------------
# Parâmetros editáveis da análise
# -----------------------------------------------------------------------------
PARAMETROS = {
    # Critérios para definir a condição-alvo da SFT usada no acoplamento conceitual.
    # Estes critérios preservam a lógica defendida na dissertação: não basta escolher
    # o maior valor global, pois baixa pressão pode representar apenas resposta do
    # domínio computacional.
    "SFT_P_MIN_BAR": 15.0,
    "SFT_H2CO_MIN": 1.90,
    "SFT_H2CO_MAX": 2.10,
    "SFT_F_SCCM_REFERENCIA": 15.0,
    "SFT_USAR_F_REFERENCIA": True,

    # Critérios de ranqueamento da RSB.
    # A proximidade da razão H2/CO é o critério primário. Conversões e seletividade
    # relativa a carbono sólido são critérios secundários para desempate e leitura.
    "RSB_TOP_N_DESTAQUE": 30,
    "RSB_TOLERANCIA_H2CO_ESTRITA": 0.10,
    "RSB_TOLERANCIA_H2CO_AMPLIADA": 0.20,

    # Pesos usados apenas como índice auxiliar editável. O ranking principal continua
    # registrado por proximidade de H2/CO, conversão média e seletividade a carbono.
    "PESO_H2CO": 0.50,
    "PESO_CONVERSAO": 0.30,
    "PESO_CARBONO": 0.20,
}


# -----------------------------------------------------------------------------
# Utilitários gerais
# -----------------------------------------------------------------------------

def normalizar_texto(valor: Any) -> str:
    """Normaliza texto para busca robusta em rótulos de planilhas."""
    if valor is None:
        return ""
    texto = str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = texto.replace("₂", "2").replace("₅", "5").replace("₁", "1").replace("₁₂", "12")
    texto = texto.replace("–", "-").replace("—", "-")
    texto = re.sub(r"\s+", " ", texto.strip().lower())
    return texto


def para_float(valor: Any) -> Optional[float]:
    """Converte valores numéricos com vírgula/ponto/notação científica para float."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        if math.isfinite(float(valor)):
            return float(valor)
        return None
    texto = str(valor).strip()
    if not texto:
        return None
    texto = texto.replace("K", "").replace("°C", "").replace("ºC", "")
    texto = texto.replace("%", "").strip()
    # Trata vírgula decimal apenas quando não há ponto decimal.
    if "," in texto and "." not in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        out = float(texto)
    except ValueError:
        return None
    return out if math.isfinite(out) else None


def letra_coluna(indice_1_base: int) -> str:
    return get_column_letter(indice_1_base)


def chave_dedup(valores: Iterable[Optional[float]], casas: int = 8) -> Tuple[Any, ...]:
    chave = []
    for v in valores:
        if v is None:
            chave.append(None)
        else:
            chave.append(round(float(v), casas))
    return tuple(chave)


# -----------------------------------------------------------------------------
# Leitor .xlsx rápido por XML; evita depender de openpyxl para ler arquivos grandes.
# -----------------------------------------------------------------------------

class XLSXLeitor:
    NS = {
        "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    def __init__(self, caminho: Path):
        self.caminho = Path(caminho)
        self.zip = zipfile.ZipFile(self.caminho)
        self.shared_strings = self._ler_shared_strings()
        self.sheets = self._ler_sheets()

    def _ler_shared_strings(self) -> List[str]:
        if "xl/sharedStrings.xml" not in self.zip.namelist():
            return []
        root = ET.fromstring(self.zip.read("xl/sharedStrings.xml"))
        strings = []
        for si in root.findall("a:si", self.NS):
            partes = [t.text or "" for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
            strings.append("".join(partes))
        return strings

    def _ler_sheets(self) -> Dict[str, str]:
        wb = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        out = {}
        for sh in wb.find("a:sheets", self.NS):
            nome = sh.attrib["name"]
            rid = sh.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = relmap[rid]
            if not target.startswith("xl/"):
                target = "xl/" + target
            out[nome] = target
        return out

    @staticmethod
    def _rc(celula_ref: str) -> Tuple[int, int]:
        m = re.match(r"([A-Z]+)(\d+)", celula_ref)
        if not m:
            raise ValueError(f"Referência de célula inválida: {celula_ref}")
        col = 0
        for ch in m.group(1):
            col = col * 26 + ord(ch) - 64
        return int(m.group(2)), col

    def matriz(self, sheet_name: str) -> List[List[Any]]:
        target = self.sheets[sheet_name]
        root = ET.fromstring(self.zip.read(target))
        data: Dict[Tuple[int, int], Any] = {}
        maxr = 0
        maxc = 0
        for c in root.findall(".//a:sheetData/a:row/a:c", self.NS):
            ref = c.attrib.get("r")
            if not ref:
                continue
            r, col = self._rc(ref)
            tipo = c.attrib.get("t")
            valor_node = c.find("a:v", self.NS)
            valor: Any = "" if valor_node is None else valor_node.text
            if tipo == "s" and valor != "":
                valor = self.shared_strings[int(valor)]
            elif tipo == "inlineStr":
                partes = [t.text or "" for t in c.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
                valor = "".join(partes)
            data[(r, col)] = valor
            maxr = max(maxr, r)
            maxc = max(maxc, col)
        return [[data.get((r, c), "") for c in range(1, maxc + 1)] for r in range(1, maxr + 1)]


# -----------------------------------------------------------------------------
# Extração dos resultados da SFT
# -----------------------------------------------------------------------------

def localizar_linha(matriz: List[List[Any]], ini: int, fim: int, termo: str) -> Optional[int]:
    termo_n = normalizar_texto(termo)
    for r in range(ini, min(fim, len(matriz))):
        texto_linha = " | ".join(normalizar_texto(v) for v in matriz[r])
        if termo_n in texto_linha:
            return r
    return None


def valor_linha_coluna(matriz: List[List[Any]], row_idx: Optional[int], col_idx: int) -> Optional[float]:
    if row_idx is None or row_idx < 0 or row_idx >= len(matriz):
        return None
    if col_idx < 0 or col_idx >= len(matriz[row_idx]):
        return None
    return para_float(matriz[row_idx][col_idx])


def detectar_blocos_sft(matriz: List[List[Any]]) -> List[Tuple[int, int, Optional[float]]]:
    starts: List[Tuple[int, Optional[float]]] = []
    for r, linha in enumerate(matriz):
        texto = " | ".join(str(v) for v in linha if v not in (None, ""))
        if "SFT Fe-K" in texto and "F =" in texto:
            m = re.search(r"F\s*=\s*([0-9]+(?:[\.,][0-9]+)?)\s*SCCM", texto, flags=re.I)
            f = para_float(m.group(1)) if m else None
            starts.append((r, f))
    blocos = []
    for i, (start, f) in enumerate(starts):
        end = starts[i + 1][0] if i + 1 < len(starts) else len(matriz)
        blocos.append((start, end, f))
    return blocos


def extrair_sft(caminho_sft: Path) -> List[Dict[str, Any]]:
    leitor = XLSXLeitor(caminho_sft)
    resultados: List[Dict[str, Any]] = []
    for aba in leitor.sheets:
        if not normalizar_texto(aba).startswith("h2_co"):
            continue
        matriz = leitor.matriz(aba)
        # H2/CO também é lido de linha interna; sheet name é apenas fallback.
        m = re.search(r"([0-9]+)[\.,]([0-9]+)", aba)
        h2co_sheet = float(f"{m.group(1)}.{m.group(2)}") if m else None
        for ini, fim, f_bloco in detectar_blocos_sft(matriz):
            rows = {
                "T_K": localizar_linha(matriz, ini, fim, "Temperatura (K)"),
                "T_C": localizar_linha(matriz, ini, fim, "Temperatura (C)"),
                "P_Pa": localizar_linha(matriz, ini, fim, "Pressao (Pa)"),
                "P_bar": localizar_linha(matriz, ini, fim, "Pressao (bar)"),
                "F_SCCM": localizar_linha(matriz, ini, fim, "SYNGAS Flow IN (SCCM)"),
                "F_mol_s": localizar_linha(matriz, ini, fim, "SYNGAS Flow IN (mol/s)"),
                "H2CO": localizar_linha(matriz, ini, fim, "H2/CO"),
                "FH2_in": localizar_linha(matriz, ini, fim, "FH2,IN"),
                "FCO_in": localizar_linha(matriz, ini, fim, "FCO,IN"),
                "WF_total": localizar_linha(matriz, ini, fim, "W/F total"),
                "WFCO_in": localizar_linha(matriz, ini, fim, "W/FCO,in"),
                "XCO": localizar_linha(matriz, ini, fim, "XCO (%)"),
                "XH2": localizar_linha(matriz, ini, fim, "XH2 (%)"),
                "alpha": localizar_linha(matriz, ini, fim, "alpha aparente C5-C11"),
                "S_C5_C11": localizar_linha(matriz, ini, fim, "S C5-C11 (%)"),
                "S_C12p": localizar_linha(matriz, ini, fim, "S C12+ (%)"),
                "S_CO2": localizar_linha(matriz, ini, fim, "S CO2 (%)"),
                "eta_C5_C11": localizar_linha(matriz, ini, fim, "eta_C C5-C11 (%)"),
                "Y_C5_C11": localizar_linha(matriz, ini, fim, "Y C5-C11 molar (%)"),
                "OP_C5_C11": localizar_linha(matriz, ini, fim, "olefinas/parafinas C5-C11"),
            }
            max_cols = max(len(r) for r in matriz) if matriz else 0
            for c in range(max_cols):
                t_c = valor_linha_coluna(matriz, rows["T_C"], c)
                p_bar = valor_linha_coluna(matriz, rows["P_bar"], c)
                h2co = valor_linha_coluna(matriz, rows["H2CO"], c) or h2co_sheet
                f_sccm = valor_linha_coluna(matriz, rows["F_SCCM"], c) or f_bloco
                eta = valor_linha_coluna(matriz, rows["eta_C5_C11"], c)
                xco = valor_linha_coluna(matriz, rows["XCO"], c)
                if t_c is None or p_bar is None or h2co is None or f_sccm is None or eta is None or xco is None:
                    continue
                resultados.append({
                    "fonte_arquivo": caminho_sft.name,
                    "fonte_aba": aba,
                    "linha_bloco_inicio": ini + 1,
                    "coluna_origem": letra_coluna(c + 1),
                    "H2CO_SFT": h2co,
                    "T_C": t_c,
                    "T_K": valor_linha_coluna(matriz, rows["T_K"], c),
                    "P_bar": p_bar,
                    "P_Pa": valor_linha_coluna(matriz, rows["P_Pa"], c),
                    "F_SCCM": f_sccm,
                    "F_mol_s": valor_linha_coluna(matriz, rows["F_mol_s"], c),
                    "FH2_in": valor_linha_coluna(matriz, rows["FH2_in"], c),
                    "FCO_in": valor_linha_coluna(matriz, rows["FCO_in"], c),
                    "W_Ftotal": valor_linha_coluna(matriz, rows["WF_total"], c),
                    "W_FCO_in": valor_linha_coluna(matriz, rows["WFCO_in"], c),
                    "XCO_pct": xco,
                    "XH2_pct": valor_linha_coluna(matriz, rows["XH2"], c),
                    "alpha_aparente": valor_linha_coluna(matriz, rows["alpha"], c),
                    "S_C5_C11_pct": valor_linha_coluna(matriz, rows["S_C5_C11"], c),
                    "S_C12p_pct": valor_linha_coluna(matriz, rows["S_C12p"], c),
                    "S_CO2_pct": valor_linha_coluna(matriz, rows["S_CO2"], c),
                    "eta_C5_C11_pct": eta,
                    "Y_C5_C11_molar_pct": valor_linha_coluna(matriz, rows["Y_C5_C11"], c),
                    "OP_C5_C11": valor_linha_coluna(matriz, rows["OP_C5_C11"], c),
                })
    return resultados


# -----------------------------------------------------------------------------
# Extração dos resultados da RSB
# -----------------------------------------------------------------------------

def sheet_rsb_valida(nome: str) -> bool:
    n = normalizar_texto(nome)
    if "escolh" in n or "acopla" in n:
        return False
    return "pa" in n or "temperatura" in n or "exp" in n


def extrair_rsb(caminho_rsb: Path) -> List[Dict[str, Any]]:
    leitor = XLSXLeitor(caminho_rsb)
    registros: List[Dict[str, Any]] = []
    for aba in leitor.sheets:
        if not sheet_rsb_valida(aba):
            continue
        matriz = leitor.matriz(aba)
        rows = {
            "T_K": localizar_linha(matriz, 0, len(matriz), "Temperatura ( K )"),
            "P_Pa": localizar_linha(matriz, 0, len(matriz), "Pressão ( Pa )"),
            "F_biogas": localizar_linha(matriz, 0, len(matriz), "Flow (mol/s)"),
            "CH4_frac": localizar_linha(matriz, 0, len(matriz), "CH4 (fração molar)"),
            "CO2_frac": localizar_linha(matriz, 0, len(matriz), "CO2 (fração molar)"),
            "CH4_out": localizar_linha(matriz, 0, len(matriz), "CH4 ( mol/s )"),
            "CO2_out": localizar_linha(matriz, 0, len(matriz), "CO2 ( mol/s)"),
            "CO_out": localizar_linha(matriz, 0, len(matriz), "CO (mol/s)"),
            "H2O_out": localizar_linha(matriz, 0, len(matriz), "H2O ( mol/s)"),
            "C_out": localizar_linha(matriz, 0, len(matriz), "C ( mol/s)"),
            "H2_out": localizar_linha(matriz, 0, len(matriz), "H2 (mol/s)"),
            "OUT_total": localizar_linha(matriz, 0, len(matriz), "OUT Molar Flow"),
            "DRM": localizar_linha(matriz, 0, len(matriz), "Reaction Coordinate  DRM"),
            "RWGS": localizar_linha(matriz, 0, len(matriz), "Reaction Coordinate RWGS"),
            "DM": localizar_linha(matriz, 0, len(matriz), "Reaction Coordinate DM"),
            "XCH4": localizar_linha(matriz, 0, len(matriz), "Convertion CH4"),
            "XCO2": localizar_linha(matriz, 0, len(matriz), "Convertion CO2"),
            "SC": localizar_linha(matriz, 0, len(matriz), "Selec ( C )"),
            "H2CO": localizar_linha(matriz, 0, len(matriz), "H2 / CO"),
        }
        if rows["T_K"] is None or rows["P_Pa"] is None or rows["CH4_frac"] is None or rows["H2CO"] is None:
            continue
        max_cols = max(len(r) for r in matriz) if matriz else 0
        for c in range(max_cols):
            h2co = valor_linha_coluna(matriz, rows["H2CO"], c)
            t_k = valor_linha_coluna(matriz, rows["T_K"], c)
            p_pa = valor_linha_coluna(matriz, rows["P_Pa"], c)
            ch4 = valor_linha_coluna(matriz, rows["CH4_frac"], c)
            co2 = valor_linha_coluna(matriz, rows["CO2_frac"], c)
            if h2co is None or t_k is None or p_pa is None or ch4 is None or co2 is None:
                continue
            xch4 = valor_linha_coluna(matriz, rows["XCH4"], c)
            xco2 = valor_linha_coluna(matriz, rows["XCO2"], c)
            registros.append({
                "fonte_arquivo": caminho_rsb.name,
                "fonte_aba": aba,
                "coluna_origem": letra_coluna(c + 1),
                "T_K": t_k,
                "T_C": t_k - 273.15 if t_k is not None else None,
                "P_Pa": p_pa,
                "P_atm_aprox": p_pa / 101325.0 if p_pa is not None else None,
                "F_biogas_mol_s": valor_linha_coluna(matriz, rows["F_biogas"], c),
                "CH4_frac": ch4,
                "CO2_frac": co2,
                "H2CO_RSB": h2co,
                "XCH4_pct": xch4,
                "XCO2_pct": xco2,
                "X_media_pct": (xch4 + xco2) / 2.0 if xch4 is not None and xco2 is not None else None,
                "SC_relativo": valor_linha_coluna(matriz, rows["SC"], c),
                "CO_out_mol_s": valor_linha_coluna(matriz, rows["CO_out"], c),
                "H2_out_mol_s": valor_linha_coluna(matriz, rows["H2_out"], c),
                "CO2_out_mol_s": valor_linha_coluna(matriz, rows["CO2_out"], c),
                "CH4_out_mol_s": valor_linha_coluna(matriz, rows["CH4_out"], c),
                "H2O_out_mol_s": valor_linha_coluna(matriz, rows["H2O_out"], c),
                "C_out_mol_s": valor_linha_coluna(matriz, rows["C_out"], c),
                "OUT_total_mol_s": valor_linha_coluna(matriz, rows["OUT_total"], c),
                "DRM_mol_s": valor_linha_coluna(matriz, rows["DRM"], c),
                "RWGS_mol_s": valor_linha_coluna(matriz, rows["RWGS"], c),
                "DM_mol_s": valor_linha_coluna(matriz, rows["DM"], c),
            })

    # Deduplicação: mantém um registro por T/P/composição/vazão, preferindo abas EXP quando houver empate.
    dedup: Dict[Tuple[Any, ...], Dict[str, Any]] = {}
    for r in registros:
        key = chave_dedup([r.get("T_K"), r.get("P_Pa"), r.get("F_biogas_mol_s"), r.get("CH4_frac"), r.get("CO2_frac")])
        atual = dedup.get(key)
        if atual is None:
            dedup[key] = r
        else:
            score_novo = 1 if "exp" in normalizar_texto(r.get("fonte_aba")) else 0
            score_atual = 1 if "exp" in normalizar_texto(atual.get("fonte_aba")) else 0
            if score_novo > score_atual:
                dedup[key] = r
    return list(dedup.values())


# -----------------------------------------------------------------------------
# Cálculos de seleção e memória
# -----------------------------------------------------------------------------

def selecionar_alvo_sft(sft: List[Dict[str, Any]], params: Dict[str, Any]) -> Tuple[Dict[str, Any], Dict[str, List[Dict[str, Any]]]]:
    def eta(row: Dict[str, Any]) -> float:
        return row.get("eta_C5_C11_pct") if row.get("eta_C5_C11_pct") is not None else -1e99

    p_min = params["SFT_P_MIN_BAR"]
    hmin = params["SFT_H2CO_MIN"]
    hmax = params["SFT_H2CO_MAX"]
    fref = params["SFT_F_SCCM_REFERENCIA"]
    usar_f = params["SFT_USAR_F_REFERENCIA"]

    global_top = sorted(sft, key=lambda r: eta(r), reverse=True)[:20]
    p_ge = [r for r in sft if r.get("P_bar") is not None and r["P_bar"] >= p_min]
    p_ge_top = sorted(p_ge, key=lambda r: eta(r), reverse=True)[:20]
    pref = [r for r in sft if (r.get("P_bar") is not None and r.get("H2CO_SFT") is not None and
                               r["P_bar"] >= p_min and hmin <= r["H2CO_SFT"] <= hmax)]
    if usar_f:
        pref_f = [r for r in pref if r.get("F_SCCM") is not None and abs(r["F_SCCM"] - fref) < 1e-6]
        if pref_f:
            pref = pref_f
    if not pref:
        pref = p_ge if p_ge else sft

    # Critério de ordenação: eficiência em C5-C11, proximidade de H2/CO=2,10,
    # proximidade da vazão de referência e pressão intermediária.
    href = hmax
    alvo = sorted(
        pref,
        key=lambda r: (
            -eta(r),
            abs((r.get("H2CO_SFT") or 0) - href),
            abs((r.get("F_SCCM") or fref) - fref),
            abs((r.get("P_bar") or p_min) - p_min),
        ),
    )[0]
    grupos = {
        "maior_valor_observado_dominio_total": global_top,
        "maior_valor_observado_P_ge_min": p_ge_top,
        "cenarios_alvo_preferencial": sorted(pref, key=lambda r: eta(r), reverse=True)[:20],
    }
    return alvo, grupos


def calcular_memoria_rsb(rsb: List[Dict[str, Any]], alvo_sft: Dict[str, Any], params: Dict[str, Any]) -> List[Dict[str, Any]]:
    h_alvo = alvo_sft["H2CO_SFT"]
    sc_vals = [r["SC_relativo"] for r in rsb if r.get("SC_relativo") is not None]
    sc_min = min(sc_vals) if sc_vals else 0.0
    sc_max = max(sc_vals) if sc_vals else 1.0
    delta_max = max([abs((r.get("H2CO_RSB") or 0) - h_alvo) for r in rsb] or [1.0])
    out = []
    for r in rsb:
        row = dict(r)
        h = row.get("H2CO_RSB")
        delta = abs(h - h_alvo) if h is not None else None
        row["H2CO_ALVO_SFT"] = h_alvo
        row["delta_H2CO_abs"] = delta
        row["delta_H2CO_pct_relativo"] = (delta / h_alvo * 100.0) if delta is not None and h_alvo else None
        tol_e = params["RSB_TOLERANCIA_H2CO_ESTRITA"]
        tol_a = params["RSB_TOLERANCIA_H2CO_AMPLIADA"]
        if delta is None:
            classe = "sem H2/CO calculável"
        elif delta <= tol_e:
            classe = "aderência estrita"
        elif delta <= tol_a:
            classe = "aderência ampliada"
        else:
            classe = "fora da faixa de aderência"
        row["classe_aderencia"] = classe
        if delta is None:
            score_h = 0.0
        else:
            score_h = 1.0 - min(delta / delta_max, 1.0) if delta_max > 0 else 1.0
        xmedia = row.get("X_media_pct")
        score_conv = max(0.0, min((xmedia or 0.0) / 100.0, 1.0))
        sc = row.get("SC_relativo")
        if sc is None or sc_max == sc_min:
            score_c = 0.0
        else:
            # menor seletividade relativa a carbono sólido é melhor
            score_c = 1.0 - ((sc - sc_min) / (sc_max - sc_min))
        row["score_H2CO_0a1"] = score_h
        row["score_conversao_0a1"] = score_conv
        row["score_carbono_0a1"] = score_c
        row["indice_auxiliar_0a1"] = (
            params["PESO_H2CO"] * score_h +
            params["PESO_CONVERSAO"] * score_conv +
            params["PESO_CARBONO"] * score_c
        )
        if h is None:
            necessidade = "não avaliável"
        elif h < h_alvo:
            necessidade = "RSB abaixo do alvo: requer elevar H2/CO ou aceitar alimentação subestequiométrica"
        elif h > h_alvo:
            necessidade = "RSB acima do alvo: requer reduzir H2/CO ou aceitar alimentação mais hidrogenada"
        else:
            necessidade = "aderente ao alvo"
        row["leitura_condicionamento"] = necessidade
        out.append(row)
    out.sort(key=lambda r: (
        r.get("delta_H2CO_abs") if r.get("delta_H2CO_abs") is not None else 1e99,
        -(r.get("X_media_pct") if r.get("X_media_pct") is not None else -1e99),
        r.get("SC_relativo") if r.get("SC_relativo") is not None else 1e99,
    ))
    for i, row in enumerate(out, start=1):
        row["rank_primario"] = i
    out.sort(key=lambda r: r["indice_auxiliar_0a1"], reverse=True)
    for i, row in enumerate(out, start=1):
        row["rank_indice_auxiliar"] = i
    # Volta ao ranking primário para leitura principal.
    out.sort(key=lambda r: r["rank_primario"])
    return out


# -----------------------------------------------------------------------------
# Escrita da planilha de saída
# -----------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
SUBHEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
LIGHT_FILL = PatternFill("solid", fgColor="F3F6FA")
YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")
GREEN_FILL = PatternFill("solid", fgColor="E2F0D9")
BORDER = Border(
    left=Side(style="thin", color="D9E2F3"),
    right=Side(style="thin", color="D9E2F3"),
    top=Side(style="thin", color="D9E2F3"),
    bottom=Side(style="thin", color="D9E2F3"),
)


def escrever_tabela(ws, headers: List[str], rows: List[Dict[str, Any]], start_row: int = 1, table_name: Optional[str] = None) -> int:
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(start_row, j, h)
        cell.fill = HEADER_FILL
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    for i, row in enumerate(rows, start=start_row + 1):
        for j, h in enumerate(headers, start=1):
            val = row.get(h, "")
            cell = ws.cell(i, j, val)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    end_row = start_row + len(rows)
    if table_name and rows:
        ref = f"A{start_row}:{get_column_letter(len(headers))}{end_row}"
        tab = Table(displayName=table_name, ref=ref)
        style = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws.add_table(tab)
    congelar = start_row + 1
    ws.freeze_panes = f"A{congelar}"
    return end_row


def ajustar_colunas(ws, max_width: int = 38) -> None:
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        width = 8
        for cell in col:
            if cell.value is None:
                continue
            width = max(width, min(max_width, len(str(cell.value)) + 2))
        ws.column_dimensions[col_letter].width = width
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 18


def aplicar_formatos_numericos(ws) -> None:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, float):
                if abs(cell.value) < 0.001 and cell.value != 0:
                    cell.number_format = "0.000E+00"
                else:
                    cell.number_format = "0.000"


def gerar_planilha(saida: Path, rsb: List[Dict[str, Any]], sft: List[Dict[str, Any]], alvo_sft: Dict[str, Any], grupos_sft: Dict[str, List[Dict[str, Any]]], memoria_rsb: List[Dict[str, Any]], params: Dict[str, Any]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "00_LEIA-ME"
    ws["A1"] = "Memória final de cálculo do acoplamento conceitual RSB-SFT"
    ws["A1"].font = Font(size=14, bold=True, color="1F4E79")
    linhas = [
        "Esta planilha foi gerada por script Python a partir de dois arquivos de resultados: RSB e SFT.",
        "A análise não executa novas simulações; apenas reorganiza, ranqueia e documenta os critérios de compatibilização.",
        "A condição-alvo preferencial da SFT é definida dentro de critérios editáveis: P >= 15 bar, H2/CO entre 1,90 e 2,10 e, por padrão, F = 15 SCCM.",
        "A RSB é avaliada pela aderência da razão H2/CO ao alvo da SFT, pelas conversões de CH4 e CO2 e pela seletividade relativa à formação de carbono sólido.",
        "Os cenários destacados não representam uma seleção artificial fechada; são os casos de maior aderência segundo os critérios documentados, dentro de todo o universo lido nas planilhas.",
        "O índice auxiliar usa pesos editáveis apenas para visualização. A leitura defensável principal é o ranking por delta H2/CO, conversões e seletividade a carbono.",
    ]
    for i, texto in enumerate(linhas, start=3):
        ws.cell(i, 1, texto)
        ws.cell(i, 1).alignment = Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 120

    # Parâmetros
    ws = wb.create_sheet("01_PARAMETROS")
    param_rows = []
    for k, v in params.items():
        param_rows.append({"Parametro": k, "Valor": v, "Descrição": descricao_parametro(k)})
    escrever_tabela(ws, ["Parametro", "Valor", "Descrição"], param_rows, 1, "TabelaParametros")
    ajustar_colunas(ws)

    # SFT longa
    ws = wb.create_sheet("02_SFT_LONGA")
    sft_headers = [
        "fonte_aba", "linha_bloco_inicio", "coluna_origem", "H2CO_SFT", "T_C", "P_bar", "F_SCCM",
        "W_Ftotal", "W_FCO_in", "XCO_pct", "XH2_pct", "alpha_aparente", "S_C5_C11_pct",
        "S_C12p_pct", "eta_C5_C11_pct", "Y_C5_C11_molar_pct", "OP_C5_C11"
    ]
    escrever_tabela(ws, sft_headers, sft, 1, "TabelaSFTLonga")
    ajustar_colunas(ws)
    aplicar_formatos_numericos(ws)

    # SFT alvo
    ws = wb.create_sheet("03_SFT_ALVO")
    ws["A1"] = "Condição-alvo adotada para a integração conceitual"
    ws["A1"].font = Font(size=13, bold=True, color="1F4E79")
    alvo_headers = ["critério", "H2CO_SFT", "T_C", "P_bar", "F_SCCM", "W_Ftotal", "W_FCO_in", "XCO_pct", "eta_C5_C11_pct", "alpha_aparente", "fonte_aba", "coluna_origem"]
    alvo_row = {h: alvo_sft.get(h, "") for h in alvo_headers}
    alvo_row["critério"] = "Alvo SFT preferencial"
    escrever_tabela(ws, alvo_headers, [alvo_row], 3, "TabelaAlvoSFT")
    r = 7
    for nome_grupo, linhas_grupo in grupos_sft.items():
        ws.cell(r, 1, nome_grupo).font = Font(bold=True, color="1F4E79")
        mini = []
        for item in linhas_grupo[:10]:
            mini.append({"critério": nome_grupo, **{h: item.get(h, "") for h in alvo_headers if h != "critério"}})
        r = escrever_tabela(ws, alvo_headers, mini, r + 1, None) + 3
    ajustar_colunas(ws)
    aplicar_formatos_numericos(ws)

    # RSB longa
    ws = wb.create_sheet("04_RSB_LONGA")
    rsb_headers = [
        "fonte_aba", "coluna_origem", "T_K", "T_C", "P_atm_aprox", "P_Pa", "F_biogas_mol_s", "CH4_frac", "CO2_frac",
        "H2CO_RSB", "XCH4_pct", "XCO2_pct", "X_media_pct", "SC_relativo", "CO_out_mol_s", "H2_out_mol_s",
        "C_out_mol_s", "DRM_mol_s", "RWGS_mol_s", "DM_mol_s"
    ]
    escrever_tabela(ws, rsb_headers, rsb, 1, "TabelaRSBLonga")
    ajustar_colunas(ws)
    aplicar_formatos_numericos(ws)

    # Memória RSB
    ws = wb.create_sheet("05_RSB_MEMORIA")
    memoria_headers = [
        "rank_primario", "rank_indice_auxiliar", "fonte_aba", "coluna_origem", "T_K", "P_atm_aprox", "CH4_frac", "CO2_frac",
        "H2CO_RSB", "H2CO_ALVO_SFT", "delta_H2CO_abs", "delta_H2CO_pct_relativo", "classe_aderencia",
        "XCH4_pct", "XCO2_pct", "X_media_pct", "SC_relativo", "score_H2CO_0a1", "score_conversao_0a1", "score_carbono_0a1",
        "indice_auxiliar_0a1", "leitura_condicionamento", "CO_out_mol_s", "H2_out_mol_s", "C_out_mol_s", "DRM_mol_s", "RWGS_mol_s", "DM_mol_s"
    ]
    escrever_tabela(ws, memoria_headers, memoria_rsb, 1, "TabelaRSBMemoria")
    ajustar_colunas(ws)
    aplicar_formatos_numericos(ws)

    # Candidatos destacados
    ws = wb.create_sheet("06_CENARIOS_DESTAQUE")
    top_n = int(params["RSB_TOP_N_DESTAQUE"])
    destaques = memoria_rsb[:top_n]
    escrever_tabela(ws, memoria_headers, destaques, 1, "TabelaCenariosDestaque")
    ajustar_colunas(ws)
    aplicar_formatos_numericos(ws)

    # Compatibilização
    ws = wb.create_sheet("07_COMPATIBILIZACAO")
    comp_headers = [
        "rank_primario", "T_RSB_K", "P_RSB_atm", "CH4_RSB", "CO2_RSB", "H2CO_RSB", "H2CO_ALVO_SFT", "delta_H2CO_abs",
        "classe_aderencia", "XCH4_pct", "XCO2_pct", "SC_relativo", "H2CO_SFT", "T_SFT_C", "P_SFT_bar", "F_SFT_SCCM",
        "eta_C5_C11_SFT_pct", "XCO_SFT_pct", "leitura_integracao"
    ]
    comp_rows = []
    for r in destaques:
        comp_rows.append({
            "rank_primario": r.get("rank_primario"),
            "T_RSB_K": r.get("T_K"),
            "P_RSB_atm": r.get("P_atm_aprox"),
            "CH4_RSB": r.get("CH4_frac"),
            "CO2_RSB": r.get("CO2_frac"),
            "H2CO_RSB": r.get("H2CO_RSB"),
            "H2CO_ALVO_SFT": r.get("H2CO_ALVO_SFT"),
            "delta_H2CO_abs": r.get("delta_H2CO_abs"),
            "classe_aderencia": r.get("classe_aderencia"),
            "XCH4_pct": r.get("XCH4_pct"),
            "XCO2_pct": r.get("XCO2_pct"),
            "SC_relativo": r.get("SC_relativo"),
            "H2CO_SFT": alvo_sft.get("H2CO_SFT"),
            "T_SFT_C": alvo_sft.get("T_C"),
            "P_SFT_bar": alvo_sft.get("P_bar"),
            "F_SFT_SCCM": alvo_sft.get("F_SCCM"),
            "eta_C5_C11_SFT_pct": alvo_sft.get("eta_C5_C11_pct"),
            "XCO_SFT_pct": alvo_sft.get("XCO_pct"),
            "leitura_integracao": r.get("leitura_condicionamento"),
        })
    escrever_tabela(ws, comp_headers, comp_rows, 1, "TabelaCompatibilizacao")
    ajustar_colunas(ws)
    aplicar_formatos_numericos(ws)

    # Gráficos editáveis
    ws = wb.create_sheet("08_GRAFICOS")
    ws["A1"] = "Gráficos editáveis gerados a partir dos cenários destacados"
    ws["A1"].font = Font(size=13, bold=True, color="1F4E79")
    chart_data_headers = ["rank", "H2CO_RSB", "delta_H2CO", "X_media_pct", "SC_relativo", "indice_auxiliar"]
    for j, h in enumerate(chart_data_headers, 1):
        ws.cell(3, j, h).fill = HEADER_FILL
        ws.cell(3, j).font = Font(color="FFFFFF", bold=True)
    for i, r in enumerate(destaques, 4):
        ws.cell(i, 1, r.get("rank_primario"))
        ws.cell(i, 2, r.get("H2CO_RSB"))
        ws.cell(i, 3, r.get("delta_H2CO_abs"))
        ws.cell(i, 4, r.get("X_media_pct"))
        ws.cell(i, 5, r.get("SC_relativo"))
        ws.cell(i, 6, r.get("indice_auxiliar_0a1"))
    nrows = len(destaques)
    if nrows:
        # Barras: delta H2/CO por ranking
        chart = BarChart()
        chart.title = "Aderência ao alvo da SFT: |H2/CO_RSB - H2/CO_alvo|"
        chart.y_axis.title = "Delta H2/CO"
        chart.x_axis.title = "Rank primário"
        data = Reference(ws, min_col=3, min_row=3, max_row=3 + nrows)
        cats = Reference(ws, min_col=1, min_row=4, max_row=3 + nrows)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 8
        chart.width = 18
        ws.add_chart(chart, "H3")

        # Linha: H2/CO RSB
        chart2 = LineChart()
        chart2.title = "H2/CO das correntes RSB destacadas"
        chart2.y_axis.title = "H2/CO"
        chart2.x_axis.title = "Rank primário"
        data2 = Reference(ws, min_col=2, min_row=3, max_row=3 + nrows)
        chart2.add_data(data2, titles_from_data=True)
        chart2.set_categories(cats)
        chart2.height = 8
        chart2.width = 18
        ws.add_chart(chart2, "H20")

        # Barras: conversão média e SC
        chart3 = BarChart()
        chart3.title = "Conversão média e seletividade relativa a carbono"
        chart3.y_axis.title = "Valor"
        data3 = Reference(ws, min_col=4, max_col=5, min_row=3, max_row=3 + nrows)
        chart3.add_data(data3, titles_from_data=True)
        chart3.set_categories(cats)
        chart3.height = 8
        chart3.width = 18
        ws.add_chart(chart3, "H37")
    ajustar_colunas(ws)
    aplicar_formatos_numericos(ws)

    # Aba com registro das fórmulas textuais
    ws = wb.create_sheet("09_FORMULAS_REGRAS")
    regras = [
        {"Item": "Alvo SFT", "Regra": "Selecionar cenário SFT em P >= SFT_P_MIN_BAR, H2/CO entre SFT_H2CO_MIN e SFT_H2CO_MAX e, se habilitado, F = SFT_F_SCCM_REFERENCIA; ordenar por maior eta_C5_C11."},
        {"Item": "Delta H2/CO", "Regra": "delta_H2CO_abs = ABS(H2CO_RSB - H2CO_ALVO_SFT)."},
        {"Item": "Classe de aderência", "Regra": "aderência estrita se delta <= RSB_TOLERANCIA_H2CO_ESTRITA; aderência ampliada se delta <= RSB_TOLERANCIA_H2CO_AMPLIADA."},
        {"Item": "Conversão média", "Regra": "X_media_pct = (XCH4_pct + XCO2_pct)/2."},
        {"Item": "Score H2/CO", "Regra": "score_H2CO_0a1 = 1 - MIN(delta_H2CO_abs/delta_max_observado, 1)."},
        {"Item": "Score conversão", "Regra": "score_conversao_0a1 = X_media_pct/100."},
        {"Item": "Score carbono", "Regra": "score_carbono_0a1 = 1 - (SC_relativo - SC_min)/(SC_max - SC_min). Menor SC é mais favorável."},
        {"Item": "Índice auxiliar", "Regra": "indice_auxiliar = PESO_H2CO*score_H2CO + PESO_CONVERSAO*score_conversao + PESO_CARBONO*score_carbono."},
        {"Item": "Ranking primário", "Regra": "Ordenar por menor delta_H2CO_abs; em seguida maior X_media_pct; em seguida menor SC_relativo."},
        {"Item": "Leitura", "Regra": "Os cenários destacados são representativos e não esgotam todas as possibilidades de acoplamento conceitual."},
    ]
    escrever_tabela(ws, ["Item", "Regra"], regras, 1, "TabelaFormulasRegras")
    ajustar_colunas(ws, max_width=90)

    # Metadados finais
    ws = wb.create_sheet("10_AUDITORIA")
    auditoria = [
        {"Campo": "Total de casos SFT lidos", "Valor": len(sft)},
        {"Campo": "Total de casos RSB lidos após deduplicação", "Valor": len(rsb)},
        {"Campo": "H2/CO alvo SFT", "Valor": alvo_sft.get("H2CO_SFT")},
        {"Campo": "T alvo SFT (°C)", "Valor": alvo_sft.get("T_C")},
        {"Campo": "P alvo SFT (bar)", "Valor": alvo_sft.get("P_bar")},
        {"Campo": "F alvo SFT (SCCM)", "Valor": alvo_sft.get("F_SCCM")},
        {"Campo": "eta_C5-C11 alvo SFT (%)", "Valor": alvo_sft.get("eta_C5_C11_pct")},
    ]
    escrever_tabela(ws, ["Campo", "Valor"], auditoria, 1, "TabelaAuditoria")
    ajustar_colunas(ws)
    aplicar_formatos_numericos(ws)

    # Ajustes gerais
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = True
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(saida)


def descricao_parametro(nome: str) -> str:
    descr = {
        "SFT_P_MIN_BAR": "Pressão mínima considerada coerente com o alvo de produtos líquidos na SFT.",
        "SFT_H2CO_MIN": "Limite inferior da faixa H2/CO usada para definir o alvo da SFT.",
        "SFT_H2CO_MAX": "Limite superior da faixa H2/CO usada para definir o alvo da SFT.",
        "SFT_F_SCCM_REFERENCIA": "Vazão de syngas usada como referência interna para a SFT.",
        "SFT_USAR_F_REFERENCIA": "Quando verdadeiro, restringe a escolha do alvo da SFT à vazão de referência.",
        "RSB_TOP_N_DESTAQUE": "Número de cenários RSB destacados na planilha de compatibilização.",
        "RSB_TOLERANCIA_H2CO_ESTRITA": "Tolerância para classificar aderência estrita ao alvo H2/CO.",
        "RSB_TOLERANCIA_H2CO_AMPLIADA": "Tolerância para classificar aderência ampliada ao alvo H2/CO.",
        "PESO_H2CO": "Peso do score de proximidade H2/CO no índice auxiliar.",
        "PESO_CONVERSAO": "Peso do score de conversão média no índice auxiliar.",
        "PESO_CARBONO": "Peso do score de menor seletividade relativa a carbono no índice auxiliar.",
    }
    return descr.get(nome, "Parâmetro da análise.")


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------

def localizar_arquivo(base_dir: Path, padrao: str, fallback_contains: str) -> Path:
    exato = base_dir / padrao
    if exato.exists():
        return exato
    candidatos = sorted(base_dir.glob("*.xlsx"))
    for c in candidatos:
        if fallback_contains.lower() in c.name.lower():
            return c
    raise FileNotFoundError(f"Não encontrei arquivo .xlsx contendo '{fallback_contains}' em {base_dir}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Gera planilha de memória final de cálculo para acoplamento conceitual RSB-SFT.")
    parser.add_argument("--rsb", type=str, default=None, help="Arquivo RSB_Results.xlsx")
    parser.add_argument("--sft", type=str, default=None, help="Arquivo RESULTADOS_Gray2024_...xlsx")
    parser.add_argument("--saida", type=str, default="Selecao_acoplamento_RSB_SFT.xlsx", help="Nome do arquivo .xlsx de saída")
    parser.add_argument("--top", type=int, default=None, help="Número de cenários destacados da RSB")
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    caminho_rsb = Path(args.rsb) if args.rsb else localizar_arquivo(base_dir, "RSB_Results.xlsx", "RSB")
    caminho_sft = Path(args.sft) if args.sft else localizar_arquivo(base_dir, "RESULTADOS_Gray2024_FeKSiO2_HTFT_H2CO_1.80a2.20_WF_P5a40.xlsx", "RESULTADOS_Gray")
    if not caminho_rsb.is_absolute():
        caminho_rsb = base_dir / caminho_rsb
    if not caminho_sft.is_absolute():
        caminho_sft = base_dir / caminho_sft
    params = dict(PARAMETROS)
    if args.top is not None:
        params["RSB_TOP_N_DESTAQUE"] = int(args.top)
    saida = Path(args.saida)
    if not saida.is_absolute():
        saida = base_dir / saida

    print("Lendo SFT:", caminho_sft.name)
    sft = extrair_sft(caminho_sft)
    if not sft:
        raise SystemExit("Nenhum caso SFT foi lido. Verifique o arquivo e o padrão de abas H2_CO.")
    print(f"  Casos SFT lidos: {len(sft)}")

    print("Lendo RSB:", caminho_rsb.name)
    rsb = extrair_rsb(caminho_rsb)
    if not rsb:
        raise SystemExit("Nenhum caso RSB foi lido. Verifique o arquivo e os rótulos das abas.")
    print(f"  Casos RSB lidos após deduplicação: {len(rsb)}")

    alvo_sft, grupos_sft = selecionar_alvo_sft(sft, params)
    print("Alvo SFT selecionado:", {
        "H2CO": alvo_sft.get("H2CO_SFT"),
        "T_C": alvo_sft.get("T_C"),
        "P_bar": alvo_sft.get("P_bar"),
        "F_SCCM": alvo_sft.get("F_SCCM"),
        "eta_C5_C11_pct": alvo_sft.get("eta_C5_C11_pct"),
    })

    memoria_rsb = calcular_memoria_rsb(rsb, alvo_sft, params)
    gerar_planilha(saida, rsb, sft, alvo_sft, grupos_sft, memoria_rsb, params)
    print("Planilha gerada:", saida)


if __name__ == "__main__":
    main()
