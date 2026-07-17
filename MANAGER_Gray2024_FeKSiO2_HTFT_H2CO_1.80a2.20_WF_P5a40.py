# -*- coding: utf-8 -*-
# =========================================================
# MANAGER_Gray2024_FeKSiO2_HTFT_H2CO_1.80a2.20_WF_P5a40.py
#
# Rotina de automacao para o DWSIM usada na etapa de SFT.
#
# Finalidade:
#   Executar varreduras parametricas da sintese de Fischer-Tropsch
#   em regime HTFT, com catalisador Fe-K/SiO2 e modelo cinetico
#   baseado em Gray et al. (2024), utilizando o CUSTOM:
#
#   CUSTOM_Gray2024_FeKSiO2_HTFT_WGS_CHO_lumpC12plus.py
#
# Escopo da varredura:
#   - razao molar H2/CO de 1.80 a 2.20;
#   - temperaturas de 290 a 360 C;
#   - pressoes de 5 a 40 bar;
#   - cinco vazoes molares equivalentes a 3.75, 7.50, 15, 30 e 60 SCCM;
#   - Wcat fixo no CUSTOM em 0.001 kg.
#
# Interpretacao fisica:
#   A variacao da vazao molar, mantendo Wcat constante, permite avaliar
#   o efeito de W/Ftotal e W/FCO,in. Assim, os resultados nao devem ser
#   interpretados como efeito isolado de massa catalitica, mas como efeito
#   da quantidade de catalisador disponivel por unidade de vazao molar
#   alimentada, especialmente por unidade de CO de entrada.
#
# Organizacao da saida:
#   - uma aba por razao H2/CO;
#   - dentro de cada aba, um bloco vertical por vazao de syngas;
#   - dentro de cada bloco, pressoes em blocos horizontais;
#   - temperaturas dispostas nas colunas de cada bloco de pressao.
#
# Arquivos gerados:
#   RESULTADOS_Gray2024_FeKSiO2_HTFT_H2CO_1.80a2.20_WF_P5a40.xlsx
#   e arquivos CSV de backup, um por aba.
#
# Observacao:
#   Este script nao modifica as equacoes cineticas. Ele apenas define
#   as condicoes de alimentacao, executa o flowsheet e organiza os
#   resultados calculados pelo CUSTOM e pelo DWSIM.
# =========================================================

import clr
import math
import os
import traceback

clr.AddReference("System")
import System
from System import Environment
from System.IO import Path, Directory, StreamWriter

# =========================================================
# CONFIGURACAO DO FLOWSHEET
# =========================================================
FEED_STREAM_NAME = "SYNGAS_FEED"
PRODUCT_STREAM_NAME = "FT_OUT_NEW"
CUSTOM_UO_NAME = "CUSTOM-FT-NEW"

# Deve coincidir com o valor de Wcat_total definido no CUSTOM usado na simulacao.
WCAT_TOTAL_KG = 0.001

# Conversao adotada: 15 SCCM = 1.02186E-05 mol/s.
MOL_S_PER_SCCM = 1.02186e-5 / 15.0

# Vazoes usadas para avaliar W/Ftotal e W/FCO,in com Wcat fixo.
FLOW_CASES = [
    (3.75,  3.75  * MOL_S_PER_SCCM),
    (7.50,  7.50  * MOL_S_PER_SCCM),
    (15.00, 15.00 * MOL_S_PER_SCCM),
    (30.00, 30.00 * MOL_S_PER_SCCM),
    (60.00, 60.00 * MOL_S_PER_SCCM),
]

# Faixa HTFT avaliada. O ponto de 360 C aproxima a varredura da faixa experimental de Gray et al. (2024).
TEMPERATURES_C = [290.0, 300.0, 310.0, 320.0, 330.0, 340.0, 350.0, 360.0]
TEMPERATURES_K = [t + 273.15 for t in TEMPERATURES_C]

# Faixa de pressao avaliada: 5-40 bar.
PRESSURES_BAR = [5.0, 10.0, 15.0, 20.0, 30.0, 40.0]
PRESSURES_PA = [p * 100000.0 for p in PRESSURES_BAR]

# Razoes H2/CO avaliadas em torno da faixa de interesse para a SFT.
H2CO_VALUES = [1.80, 1.85, 1.90, 1.95, 2.00, 2.05, 2.10, 2.15, 2.20]

# Arquivos de saida.
OUTPUT_BASENAME = "RESULTADOS_Gray2024_FeKSiO2_HTFT_H2CO_1.80a2.20_WF_P5a40"
OUTPUT_ROOT = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), OUTPUT_BASENAME)
CSV_DIR = Path.Combine(OUTPUT_ROOT, "CSV_backup")

# Para evitar perda de resultados, CSV e gravado sempre.
WRITE_CSV_BACKUP = True
# XLSX usa automacao do Excel no final. Se travar, os CSVs ja estarao salvos.
CREATE_XLSX_WITH_EXCEL = True
# Escrita em bloco no Excel; se falhar, cai para celula a celula.
USE_BULK_EXCEL_WRITE = True

# =========================================================
# COMPONENTES E ATOMOS
# =========================================================
# Nome DWSIM, rotulo, C, H, O, familia.
# familia: "par", "olef" ou "base". A classificacao e usada nos cortes C2-C4, C5-C11 e C12+.
COMPONENT_ROWS = [
    ("Carbon monoxide", "FCO,OUT (mol/s)", 1, 0, 1, "base"),
    ("Hydrogen", "H2 (mol/s)", 0, 2, 0, "base"),
    ("Water", "H2O (mol/s)", 0, 2, 1, "base"),
    ("Carbon dioxide", "CO2 (mol/s)", 1, 0, 2, "base"),
    ("Methane", "CH4 (mol/s)", 1, 4, 0, "par"),
    ("Ethylene", "C2H4 (mol/s)", 2, 4, 0, "olef"),
    ("Ethane", "C2H6 (mol/s)", 2, 6, 0, "par"),
    ("Propylene", "C3H6 (mol/s)", 3, 6, 0, "olef"),
    ("Propane", "C3H8 (mol/s)", 3, 8, 0, "par"),
    ("1-butene", "C4H8 (mol/s)", 4, 8, 0, "olef"),
    ("N-butane", "C4H10 (mol/s)", 4, 10, 0, "par"),
    ("1-pentene", "C5H10 (mol/s)", 5, 10, 0, "olef"),
    ("N-pentane", "C5H12 (mol/s)", 5, 12, 0, "par"),
    ("1-hexene", "C6H12 (mol/s)", 6, 12, 0, "olef"),
    ("N-hexane", "C6H14 (mol/s)", 6, 14, 0, "par"),
    ("1-heptene", "C7H14 (mol/s)", 7, 14, 0, "olef"),
    ("N-heptane", "C7H16 (mol/s)", 7, 16, 0, "par"),
    ("1-octene", "C8H16 (mol/s)", 8, 16, 0, "olef"),
    ("N-octane", "C8H18 (mol/s)", 8, 18, 0, "par"),
    ("1-nonene", "C9H18 (mol/s)", 9, 18, 0, "olef"),
    ("N-nonane", "C9H20 (mol/s)", 9, 20, 0, "par"),
    ("1-decene", "C10H20 (mol/s)", 10, 20, 0, "olef"),
    ("N-decane", "C10H22 (mol/s)", 10, 22, 0, "par"),
    ("1-undecene", "C11H22 (mol/s)", 11, 22, 0, "olef"),
    ("N-undecane", "C11H24 (mol/s)", 11, 24, 0, "par"),
    ("1-dodecene", "C12H24 lump olef. (mol/s)", 12, 24, 0, "olef"),
    ("N-dodecane", "C12H26 lump paraf. (mol/s)", 12, 26, 0, "par"),
    ("1-eicosene", "C20H40 lump olef. (mol/s)", 20, 40, 0, "olef"),
    ("N-tetracosane", "C24H50 lump paraf. (mol/s)", 24, 50, 0, "par"),
]

# =========================================================
# FUNCOES AUXILIARES
# =========================================================
def safe_div(a, b):
    # Divisao protegida para evitar erro numerico quando o denominador se aproxima de zero.
    if abs(b) <= 1.0e-30:
        return 0.0
    return a / b


def fmt_num(x, nd=2):
    return ("%.*f" % (nd, x)).replace(".", ",")


def sheet_name_for_h2co(x):
    return "H2_CO " + fmt_num(x, 2)


def safe_sheet_filename(sheet_name):
    s = sheet_name.replace("/", "_").replace(" ", "_").replace(",", "_")
    s = s.replace(":", "_").replace("*", "_").replace("?", "_")
    s = s.replace("[", "_").replace("]", "_")
    return s


def ensure_dir(path):
    if not Directory.Exists(path):
        Directory.CreateDirectory(path)


def get_object(tag):
    try:
        return Flowsheet.GetFlowsheetSimulationObject(tag)
    except:
        pass
    try:
        return Flowsheet.SimulationObjects[tag]
    except:
        pass
    try:
        for obj in Flowsheet.SimulationObjects.Values:
            try:
                if str(obj.GraphicObject.Tag) == tag:
                    return obj
            except:
                pass
            try:
                if str(obj.Name) == tag:
                    return obj
            except:
                pass
    except:
        pass
    raise Exception("Objeto nao encontrado no flowsheet: " + tag)


def get_component_names(stream):
    names = []
    try:
        for k in stream.Phases[0].Compounds.Keys:
            names.append(str(k))
        if len(names) > 0:
            return names
    except:
        pass
    try:
        for k in stream.GetPhase("Overall").Compounds.Keys:
            names.append(str(k))
        if len(names) > 0:
            return names
    except:
        pass
    raise Exception("Nao foi possivel obter os nomes dos compostos.")


def set_stream_case(feed, comp_idx, T_K, P_Pa, flow_mol_s, h2co):
    # Define a corrente de alimentacao SYNGAS_FEED para um caso da varredura.
    # A alimentacao reacional contem apenas CO e H2; os demais componentes sao zerados.
    # Para uma razao R = H2/CO: z_CO = 1/(1+R) e z_H2 = R/(1+R).
    z_co = 1.0 / (1.0 + h2co)
    z_h2 = h2co / (1.0 + h2co)

    try:
        feed.SetTemperature(float(T_K))
    except:
        feed.GetPhase("Overall").Properties.temperature = float(T_K)

    try:
        feed.SetPressure(float(P_Pa))
    except:
        feed.GetPhase("Overall").Properties.pressure = float(P_Pa)

    try:
        feed.SetMolarFlow(float(flow_mol_s))
    except:
        feed.GetPhase("Overall").Properties.molarflow = float(flow_mol_s)

    ncomp = len(comp_idx)
    try:
        for i in range(ncomp):
            feed.SetOverallCompoundMolarFlow(i, 0.0)
        feed.SetOverallCompoundMolarFlow(comp_idx["Carbon monoxide"], flow_mol_s * z_co)
        feed.SetOverallCompoundMolarFlow(comp_idx["Hydrogen"], flow_mol_s * z_h2)
    except:
        from System import Array, Double
        z = [0.0 for i in range(ncomp)]
        z[comp_idx["Carbon monoxide"]] = z_co
        z[comp_idx["Hydrogen"]] = z_h2
        feed.SetOverallComposition(Array[Double](z))

    try:
        feed.Calculate()
    except:
        pass

    return z_h2, z_co, flow_mol_s * z_h2, flow_mol_s * z_co


def solve_case(custom_uo, product):
    # Tenta resolver o flowsheet completo; se nao for possivel, calcula diretamente
    # a unidade CUSTOM e a corrente de produto. Essa redundancia torna o script
    # mais robusto em diferentes configuracoes do DWSIM.
    try:
        Solver.SolveFlowsheet(Flowsheet, True)
        return
    except:
        pass
    try:
        Solver.SolveFlowsheet(Flowsheet)
        return
    except:
        pass
    try:
        custom_uo.Calculate()
    except:
        pass
    try:
        product.Calculate()
    except:
        pass


def get_molar_flows(stream, comp_names):
    # Converte composicoes molares globais da corrente de produto em vazoes molares
    # individuais, que serao usadas no calculo dos cortes e indicadores.
    Ftot = float(stream.GetMolarFlow())
    z = list(stream.GetOverallComposition())
    flows = {}
    for i, name in enumerate(comp_names):
        try:
            flows[name] = Ftot * float(z[i])
        except:
            flows[name] = 0.0
    return flows


def get_flow(flows, name):
    try:
        return float(flows[name])
    except:
        return 0.0


def atoms_out(flows, atom_index):
    # atom_index: 2=C, 3=H, 4=O em COMPONENT_ROWS.
    total = 0.0
    for row in COMPONENT_ROWS:
        name = row[0]
        natoms = row[atom_index]
        total += natoms * get_flow(flows, name)
    return total


def hydrocarbon_carbon(flows, n_min, n_max):
    # Soma o fluxo de carbono, e nao apenas mols de moleculas, em uma faixa Cn-Cm.
    total_c = 0.0
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        if family in ["par", "olef"] and nC >= n_min and nC <= n_max:
            total_c += nC * get_flow(flows, name)
    return total_c


def hydrocarbon_moles(flows, n_min, n_max):
    # Soma as vazoes molares das moleculas em uma faixa Cn-Cm.
    total_m = 0.0
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        if family in ["par", "olef"] and nC >= n_min and nC <= n_max:
            total_m += get_flow(flows, name)
    return total_m


def hydrocarbon_carbon_total(flows):
    total_c = 0.0
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        if family in ["par", "olef"]:
            total_c += nC * get_flow(flows, name)
    return total_c


def hydrocarbon_moles_total(flows):
    total_m = 0.0
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        if family in ["par", "olef"]:
            total_m += get_flow(flows, name)
    return total_m


def carbon_by_number(flows, n):
    total = 0.0
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        if family in ["par", "olef"] and nC == n:
            total += get_flow(flows, name)
    return total


def alpha_apparent_C5_C11(flows):
    ratios = []
    for n in range(6, 12):
        f_prev = carbon_by_number(flows, n - 1)
        f_now = carbon_by_number(flows, n)
        if f_prev > 1.0e-30 and f_now >= 0.0:
            ratios.append(f_now / f_prev)
    if len(ratios) == 0:
        return 0.0
    return sum(ratios) / float(len(ratios))


def olefin_paraffin_ratio_C5_C11(flows):
    olef = 0.0
    par = 0.0
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        if nC >= 5 and nC <= 11:
            if family == "olef":
                olef += get_flow(flows, name)
            elif family == "par":
                par += get_flow(flows, name)
    return safe_div(olef, par)


def build_results(flows, F_in, FH2_in, FCO_in, h2co, T_K, P_Pa):
    # Calcula indicadores de desempenho a partir das vazoes molares de saida.
    # Os indicadores de carbono usam FCO,in como base, pois o carbono dos
    # hidrocarbonetos formados deriva do CO alimentado no modelo da SFT.
    FCO_out = get_flow(flows, "Carbon monoxide")
    FH2_out = get_flow(flows, "Hydrogen")
    FCO2_out = get_flow(flows, "Carbon dioxide")

    CO_conv = FCO_in - FCO_out
    H2_conv = FH2_in - FH2_out

    C_HC_total = hydrocarbon_carbon_total(flows)
    M_HC_total = hydrocarbon_moles_total(flows)
    C_CH4 = get_flow(flows, "Methane")
    C_C2_C4 = hydrocarbon_carbon(flows, 2, 4)
    C_C5_C11 = hydrocarbon_carbon(flows, 5, 11)
    C_C12plus = hydrocarbon_carbon(flows, 12, 100)
    M_C5_C11 = hydrocarbon_moles(flows, 5, 11)

    C_in = FCO_in
    H_in = 2.0 * FH2_in
    O_in = FCO_in

    C_out = atoms_out(flows, 2)
    H_out = atoms_out(flows, 3)
    O_out = atoms_out(flows, 4)

    rows = {}
    rows["XCO (%)"] = safe_div(CO_conv, FCO_in) * 100.0
    rows["XH2 (%)"] = safe_div(H2_conv, FH2_in) * 100.0
    rows["alpha aparente C5-C11"] = alpha_apparent_C5_C11(flows)
    rows["S CH4 (%)"] = safe_div(C_CH4, C_HC_total) * 100.0
    rows["S C2-C4 (%)"] = safe_div(C_C2_C4, C_HC_total) * 100.0
    rows["S C5-C11 (%)"] = safe_div(C_C5_C11, C_HC_total) * 100.0
    rows["S C12+ (%)"] = safe_div(C_C12plus, C_HC_total) * 100.0
    rows["S CO2 (%)"] = safe_div(FCO2_out, CO_conv) * 100.0
    rows["eta_C C5-C11 (%)"] = safe_div(C_C5_C11, FCO_in) * 100.0
    rows["Y C5-C11 molar (%)"] = safe_div(M_C5_C11, FCO_in) * 100.0
    rows["H2/CO saida"] = safe_div(FH2_out, FCO_out)
    rows["olefinas/parafinas C5-C11"] = olefin_paraffin_ratio_C5_C11(flows)
    rows["balanco C erro (%)"] = safe_div(C_out - C_in, C_in) * 100.0
    rows["balanco H erro (%)"] = safe_div(H_out - H_in, H_in) * 100.0
    rows["balanco O erro (%)"] = safe_div(O_out - O_in, O_in) * 100.0
    rows["CTOT em HC (mol C/s)"] = C_HC_total
    rows["MTOT HC (mol/s)"] = M_HC_total
    return rows

# =========================================================
# MATRIZ DE SAIDA
# =========================================================
def empty_matrix(rows, cols):
    return [[None for c in range(cols)] for r in range(rows)]


def set_cell(mat, r, c, value):
    while len(mat) < r:
        mat.append([])
    while len(mat[r - 1]) < c:
        mat[r - 1].append(None)
    mat[r - 1][c - 1] = value


def block_start_col(pressure_index):
    # Comeca em E; deixa uma coluna em branco entre pressoes.
    return 5 + pressure_index * (len(TEMPERATURES_K) + 1)


BLOCK_ROWS = 72
BASE_BLOCK_DATA_START = 1


def flow_block_start(flow_index):
    return 1 + flow_index * BLOCK_ROWS


def put_block_labels(mat, row0, h2co, flow_sccm, flow_mol_s):
    # Insere, em cada bloco de vazao, os cabecalhos fixos da planilha.
    # row0 e 1-based.
    z_co = 1.0 / (1.0 + h2co)
    z_h2 = h2co / (1.0 + h2co)
    fh2_in = flow_mol_s * z_h2
    fco_in = flow_mol_s * z_co

    set_cell(mat, row0 + 0, 1, "SFT Fe-K/SiO2 - Gray2024 HTFT | F = %.2f SCCM" % flow_sccm)
    set_cell(mat, row0 + 1, 3, "Temperatura (K) =")
    set_cell(mat, row0 + 2, 3, "Temperatura (C) =")
    set_cell(mat, row0 + 3, 3, "Pressao (Pa) =")
    set_cell(mat, row0 + 4, 3, "Pressao (bar) =")
    set_cell(mat, row0 + 5, 3, "SYNGAS Flow IN (SCCM) =")
    set_cell(mat, row0 + 6, 3, "SYNGAS Flow IN (mol/s) =")
    set_cell(mat, row0 + 8, 3, "Basis: MOLAR FRACTION")
    set_cell(mat, row0 + 9, 3, "H2/CO =")
    set_cell(mat, row0 + 10, 3, "H2 (fracao molar) =")
    set_cell(mat, row0 + 11, 3, "CO (fracao molar) =")
    set_cell(mat, row0 + 12, 3, "FH2,IN (mol/s) =")
    set_cell(mat, row0 + 13, 3, "FCO,IN (mol/s) =")
    set_cell(mat, row0 + 15, 3, "Wcat_total (kg) =")
    set_cell(mat, row0 + 16, 3, "W/F total (kg/(mol/s)) =")
    set_cell(mat, row0 + 17, 3, "W/FCO,in (kg/(mol CO/s)) =")
    set_cell(mat, row0 + 19, 3, "Basis: MOLAR FLOW")

    # Letras laterais apenas organizam visualmente os blocos da planilha.
    set_cell(mat, row0 + 3, 2, "F")
    set_cell(mat, row0 + 4, 2, "E")
    set_cell(mat, row0 + 5, 2, "E")
    set_cell(mat, row0 + 6, 2, "D")
    set_cell(mat, row0 + 10, 2, "IN")
    set_cell(mat, row0 + 24, 2, "P")
    set_cell(mat, row0 + 26, 2, "R")
    set_cell(mat, row0 + 28, 2, "O")
    set_cell(mat, row0 + 30, 2, "D")
    set_cell(mat, row0 + 32, 2, "U")
    set_cell(mat, row0 + 34, 2, "T")
    set_cell(mat, row0 + 36, 2, "O")
    set_cell(mat, row0 + 38, 2, "S")

    r = row0 + 20
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        set_cell(mat, r, 3, label)
        if family in ["par", "olef"]:
            set_cell(mat, r, 4, nC)
        r += 1

    r += 1
    set_cell(mat, r, 3, "CTOT em HC (mol C/s)"); r += 1
    set_cell(mat, r, 3, "MTOT HC (mol/s)"); r += 1
    set_cell(mat, r, 3, "Indices de DESEMPENHO"); r += 1

    metric_labels = [
        "XCO (%)",
        "XH2 (%)",
        "alpha aparente C5-C11",
        "S CH4 (%)",
        "S C2-C4 (%)",
        "S C5-C11 (%)",
        "S C12+ (%)",
        "S CO2 (%)",
        "eta_C C5-C11 (%)",
        "Y C5-C11 molar (%)",
        "H2/CO saida",
        "olefinas/parafinas C5-C11",
        "balanco C erro (%)",
        "balanco H erro (%)",
        "balanco O erro (%)",
        "Status",
    ]
    for lab in metric_labels:
        set_cell(mat, r, 3, lab)
        r += 1

    # Cabeçalhos de cada bloco de pressao.
    for p_i, P in enumerate(PRESSURES_PA):
        c0 = block_start_col(p_i)
        for t_i, T in enumerate(TEMPERATURES_K):
            c = c0 + t_i
            set_cell(mat, row0 + 1, c, T)
            set_cell(mat, row0 + 2, c, TEMPERATURES_C[t_i])
            set_cell(mat, row0 + 3, c, P)
            set_cell(mat, row0 + 4, c, PRESSURES_BAR[p_i])
            set_cell(mat, row0 + 5, c, flow_sccm)
            set_cell(mat, row0 + 6, c, flow_mol_s)
            set_cell(mat, row0 + 9, c, h2co)
            set_cell(mat, row0 + 10, c, z_h2)
            set_cell(mat, row0 + 11, c, z_co)
            set_cell(mat, row0 + 12, c, fh2_in)
            set_cell(mat, row0 + 13, c, fco_in)
            set_cell(mat, row0 + 15, c, WCAT_TOTAL_KG)
            set_cell(mat, row0 + 16, c, safe_div(WCAT_TOTAL_KG, flow_mol_s))
            set_cell(mat, row0 + 17, c, safe_div(WCAT_TOTAL_KG, fco_in))


def create_sheet_matrix(h2co):
    ncols = 4 + len(PRESSURES_PA) * (len(TEMPERATURES_K) + 1)
    nrows = len(FLOW_CASES) * BLOCK_ROWS
    mat = empty_matrix(nrows, ncols)
    for f_i, case in enumerate(FLOW_CASES):
        flow_sccm, flow_mol_s = case
        row0 = flow_block_start(f_i)
        put_block_labels(mat, row0, h2co, flow_sccm, flow_mol_s)
    return mat


def output_row_for_components(row0):
    return row0 + 20


def metrics_start_row(row0):
    return row0 + 20 + len(COMPONENT_ROWS) + 3


def fill_case_results(mat, row0, col, flows, metrics):
    # Escreve na matriz da planilha as vazoes dos componentes e os indicadores
    # calculados para um caso especifico de T, P, H2/CO e vazao.
    r = output_row_for_components(row0)
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        set_cell(mat, r, col, get_flow(flows, name))
        r += 1

    r += 1
    set_cell(mat, r, col, metrics["CTOT em HC (mol C/s)"]); r += 1
    set_cell(mat, r, col, metrics["MTOT HC (mol/s)"]); r += 1
    r += 1

    metric_labels = [
        "XCO (%)",
        "XH2 (%)",
        "alpha aparente C5-C11",
        "S CH4 (%)",
        "S C2-C4 (%)",
        "S C5-C11 (%)",
        "S C12+ (%)",
        "S CO2 (%)",
        "eta_C C5-C11 (%)",
        "Y C5-C11 molar (%)",
        "H2/CO saida",
        "olefinas/parafinas C5-C11",
        "balanco C erro (%)",
        "balanco H erro (%)",
        "balanco O erro (%)",
    ]
    for lab in metric_labels:
        set_cell(mat, r, col, metrics[lab])
        r += 1
    set_cell(mat, r, col, "OK")


def fill_case_error(mat, row0, col, message):
    status_row = metrics_start_row(row0) + 15
    set_cell(mat, status_row, col, "ERRO: " + str(message)[0:200])

# =========================================================
# SAIDAS CSV E XLSX
# =========================================================
def write_csv_sheet(sheet_name, mat):
    # Gera backup em CSV para cada aba. Esse arquivo e util caso a automacao
    # do Excel nao consiga gerar o XLSX ao final da varredura.
    ensure_dir(CSV_DIR)
    path = Path.Combine(CSV_DIR, OUTPUT_BASENAME + "_" + safe_sheet_filename(sheet_name) + ".csv")
    sw = StreamWriter(path, False, System.Text.Encoding.UTF8)
    try:
        for row in mat:
            cells = []
            for val in row:
                if val is None:
                    cells.append("")
                else:
                    s = str(val).replace(";", ",")
                    cells.append(s)
            sw.WriteLine(";".join(cells))
    finally:
        sw.Close()
    return path


def excel_write_matrix(ws, mat):
    nrows = len(mat)
    ncols = max([len(r) for r in mat])

    if USE_BULK_EXCEL_WRITE:
        try:
            arr = System.Array.CreateInstance(System.Object, nrows, ncols)
            for r in range(nrows):
                for c in range(ncols):
                    val = None
                    if c < len(mat[r]):
                        val = mat[r][c]
                    if val is not None:
                        arr.SetValue(val, r, c)
            rng = ws.Range(ws.Cells(1, 1), ws.Cells(nrows, ncols))
            rng.Value2 = arr
            return
        except Exception as e:
            print("Aviso: escrita em bloco no Excel falhou; usando celula a celula. Motivo: " + str(e))

    for r in range(nrows):
        for c in range(ncols):
            val = None
            if c < len(mat[r]):
                val = mat[r][c]
            if val is not None:
                ws.Cells(r + 1, c + 1).Value2 = val


def write_xlsx_with_excel(all_sheets):
    # Consolida todas as matrizes em um arquivo XLSX usando a automacao COM do Excel.
    # Em ambiente sem Excel instalado, os CSVs de backup permanecem como saida principal.
    clr.AddReference("Microsoft.Office.Interop.Excel")
    import Microsoft.Office.Interop.Excel as Excel

    ensure_dir(OUTPUT_ROOT)

    excel = Excel.ApplicationClass()
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Add()

    while wb.Worksheets.Count < len(all_sheets):
        wb.Worksheets.Add()

    sheet_index = 1
    for sheet_name, mat in all_sheets:
        ws = wb.Worksheets[sheet_index]
        ws.Name = sheet_name
        excel_write_matrix(ws, mat)

        try:
            ws.Rows(1).Font.Bold = True
            ws.Columns.AutoFit()
            ws.Range("A1:A1").Font.Bold = True
        except:
            pass
        sheet_index += 1

    while wb.Worksheets.Count > len(all_sheets):
        wb.Worksheets[wb.Worksheets.Count].Delete()

    output_path = Path.Combine(OUTPUT_ROOT, OUTPUT_BASENAME + ".xlsx")
    wb.SaveAs(output_path)
    wb.Close(True)
    excel.Quit()
    return output_path

# =========================================================
# EXECUCAO PRINCIPAL
# =========================================================
def main():
    # Executa a varredura completa. A ordem dos loops segue a organizacao da saida:
    # H2/CO define a aba; vazao define o bloco vertical; pressao e temperatura
    # definem as colunas internas de cada bloco.
    ensure_dir(OUTPUT_ROOT)
    ensure_dir(CSV_DIR)

    feed = get_object(FEED_STREAM_NAME)
    prod = get_object(PRODUCT_STREAM_NAME)
    custom = get_object(CUSTOM_UO_NAME)

    comp_names = get_component_names(feed)
    comp_idx = {}
    for i, name in enumerate(comp_names):
        comp_idx[name] = i

    required = ["Carbon monoxide", "Hydrogen", "Water", "Carbon dioxide"]
    for name, label, nC, nH, nO, family in COMPONENT_ROWS:
        if name not in required:
            required.append(name)

    missing = []
    for name in required:
        if name not in comp_idx:
            missing.append(name)
    if len(missing) > 0:
        raise Exception("Faltam componentes no flowsheet: " + "; ".join(missing))

    all_sheets = []
    total_cases = 0

    for h2co in H2CO_VALUES:
        sheet_name = sheet_name_for_h2co(h2co)
        mat = create_sheet_matrix(h2co)

        for f_i, case in enumerate(FLOW_CASES):
            flow_sccm, flow_mol_s = case
            row0 = flow_block_start(f_i)
            z_co = 1.0 / (1.0 + h2co)
            z_h2 = h2co / (1.0 + h2co)
            fh2_in = flow_mol_s * z_h2
            fco_in = flow_mol_s * z_co

            for p_i, P in enumerate(PRESSURES_PA):
                c0 = block_start_col(p_i)
                for t_i, T in enumerate(TEMPERATURES_K):
                    col = c0 + t_i
                    total_cases += 1
                    try:
                        set_stream_case(feed, comp_idx, T, P, flow_mol_s, h2co)
                        solve_case(custom, prod)
                        flows = get_molar_flows(prod, comp_names)
                        metrics = build_results(flows, flow_mol_s, fh2_in, fco_in, h2co, T, P)
                        fill_case_results(mat, row0, col, flows, metrics)
                        print("OK - F=%.2f SCCM  H2/CO=%.2f  T=%.2f K  P=%.0f Pa" % (flow_sccm, h2co, T, P))
                    except Exception as e:
                        msg = str(e)
                        fill_case_error(mat, row0, col, msg)
                        print("ERRO - F=%.2f SCCM  H2/CO=%.2f  T=%.2f K  P=%.0f Pa: %s" % (flow_sccm, h2co, T, P, msg))
                        print(traceback.format_exc())

        # Backup por aba imediatamente apos completar cada razao H2/CO.
        if WRITE_CSV_BACKUP:
            p_csv = write_csv_sheet(sheet_name, mat)
            print("CSV backup gerado: " + str(p_csv))

        all_sheets.append((sheet_name, mat))

    if CREATE_XLSX_WITH_EXCEL:
        try:
            out = write_xlsx_with_excel(all_sheets)
            print("Planilha XLSX gerada: " + str(out))
        except Exception as e:
            print("XLSX nao gerado. Os CSVs de backup permanecem salvos. Motivo: " + str(e))

    print("Casos tentados: " + str(total_cases))
    print("Pasta de saida: " + str(OUTPUT_ROOT))

main()
