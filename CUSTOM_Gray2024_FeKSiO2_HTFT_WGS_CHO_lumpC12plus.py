# -*- coding: utf-8 -*-
import clr
import math
import os
clr.AddReference("System")
import System

feed = ims1
prod = oms1

# No DWSIM, ims1 e oms1 sao as correntes de entrada e saida da
# unidade CUSTOM, respectivamente. O script le T, P, vazao molar
# total e composicao da corrente feed, integra o modelo cinetico
# ao longo da massa catalitica W e escreve a composicao calculada
# na corrente prod.

# =========================================================
# CUSTOM_Gray2024_FeKSiO2_HTFT_WGS_CHO_lumpC12plus.py
#
# Finalidade:
#   Representar a sintese de Fischer-Tropsch em regime HTFT com
#   catalisador Fe-K/SiO2, conforme a estrutura cinetica de
#   Gray et al. (2024), em uma unidade CUSTOM do DWSIM.
#
# Escopo do modelo:
#   - Reator pseudo-homogeneo do tipo PFR, integrado em funcao
#     da massa de catalisador W.
#   - Formacao de parafinas e olefinas segundo distribuicao de
#     crescimento de cadeia.
#   - Reacao WGS acoplada ao consumo/formacao de CO, H2, H2O e CO2.
#   - Representacao explicita de C1-C11 e agrupamento das fracoes
#     C12+ em compostos representativos disponiveis no DWSIM.
#   - Limitador estequiometrico conjunto para preservar a coerencia
#     material em cada passo de integracao.
#   - Registro diagnostico dos balancos elementares de C, H e O.
#
# Observacao:
#   O script calcula distribuicoes composicionais e vazoes molares
#   no efluente. A identificacao de uma faixa C5-C11 nao equivale,
#   por si so, a comprovacao de fase liquida no efluente.
# =========================================================

# =========================================================
# MAPA DE COMPONENTES POR NOME
# =========================================================
# A busca dos componentes e feita por nome, e nao por indice fixo.
# Isso reduz o risco de erro quando a ordem dos compostos no DWSIM
# e alterada pela inclusao ou remocao de especies no flowsheet.
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

    raise Exception("Nao foi possivel obter os compostos da corrente de entrada.")

component_names = get_component_names(feed)
z = list(feed.GetOverallComposition())

if len(component_names) != len(z):
    raise Exception(
        "Numero de nomes de compostos diferente do vetor de composicao: "
        + str(len(component_names)) + " nomes e "
        + str(len(z)) + " fracoes."
    )

idx = {}
for i, name in enumerate(component_names):
    idx[name] = i

# =========================================================
# COMPONENTES UTILIZADOS
#
# C1-C11: especies explicitas.
# C12+: dois lumps para parafinas e dois para olefinas.
#
# Parafinas C12+ -> N-dodecane (C12) + N-tetracosane (C24)
# Olefinas C12+  -> 1-dodecene (C12) + 1-eicosene (C20)
#
# Na faixa representavel, o fracionamento entre os dois representantes
# preserva simultaneamente:
#   1) vazao molar total da familia;
#   2) carbono total da familia.
#
# Fora da faixa representavel, o modelo usa o representante extremo
# disponivel e preserva o carbono total, recalculando a estequiometria
# a partir das especies efetivamente escritas no DWSIM.
# =========================================================
base_names = [
    "Carbon monoxide",
    "Hydrogen",
    "Water",
    "Carbon dioxide",
]

paraffin_names = {
    1: "Methane",
    2: "Ethane",
    3: "Propane",
    4: "N-butane",
    5: "N-pentane",
    6: "N-hexane",
    7: "N-heptane",
    8: "N-octane",
    9: "N-nonane",
    10: "N-decane",
    11: "N-undecane",
}

olefin_names = {
    2: "Ethylene",
    3: "Propylene",
    4: "1-butene",
    5: "1-pentene",
    6: "1-hexene",
    7: "1-heptene",
    8: "1-octene",
    9: "1-nonene",
    10: "1-decene",
    11: "1-undecene",
}

par_lump_low_name = "N-dodecane"
par_lump_high_name = "N-tetracosane"
par_lump_low_n = 12.0
par_lump_high_n = 24.0

olef_lump_low_name = "1-dodecene"
olef_lump_high_name = "1-eicosene"
olef_lump_low_n = 12.0
olef_lump_high_n = 20.0

required_names = list(base_names)
required_names += [paraffin_names[n] for n in sorted(paraffin_names.keys())]
required_names += [olefin_names[n] for n in sorted(olefin_names.keys())]
required_names += [
    par_lump_low_name,
    par_lump_high_name,
    olef_lump_low_name,
    olef_lump_high_name,
]

# Antes de iniciar os calculos, verifica-se se todos os componentes
# requeridos foram inseridos no flowsheet do DWSIM. A ausencia de
# qualquer componente inviabiliza a escrita correta da corrente de saida.
missing = []
for name in required_names:
    if name not in idx:
        missing.append(name)

if len(missing) > 0:
    raise Exception(
        "Faltam compostos requeridos no flowsheet: "
        + "; ".join(missing)
    )

iCO = idx["Carbon monoxide"]
iH2 = idx["Hydrogen"]
iH2O = idx["Water"]
iCO2 = idx["Carbon dioxide"]

par_idx = {}
for n, name in paraffin_names.items():
    par_idx[n] = idx[name]

olef_idx = {}
for n, name in olefin_names.items():
    olef_idx[n] = idx[name]

iParLow = idx[par_lump_low_name]
iParHigh = idx[par_lump_high_name]
iOlefLow = idx[olef_lump_low_name]
iOlefHigh = idx[olef_lump_high_name]

# =========================================================
# FUNCAO DE LUMPING POR DOIS REPRESENTANTES
# =========================================================
def split_two_lumps(total_mol, total_carbon, n_low, n_high, label):
    if total_mol <= 1.0e-30 or total_carbon <= 1.0e-30:
        return 0.0, 0.0

    n_avg = total_carbon / total_mol
    tol = 1.0e-9

    # Caso normal: o numero medio de carbonos esta entre os dois
    # representantes. Nesta situacao, preservam-se simultaneamente:
    #   1) vazao molar total da familia;
    #   2) carbono total da familia.
    if n_avg >= (n_low - tol) and n_avg <= (n_high + tol):
        high_mol = (total_carbon - n_low * total_mol) / (n_high - n_low)
        low_mol = total_mol - high_mol

    # Caso fora da faixa: em baixa temperatura/alta pressao, a distribuicao
    # pode prever C12+ com numero medio de carbonos maior que o maior composto
    # disponivel no DWSIM. Nesse caso, usa-se o representante extremo e
    # preserva-se o carbono total. A vazao molar do lump passa a ser a vazao
    # equivalente necessaria para representar esse carbono no componente
    # disponivel, mantendo o fechamento elementar C/H/O.
    elif n_avg > (n_high + tol):
        low_mol = 0.0
        high_mol = total_carbon / n_high

    else:
        low_mol = total_carbon / n_low
        high_mol = 0.0

    if abs(low_mol) < 1.0e-18:
        low_mol = 0.0
    if abs(high_mol) < 1.0e-18:
        high_mol = 0.0

    if low_mol < -1.0e-18 or high_mol < -1.0e-18:
        raise Exception(label + ": vazao negativa gerada no lumping.")

    low_mol = max(low_mol, 0.0)
    high_mol = max(high_mol, 0.0)

    return low_mol, high_mol

# =========================================================
# PARAMETROS DO REATOR
# =========================================================
# A variavel independente da integracao e a massa acumulada de catalisador.
# O valor abaixo corresponde a 1 g de catalisador, coerente com a escala
# experimental usada como referencia por Gray et al. (2024). Nas varreduras
# parametricas, a severidade catalitica e avaliada por W/Ftotal e W/FCO,in,
# principalmente pela variacao da vazao de alimentacao no MANAGER.
Wcat_total = 0.001  # kg de catalisador
Nsteps = 1000       # numero de incrementos numericos em W
dW = Wcat_total / float(Nsteps)

# =========================================================
# PARAMETROS CINETICOS - HTFT Fe-K/SiO2 (Gray et al., 2024)
# =========================================================
R_gas = 8.314  # J/(mol.K)

# Parametros de Gray para as taxas de crescimento/terminacao.
# A_k1: sem energia de ativacao reportada.
# A_k5 e A_k6: definidos na base mol/(gcat.h), com pressoes em bar.
# A conversao para mol/(kgcat.s) e aplicada por conv_rate.
A_k1 = 4.5
E_k1 = 0.0

A_k5 = 2.15e7
E_k5 = 95330.0

A_k6 = 76.4
E_k6 = 16360.0

# Parametros de Gray para WGS.
# kWGS usa a energia ajustada por Gray.
A_kWGS = 5.84e2
E_kWGS = 20880.0
Tref_kWGS = 573.0

# Parametro Kv de Lox & Froment, confirmado por Gray como o parametro
# que aparece no denominador da WGS.
# Kv(573 K) = 3.6 bar^(-0.5); Ev = 27.7 kJ/mol.
Tref_Kv = 573.0
Kv_ref = 3.6
E_Kv = 27700.0

# Fatores especificos adotados para ajustar especies leves.
# - CH4: fator associado a formacao adicional de metano.
# - C2H4: fator de correcao para eteno, com a fracao suprimida
#   direcionada para metano pela rota C2H4 + 2 H2 -> 2 CH4.
CH4_CORR = 0.5
C2H4_CORR = 0.5

# Conversao das taxas baseadas em gcat.h para kgcat.s.
# 1 mol/(g.h) = 1000/3600 mol/(kg.s)
conv_rate = 1000.0 / 3600.0

# =========================================================
# LEITURA DA CORRENTE DE ENTRADA
# =========================================================
# O DWSIM fornece a vazao molar total e as fracoes molares globais.
# O vetor Fi armazena as vazoes molares individuais de todos os
# componentes, na mesma ordem em que aparecem no flowsheet.
Ftot = float(feed.GetMolarFlow())

Fi = []
for zi in z:
    Fi.append(Ftot * zi)

# Copia da alimentacao original para diagnostico de balanco elementar.
Fi_in = list(Fi)

# =========================================================
# CONDICOES DE OPERACAO
# =========================================================
# A temperatura e lida em K e a pressao, originalmente em Pa no DWSIM,
# e convertida para bar, unidade usada nas expressoes cineticas.
T_K = float(feed.GetTemperature())
P_bar = float(feed.GetPressure()) / 100000.0

k1 = A_k1 * math.exp(-E_k1 / (R_gas * T_K))
k5 = A_k5 * math.exp(-E_k5 / (R_gas * T_K))
k6 = A_k6 * math.exp(-E_k6 / (R_gas * T_K))
# Constante efetiva da WGS:
# O valor A_kWGS e tratado como referencia a 573 K e corrigido
# pela forma reparametrizada de Arrhenius.
kWGS = A_kWGS * math.exp(
    (-E_kWGS / R_gas) * ((1.0 / T_K) - (1.0 / Tref_kWGS))
)

# Reparametrizacao de Arrhenius para Kv:
# Kv(T) = Kv(Tref) * exp[(-E/R) * (1/T - 1/Tref)]
Kv = Kv_ref * math.exp(
    (-E_Kv / R_gas) * ((1.0 / T_K) - (1.0 / Tref_Kv))
)

Keq_WGS = math.exp(
    5078.0045 / T_K
    - 5.8972089
    + 13.958689e-4 * T_K
    - 27.592844e-8 * T_K * T_K
)

# =========================================================
# INTEGRACAO AO LONGO DO LEITO - EULER EXPLICITO
# =========================================================
# Em cada incremento de massa catalitica, o script recalcula composicao,
# pressoes parciais, parametro de crescimento de cadeia, taxas de formacao
# de hidrocarbonetos e taxa WGS. Em seguida, aplica o limitador
# estequiometrico e atualiza as vazoes molares.
if Ftot > 0.0:
    for step in range(Nsteps):

        Ftot_now = sum(Fi)
        if Ftot_now <= 1.0e-18:
            break

        yCO = Fi[iCO] / Ftot_now
        yH2 = Fi[iH2] / Ftot_now
        yH2O = Fi[iH2O] / Ftot_now
        yCO2 = Fi[iCO2] / Ftot_now

        pCO = max(yCO * P_bar, 1.0e-12)
        pH2 = max(yH2 * P_bar, 1.0e-12)
        pH2O = max(yH2O * P_bar, 1.0e-12)
        pCO2 = max(yCO2 * P_bar, 1.0e-12)

        # -------------------------------------------------
        # CINETICA FTS E DISTRIBUICAO DE PRODUTOS
        # -------------------------------------------------
        # As taxas de parafinas e olefinas sao calculadas a partir da
        # formulacao de Gray et al. (2024), com:
        # 1) termo de parafinas dependente de k5*pH2;
        # 2) termo de olefinas dependente de k6;
        # 3) denominador comum contendo 1/(1-alpha);
        # 4) fator especifico para CH4;
        # 5) fator especifico para C2H4, com conversao da fracao suprimida
        #    para CH4 pela rota C2H4 + 2 H2 -> 2 CH4.
        # -------------------------------------------------
        # alpha representa a probabilidade efetiva de crescimento da
        # cadeia. O intervalo e limitado numericamente para evitar
        # singularidades e distribuicoes degeneradas.
        denom_alpha = (k1 * pCO) + (k5 * pH2) + k6
        alpha = (k1 * pCO) / denom_alpha
        alpha = min(max(alpha, 0.05), 0.95)

        # N_term e D_term concentram a dependencia cinetica usada para
        # distribuir a taxa entre os numeros de carbono.
        N_term = (k1 * pCO) / ((k1 * pCO) + (k5 * pH2))
        D_term = 1.0 + (N_term * (1.0 / (1.0 - alpha)))

        Rpar = {}
        Rolef = {}

        tail_par_mol = 0.0
        tail_par_C = 0.0

        tail_olef_mol = 0.0
        tail_olef_C = 0.0

        rCO_FT = 0.0
        rH2_FT = 0.0
        R_CH4_from_C2 = 0.0
        R_CH4_meth = 0.0

        # Distribuicao completa C1-C100.
        # C1-C11 sao escritos individualmente.
        # C12+ e acumulado para lumping.
        for n in range(1, 101):

            R_p = (
                k5
                * pH2
                * N_term
                * (alpha ** (n - 1))
                / D_term
            ) * conv_rate

            # Formacao de CH4:
            # Mantem-se o metano da distribuicao de parafinas e adiciona-se
            # uma contribuicao especifica associada a metanacao:
            # CO + 3 H2 -> CH4 + H2O.
            if n == 1:
                R_CH4_meth = (
                    CH4_CORR
                    * k5
                    * pH2
                    * N_term
                    / D_term
                ) * conv_rate

            R_o = 0.0

            if n >= 2:
                R_o_ideal = (
                    k6
                    * N_term
                    * (alpha ** (n - 1))
                    / D_term
                ) * conv_rate

                if n == 2:
                    R_o = C2H4_CORR * R_o_ideal
                    r4_C2 = (1.0 - C2H4_CORR) * R_o_ideal

                    # A rota ideal de C2 olefinico consome 2 CO.
                    # A fracao suprimida e convertida em 2 CH4, com consumo
                    # adicional de 2 H2 por mol de C2H4 convertido.
                    rCO_FT += 2.0 * (R_p + R_o_ideal)

                    rH2_FT += (
                        5.0 * R_p
                        + 4.0 * R_o_ideal
                        + 2.0 * r4_C2
                    )

                    R_CH4_from_C2 += 2.0 * r4_C2

                else:
                    R_o = R_o_ideal

                    rCO_FT += n * (R_p + R_o)
                    rH2_FT += (
                        ((2.0 * n + 1.0) * R_p)
                        + ((2.0 * n) * R_o)
                    )

            else:
                # n = 1: parafina C1 + contribuicao especifica de metanacao.
                rCO_FT += (R_p + R_CH4_meth)
                rH2_FT += 3.0 * (R_p + R_CH4_meth)

            if n <= 11:
                Rpar[n] = R_p

                if n >= 2:
                    Rolef[n] = R_o
            else:
                tail_par_mol += R_p
                tail_par_C += n * R_p

                tail_olef_mol += R_o
                tail_olef_C += n * R_o

        # Metano adicional proveniente da conversao da fracao suprimida
        # de eteno e da contribuicao especifica de metanacao.
        Rpar[1] += R_CH4_from_C2 + R_CH4_meth

        # As fracoes C12+ sao escritas em dois representantes por familia
        # quando possivel, preservando simultaneamente mols e carbono.
        Rpar12, Rpar24 = split_two_lumps(
            tail_par_mol,
            tail_par_C,
            par_lump_low_n,
            par_lump_high_n,
            "Lump parafinico C12+"
        )

        Rolef12, Rolef20 = split_two_lumps(
            tail_olef_mol,
            tail_olef_C,
            olef_lump_low_n,
            olef_lump_high_n,
            "Lump olefinico C12+"
        )

        # -------------------------------------------------
        # ESTEQUIOMETRIA REPRESENTADA DOS HIDROCARBONETOS
        # -------------------------------------------------
        # A distribuicao cinetica pode gerar uma cauda C12+ com numero medio
        # de carbonos acima do maior representante disponivel. Depois do
        # relumping, o consumo de CO/H2 precisa ser calculado a partir das
        # especies efetivamente escritas no DWSIM, para fechar C/H/O no stream.
        # Para CcHh: c CO + (c + h/2) H2 -> CcHh + c H2O.
        # -------------------------------------------------
        # As demandas de CO e H2 sao recalculadas apos o lumping.
        # Isso garante que a estequiometria usada no balanco corresponda
        # exatamente as especies que serao escritas na corrente do DWSIM.
        rCO_FT = 0.0
        rH2_FT = 0.0

        for n in range(1, 12):
            fprod = Rpar[n]
            c = float(n)
            h = 2.0 * c + 2.0
            rCO_FT += c * fprod
            rH2_FT += (c + 0.5 * h) * fprod

        for n in range(2, 12):
            fprod = Rolef[n]
            c = float(n)
            h = 2.0 * c
            rCO_FT += c * fprod
            rH2_FT += (c + 0.5 * h) * fprod

        # C12+ parafinico representado por C12 e C24.
        c = par_lump_low_n
        h = 2.0 * c + 2.0
        rCO_FT += c * Rpar12
        rH2_FT += (c + 0.5 * h) * Rpar12

        c = par_lump_high_n
        h = 2.0 * c + 2.0
        rCO_FT += c * Rpar24
        rH2_FT += (c + 0.5 * h) * Rpar24

        # C12+ olefinico representado por C12 e C20.
        c = olef_lump_low_n
        h = 2.0 * c
        rCO_FT += c * Rolef12
        rH2_FT += (c + 0.5 * h) * Rolef12

        c = olef_lump_high_n
        h = 2.0 * c
        rCO_FT += c * Rolef20
        rH2_FT += (c + 0.5 * h) * Rolef20

        # -------------------------------------------------
        # CINETICA WGS - FORMA ADOTADA
        # -------------------------------------------------
        # Forma confirmada por Gray:
        # rwgs = kWGS * [pCO*pH2O - pCO2*sqrt(pH2)/Keq]
        #        / [1 + Kv(T)*pH2O/sqrt(pH2)]^2
        #
        # Pressao em bar. Kv em bar^(-0.5).
        # kWGS tratado como valor efetivo referido a 573 K.
        # Sinal de rwgs:
        #   rwgs > 0: WGS direta, CO + H2O -> CO2 + H2
        #   rwgs < 0: WGS reversa, CO2 + H2 -> CO + H2O
        # -------------------------------------------------
        sqrt_pH2 = math.sqrt(max(pH2, 1.0e-12))

        numerador_wgs = (
            (pCO * pH2O)
            - ((pCO2 * sqrt_pH2) / Keq_WGS)
        )

        denominador_wgs = (
            1.0
            + Kv * (pH2O / sqrt_pH2)
        ) ** 2

        rwgs = (
            kWGS
            * numerador_wgs
            / denominador_wgs
            * conv_rate
        )

        # -------------------------------------------------
        # LIMITADOR ESTEQUIOMETRICO CONJUNTO FTS + WGS
        # -------------------------------------------------
        # O limitador e aplicado ao balanco liquido simultaneo das rotas
        # FTS e WGS. Assim, em cada incremento dW, nenhuma especie pode ser
        # consumida em quantidade maior que a vazao molar disponivel.
        #
        # Balanco liquido de taxas:
        #   CO  : -rCO_FT - rwgs
        #   H2  : -rH2_FT + rwgs
        #   H2O : +rCO_FT - rwgs
        #   CO2 : +rwgs
        #
        # Quando rwgs < 0, ocorre WGS reversa: CO2 e H2 sao consumidos,
        # enquanto CO e H2O sao formados.
        # -------------------------------------------------
        dCO_rate = -rCO_FT - rwgs
        dH2_rate = -rH2_FT + rwgs
        dH2O_rate = rCO_FT - rwgs
        dCO2_rate = rwgs

        scale_all = 1.0

        if dCO_rate < 0.0 and (-dCO_rate * dW) > Fi[iCO]:
            scale_all = min(scale_all, Fi[iCO] / (-dCO_rate * dW))

        if dH2_rate < 0.0 and (-dH2_rate * dW) > Fi[iH2]:
            scale_all = min(scale_all, Fi[iH2] / (-dH2_rate * dW))

        if dH2O_rate < 0.0 and (-dH2O_rate * dW) > Fi[iH2O]:
            scale_all = min(scale_all, Fi[iH2O] / (-dH2O_rate * dW))

        if dCO2_rate < 0.0 and (-dCO2_rate * dW) > Fi[iCO2]:
            scale_all = min(scale_all, Fi[iCO2] / (-dCO2_rate * dW))

        if scale_all < 0.0:
            scale_all = 0.0
        if scale_all > 1.0:
            scale_all = 1.0

        rCO_FT *= scale_all
        rH2_FT *= scale_all
        rwgs *= scale_all

        for n in Rpar:
            Rpar[n] *= scale_all

        for n in Rolef:
            Rolef[n] *= scale_all

        Rpar12 *= scale_all
        Rpar24 *= scale_all
        Rolef12 *= scale_all
        Rolef20 *= scale_all

        # -------------------------------------------------
        # ATUALIZACAO DAS VAZOES MOLARES
        # -------------------------------------------------
        Fi[iCO] += (-rCO_FT - rwgs) * dW
        Fi[iH2] += (-rH2_FT + rwgs) * dW
        Fi[iH2O] += (rCO_FT - rwgs) * dW
        Fi[iCO2] += rwgs * dW

        Fi[iCO] = max(Fi[iCO], 0.0)
        Fi[iH2] = max(Fi[iH2], 0.0)
        Fi[iH2O] = max(Fi[iH2O], 0.0)
        Fi[iCO2] = max(Fi[iCO2], 0.0)

        # C1-C11 parafinas
        for n in range(1, 12):
            Fi[par_idx[n]] += Rpar[n] * dW
            Fi[par_idx[n]] = max(Fi[par_idx[n]], 0.0)

        # C2-C11 olefinas
        for n in range(2, 12):
            Fi[olef_idx[n]] += Rolef[n] * dW
            Fi[olef_idx[n]] = max(Fi[olef_idx[n]], 0.0)

        # C12+ parafinas em dois lumps
        Fi[iParLow] += Rpar12 * dW
        Fi[iParHigh] += Rpar24 * dW

        # C12+ olefinas em dois lumps
        Fi[iOlefLow] += Rolef12 * dW
        Fi[iOlefHigh] += Rolef20 * dW

        Fi[iParLow] = max(Fi[iParLow], 0.0)
        Fi[iParHigh] = max(Fi[iParHigh], 0.0)
        Fi[iOlefLow] = max(Fi[iOlefLow], 0.0)
        Fi[iOlefHigh] = max(Fi[iOlefHigh], 0.0)

# =========================================================
# ESCRITA DA CORRENTE DE SAIDA
# =========================================================
# A corrente de saida herda as propriedades gerais da alimentacao e recebe
# as vazoes molares calculadas pelo CUSTOM. Em seguida, o DWSIM recalcula
# as propriedades termodinamicas da corrente resultante.
prod.Clear()
prod.Assign(feed)

Ftot_out = sum(Fi)

if Ftot_out <= 0.0:
    try:
        prod.Calculate()
    except:
        pass
else:
    prod.SetMolarFlow(Ftot_out)

    for i in range(len(Fi)):
        prod.SetOverallCompoundMolarFlow(i, Fi[i])

    try:
        prod.Calculate()
    except:
        pass


# =========================================================
# DIAGNOSTICO - BALANCOS ELEMENTARES C/H/O
# =========================================================
# Este bloco NAO altera a corrente de saida. Ele apenas registra, em CSV,
# os balancos elementares calculados diretamente a partir dos fluxos molares
# do CUSTOM antes e depois da integracao.
#
# O diagnostico permite verificar fechamento de carbono, hidrogenio e
# oxigenio para cada caso executado no DWSIM.
#
# Local padrao de saida no Windows:
#   Desktop\06_DIAGNOSTICO_CUSTOM_Fe_v3.8_CHO.csv
#
# Se a area de trabalho nao estiver acessivel, tenta salvar na pasta corrente.
# =========================================================

def _atoms_component(name):
    # Retorna (C,H,O) por mol de componente.
    if name == "Carbon monoxide":
        return (1.0, 0.0, 1.0)
    if name == "Hydrogen":
        return (0.0, 2.0, 0.0)
    if name == "Water":
        return (0.0, 2.0, 1.0)
    if name == "Carbon dioxide":
        return (1.0, 0.0, 2.0)

    # Parafinas explicitas CnH(2n+2)
    for _n, _name in paraffin_names.items():
        if name == _name:
            return (float(_n), 2.0 * float(_n) + 2.0, 0.0)

    # Olefinas explicitas CnH(2n)
    for _n, _name in olefin_names.items():
        if name == _name:
            return (float(_n), 2.0 * float(_n), 0.0)

    # Lumps C12+
    if name == par_lump_low_name:
        return (par_lump_low_n, 2.0 * par_lump_low_n + 2.0, 0.0)
    if name == par_lump_high_name:
        return (par_lump_high_n, 2.0 * par_lump_high_n + 2.0, 0.0)
    if name == olef_lump_low_name:
        return (olef_lump_low_n, 2.0 * olef_lump_low_n, 0.0)
    if name == olef_lump_high_name:
        return (olef_lump_high_n, 2.0 * olef_lump_high_n, 0.0)

    # Componentes nao participantes, como Sulfolane, recebem zero no diagnostico.
    return (0.0, 0.0, 0.0)


def _element_totals(Fvec):
    C = 0.0
    H = 0.0
    O = 0.0
    for _i in range(len(Fvec)):
        _c, _h, _o = _atoms_component(component_names[_i])
        C += _c * Fvec[_i]
        H += _h * Fvec[_i]
        O += _o * Fvec[_i]
    return C, H, O


def _sum_paraffins(Fvec, n1, n2):
    s = 0.0
    for _n in range(n1, n2 + 1):
        if _n in par_idx:
            s += Fvec[par_idx[_n]]
    return s


def _sum_olefins(Fvec, n1, n2):
    s = 0.0
    for _n in range(n1, n2 + 1):
        if _n in olef_idx:
            s += Fvec[olef_idx[_n]]
    return s


def _fmt(x):
    try:
        return "%.15E" % float(x)
    except:
        return str(x)


def _rel_percent(err, base):
    if abs(base) <= 1.0e-30:
        return 0.0
    return 100.0 * err / base

try:
    C_in, H_in, O_in = _element_totals(Fi_in)
    C_out, H_out, O_out = _element_totals(Fi)

    errC = C_out - C_in
    errH = H_out - H_in
    errO = O_out - O_in

    FCO_in = Fi_in[iCO]
    FH2_in = Fi_in[iH2]
    FCO_out = Fi[iCO]
    FH2_out = Fi[iH2]
    FH2O_out = Fi[iH2O]
    FCO2_out = Fi[iCO2]

    H2CO_in = 0.0
    if FCO_in > 1.0e-30:
        H2CO_in = FH2_in / FCO_in

    H2CO_out = 0.0
    if FCO_out > 1.0e-30:
        H2CO_out = FH2_out / FCO_out

    XCO_diag = 0.0
    if FCO_in > 1.0e-30:
        XCO_diag = (FCO_in - FCO_out) / FCO_in

    XH2_diag = 0.0
    if FH2_in > 1.0e-30:
        XH2_diag = (FH2_in - FH2_out) / FH2_in

    F_CH4 = Fi[par_idx[1]]
    F_C2_C4 = (
        _sum_paraffins(Fi, 2, 4)
        + _sum_olefins(Fi, 2, 4)
    )
    F_C5_C11 = (
        _sum_paraffins(Fi, 5, 11)
        + _sum_olefins(Fi, 5, 11)
    )
    F_C12plus = (
        Fi[iParLow]
        + Fi[iParHigh]
        + Fi[iOlefLow]
        + Fi[iOlefHigh]
    )

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        desktop = os.getcwd()

    diag_path = os.path.join(desktop, "06_DIAGNOSTICO_CUSTOM_Fe_v3.8_CHO.csv")
    write_header = not os.path.exists(diag_path)

    header = [
        "timestamp",
        "T_K",
        "T_C",
        "P_bar",
        "Wcat_kg",
        "Ftot_in_mol_s",
        "FCO_in_mol_s",
        "FH2_in_mol_s",
        "H2CO_in",
        "Ftot_out_mol_s",
        "FCO_out_mol_s",
        "FH2_out_mol_s",
        "FH2O_out_mol_s",
        "FCO2_out_mol_s",
        "H2CO_out",
        "XCO_diag",
        "XH2_diag",
        "F_CH4_mol_s",
        "F_C2_C4_mol_s",
        "F_C5_C11_mol_s",
        "F_C12plus_mol_s",
        "C_in",
        "C_out",
        "C_err_out_minus_in",
        "C_err_percent",
        "H_in",
        "H_out",
        "H_err_out_minus_in",
        "H_err_percent",
        "O_in",
        "O_out",
        "O_err_out_minus_in",
        "O_err_percent",
    ]

    row = [
        System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
        _fmt(T_K),
        _fmt(T_K - 273.15),
        _fmt(P_bar),
        _fmt(Wcat_total),
        _fmt(sum(Fi_in)),
        _fmt(FCO_in),
        _fmt(FH2_in),
        _fmt(H2CO_in),
        _fmt(sum(Fi)),
        _fmt(FCO_out),
        _fmt(FH2_out),
        _fmt(FH2O_out),
        _fmt(FCO2_out),
        _fmt(H2CO_out),
        _fmt(XCO_diag),
        _fmt(XH2_diag),
        _fmt(F_CH4),
        _fmt(F_C2_C4),
        _fmt(F_C5_C11),
        _fmt(F_C12plus),
        _fmt(C_in),
        _fmt(C_out),
        _fmt(errC),
        _fmt(_rel_percent(errC, C_in)),
        _fmt(H_in),
        _fmt(H_out),
        _fmt(errH),
        _fmt(_rel_percent(errH, H_in)),
        _fmt(O_in),
        _fmt(O_out),
        _fmt(errO),
        _fmt(_rel_percent(errO, O_in)),
    ]

    fdiag = open(diag_path, "a")
    if write_header:
        fdiag.write(";".join(header) + "\n")
    fdiag.write(";".join(row) + "\n")
    fdiag.close()

except Exception as diag_ex:
    # Nao interromper a simulacao por falha do diagnostico.
    try:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.isdir(desktop):
            desktop = os.getcwd()
        err_path = os.path.join(desktop, "06_DIAGNOSTICO_CUSTOM_Fe_v3.8_ERRO.txt")
        ferr = open(err_path, "a")
        ferr.write(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + str(diag_ex) + "\n")
        ferr.close()
    except:
        pass

