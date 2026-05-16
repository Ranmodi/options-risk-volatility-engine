# -*- coding: utf-8 -*-

# PUBLIC-SAFE PORTFOLIO VERSION
# Real credentials, client data, private endpoints and internal paths were removed or replaced by placeholders.
"""
Fixing de operações estruturadas

- Lê "Structured_Products.xlsx" (aba 'Structured Products')
- Lê "Advisor Base.xlsx"
- Encontra operações com Data de Fixing = hoje
- Calcula resultado por estrutura (Stock, Fence KI, Smart Hedge, Aceleradora KO, TWIP, Smart Up)
- Envia 1 e-mail por assessor com 1 "cartão azul" por operação (layout igual ao exemplo)
"""

# =========================
# BOOTSTRAP OneDrive/Excel
# - força diretório de trabalho na pasta do script
# - cria pasta de logs e grava traceback em caso de erro
# =========================
from pathlib import Path as _Path
import os as _os
import sys as _sys
import traceback as _traceback
from datetime import datetime as _dt_datetime

def _get_base_dir() -> _Path:
    if getattr(_sys, "frozen", False):
        return _Path(_sys.executable).resolve().parent
    return _Path(__file__).resolve().parent

BASE_DIR = _get_base_dir()
try:
    _os.chdir(BASE_DIR)
except Exception:
    pass

def _infer_root_dir(base_dir: _Path) -> _Path:
    name = (base_dir.name or "").strip().lower()
    if name in {"scripts py", "scripts_py", "scripts python", "scripts_python", "scripts"}:
        return base_dir.parent
    return base_dir

ROOT_DIR = _infer_root_dir(BASE_DIR)

# Pastas padrão (relativas ao ROOT_DIR):
#   ROOT_DIR\Scripts py        -> onde ficam os .py
#   ROOT_DIR\logs              -> logs centralizados
#   ROOT_DIR\Planilhas base    -> onde ficam as planilhas base baixadas
PLANILHAS_DIR = ROOT_DIR / "Planilhas base"

LOG_DIR = ROOT_DIR / "logs"
if not LOG_DIR.exists():
    # fallback para manter compatibilidade com versões antigas
    LOG_DIR = BASE_DIR / "_logs"
try:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
except Exception:
    pass
def _log_unhandled_exception(prefix: str = "run_err") -> None:
    try:
        ts = _dt_datetime.now().strftime("%Y%m%d_%H%M%S")
        p_latest = LOG_DIR / f"{prefix}_LATEST.log"
        p_ts = LOG_DIR / f"{prefix}_{ts}.log"
        tb = _traceback.format_exc()
        p_latest.write_text(tb, encoding="utf-8", errors="ignore")
        p_ts.write_text(tb, encoding="utf-8", errors="ignore")
    except Exception:
        # não deixa o logging derrubar a execução
        pass


import os
import sys
import math
from datetime import datetime, date

import numpy as np
import pandas as pd

try:
    import win32com.client as win32
except ImportError:
    win32 = None

def pause():
    """
    Pausa segura para modo console.
    No executável gerado com --noconsole, não faz nada.
    """
    try:
        if sys.stdin and hasattr(sys.stdin, "isatty") and sys.stdin.isatty():
            input("Tecle ENTER para sair.")
    except Exception:
        # Em modo GUI / sem stdin, simplesmente ignora
        pass

# ---------------------- Utilidades gerais ---------------------- #

def get_base_dir() -> str:
    """Retorna a pasta base onde ficam os arquivos (Planilhas base, se existir)."""
    try:
        # Preferência: pasta padrão das automações (Planilhas base), relativa ao script
        if "PLANILHAS_DIR" in globals() and PLANILHAS_DIR.exists():
            return str(PLANILHAS_DIR)
    except Exception:
        pass

    # Fallback: mesma pasta do script / executável
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def normalize_name(s: str) -> str:
    """Normaliza nomes de colunas para facilitar matching (sem acento, espaços, hífens etc.)."""
    if s is None:
        return ""
    s = str(s).lower()
    replacements = {
        "á": "a", "à": "a", "â": "a", "ã": "a",
        "é": "e", "ê": "e",
        "í": "i",
        "ó": "o", "ô": "o", "õ": "o",
        "ú": "u",
        "ç": "c"
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    for ch in [" ", "-", "_", "."]:
        s = s.replace(ch, "")
    return s


def find_col(df: pd.DataFrame, must_contain):
    """
    Procura uma coluna cujo nome (normalizado) contenha TODOS os termos em must_contain.
    Ex.: "E-mail Assessor" -> "emailassessor" casa com ["email", "assessor"].
    """
    if isinstance(must_contain, (list, tuple)):
        must = [normalize_name(m) for m in must_contain]
    else:
        must = [normalize_name(must_contain)]
    for col in df.columns:
        name_norm = normalize_name(col)
        if all(m in name_norm for m in must):
            return col
    return None


def normalize_conta_series(s: pd.Series) -> pd.Series:
    """
    Normaliza 'Conta':
    - mantém apenas dígitos
    - remove zeros à esquerda
    Ex.: '02787004' -> '2787004'
    """
    s = s.astype(str)
    digits = s.str.replace(r"\D+", "", regex=True)

    def norm_one(x: str) -> str:
        x = x.strip()
        if x == "":
            return ""
        try:
            return str(int(x))  # int() tira zeros à esquerda
        except Exception:
            return x

    return digits.apply(norm_one)


def parse_pct(s):
    """Converte '108,73%' -> 1.0873 ou '8,73' -> 0.0873, etc."""
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return None
    if isinstance(s, (int, float, np.floating)):
        val = float(s)
        if val > 1.5:
            return val / 100.0
        return val
    txt = str(s).strip()
    if not txt:
        return None
    txt = txt.replace('%', '').replace(' ', '')
    txt = txt.replace('.', '').replace(',', '.')
    try:
        val = float(txt)
    except ValueError:
        return None
    if val > 1.5:
        val = val / 100.0
    return val


def parse_date(s):
    """Converte textos diversos para datetime."""
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return None
    if isinstance(s, (datetime, date)):
        return s
    txt = str(s).strip()
    if not txt:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(txt, fmt)
        except ValueError:
            continue
    return None


def fmt_date(d):
    """Formata datetime/date em dd/mm/aaaa."""
    if d is None or (isinstance(d, float) and math.isnan(d)):
        return ""
    if isinstance(d, datetime):
        return d.strftime("%d/%m/%Y")
    if isinstance(d, date):
        return d.strftime("%d/%m/%Y")
    dt = parse_date(d)
    if dt is None:
        return ""
    return dt.strftime("%d/%m/%Y")


def get_group_common_value(g: pd.DataFrame, col: str):
    """Retorna o primeiro valor não nulo de uma coluna dentro do grupo."""
    series = g[col]
    for v in series:
        if not (pd.isna(v) or v == ""):
            return v
    return None


def get_leg(g: pd.DataFrame, tipo=None, direcao=None):
    """Retorna um leg específico por Tipo Operação / Direção."""
    q = g
    if tipo is not None:
        q = q[q['Tipo Operação'] == tipo]
    if direcao is not None:
        q = q[q['Direção'] == direcao]
    if q.empty:
        return None
    return q.iloc[0]


def fmt_money(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return f"R$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def fmt_pct(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return f"{v*100:,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')


def find_signature_image(base_dir: str) -> str | None:
    """Procura arquivo de assinatura (assinatura*.png/jpg/jpeg/gif) na pasta."""
    for f in os.listdir(base_dir):
        lower = f.lower()
        if "assinatura" in lower and lower.endswith((".png", ".jpg", ".jpeg", ".gif")):
            return os.path.join(base_dir, f)
    return None


# ---------------------- Cálculo por operação ---------------------- #


import pandas as pd

def normalize_prod_name(nome: str | None) -> str:
    if nome is None:
        return ""
    return str(nome).strip().lower()


def build_legs_from_group(g: pd.DataFrame):
    """
    Constrói uma lista de pernas (legs) genéricas a partir do grupo.
    Cada leg carrega tipo, direção, strike em % e R$, KI e KO.
    """
    legs = []
    for _, row in g.iterrows():
        strike_px_raw = row.get("Strike (R$)")
        if strike_px_raw is None or strike_px_raw == "" or pd.isna(strike_px_raw):
            strike_px = None
        else:
            strike_px = float(strike_px_raw)

        leg = {
            "tipo": row.get("Tipo Operação"),
            "direcao": row.get("Direção"),
            "strike_pct": parse_pct(row.get("Strike")),
            "strike_px": strike_px,
            "knock_in_pct": parse_pct(row.get("KnockIn")),
            "knock_out_pct": parse_pct(row.get("KnockOut")),
        }
        legs.append(leg)
    return legs


def infer_strategy_from_legs(legs):
    """
    Tenta descobrir o tipo da estrutura olhando só para as pernas.
    Retorna uma 'chave' interna, ex.: 'stock_or_coupon', 'fence_ki', etc.
    """
    calls_buy = [l for l in legs if str(l.get("tipo")).upper() == "CALL" and str(l.get("direcao")).upper() == "BUY"]
    calls_sell = [l for l in legs if str(l.get("tipo")).upper() == "CALL" and str(l.get("direcao")).upper() == "SELL"]
    puts_buy = [l for l in legs if str(l.get("tipo")).upper() == "PUT" and str(l.get("direcao")).upper() == "BUY"]
    puts_sell = [l for l in legs if str(l.get("tipo")).upper() == "PUT" and str(l.get("direcao")).upper() == "SELL"]

    def approx(v, target, tol=1e-4):
        return v is not None and abs(v - target) <= tol

    # 1) Compra simples de PUT (seguro de queda)
    if len(puts_buy) >= 1 and not calls_buy and not calls_sell and not puts_sell:
        if all(pb.get("strike_pct") is not None for pb in puts_buy):
            return "compra_put"

    # 2) Stock or Coupon: CALL SELL + PUT BUY no mesmo strike, sem outras pernas
    if len(calls_sell) == 1 and len(puts_buy) == 1 and not calls_buy and not puts_sell:
        cs = calls_sell[0]
        pb = puts_buy[0]
        if cs.get("strike_pct") is not None and pb.get("strike_pct") is not None:
            if approx(cs["strike_pct"], pb["strike_pct"]):
                return "stock_or_coupon"

    # 3) Fence KI com Ativo:
    # CALL SELL (strike >1, KI>1) + PUT SELL (strike<1) + PUT BUY (~100%)
    if calls_sell and puts_sell and puts_buy:
        has_put_buy_100 = any(approx(pb.get("strike_pct"), 1.0) for pb in puts_buy)
        has_put_sell_below = any((ps.get("strike_pct") is not None and ps["strike_pct"] < 1.0) for ps in puts_sell)
        has_call_sell_above = any((cs.get("strike_pct") is not None and cs["strike_pct"] > 1.0) for cs in calls_sell)
        has_call_ki = any((cs.get("knock_in_pct") is not None and cs["knock_in_pct"] > 1.0) for cs in calls_sell)
        if has_put_buy_100 and has_put_sell_below and has_call_sell_above and has_call_ki:
            return "fence_ki"

    # 4) Twin Win Protected - TWIP:
    # CALL SELL ~100% com KI>1 + PUT BUY ~100% com KO<1
    if calls_sell and puts_buy and not puts_sell:
        has_call_100_ki = any(
            approx(cs.get("strike_pct"), 1.0)
            and cs.get("knock_in_pct") is not None
            and cs["knock_in_pct"] > 1.0
            for cs in calls_sell
        )
        has_put_100_ko_below = any(
            approx(pb.get("strike_pct"), 1.0)
            and pb.get("knock_out_pct") is not None
            and pb["knock_out_pct"] < 1.0
            for pb in puts_buy
        )
        if has_call_100_ki and has_put_100_ko_below:
            return "twip"

    # 5) Aceleradora KO com Ativo:
    # CALL BUY ~100% com KO>1 + CALL SELL >1, sem PUT
    if calls_buy and calls_sell and not puts_buy and not puts_sell:
        has_call_buy_100_ko = any(
            approx(cb.get("strike_pct"), 1.0)
            and cb.get("knock_out_pct") is not None
            and cb["knock_out_pct"] > 1.0
            for cb in calls_buy
        )
        has_call_sell_above = any(
            cs.get("strike_pct") is not None and cs["strike_pct"] > 1.0 for cs in calls_sell
        )
        if has_call_buy_100_ko and has_call_sell_above:
            return "aceleradora_ko"

    # 6) Smart Hedge (simplificado):
    # CALL SELL >1 com KI>1 + PUT BUY ~100%
    if calls_sell and puts_buy and not puts_sell:
        has_call_sell_above_ki = any(
            cs.get("strike_pct") is not None
            and cs["strike_pct"] > 1.0
            and cs.get("knock_in_pct") is not None
            and cs["knock_in_pct"] > 1.0
            for cs in calls_sell
        )
        has_put_100 = any(approx(pb.get("strike_pct"), 1.0) for pb in puts_buy)
        if has_call_sell_above_ki and has_put_100:
            return "smart_hedge"

    # 7) Compra de PUT Spread: PUT BUY + PUT SELL, sem CALL
    if puts_buy and puts_sell and not calls_buy and not calls_sell:
        if all(l.get("strike_pct") is not None for l in puts_buy + puts_sell):
            return "compra_put_spread"

    # 8) Compra de CALL Spread: CALL BUY + CALL SELL, sem KI/KO relevantes e sem PUT
    if calls_buy and calls_sell and not puts_buy and not puts_sell:
        has_barrier = any(
            (l.get("knock_in_pct") not in (None, 0)) or (l.get("knock_out_pct") not in (None, 0))
            for l in calls_buy + calls_sell
        )
        if not has_barrier:
            return "compra_call_spread"

    # 9) Financiamento com Ativo: apenas CALL SELL, sem KI/KO e sem PUT
    if calls_sell and not calls_buy and not puts_buy and not puts_sell:
        has_barrier = any(
            (l.get("knock_in_pct") not in (None, 0)) or (l.get("knock_out_pct") not in (None, 0))
            for l in calls_sell
        )
        if not has_barrier:
            return "financiamento_ativo"

    # Se nada bater, devolve None e cai no fallback
    return None



def get_strike_pct_from_row(row):
    if row is None:
        return None
    return parse_pct(row.get("Strike"))


def get_strike_px_from_row(row):
    if row is None:
        return None
    v = row.get("Strike (R$)")
    if v is None or v == "" or pd.isna(v):
        return None
    return float(v)


# ---------------------- Modelos de estratégia ---------------------- #

def strategy_stock_or_coupon(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                             call_sell, put_sell, put_buy, call_buy):
    model_name = "stock_or_coupon"
    barreira_flag = data_knock_out is not None
    result_bruto_pct = None

    if data_knock_out is None:
        row = call_sell if call_sell is not None else put_buy
        strike_pct = get_strike_pct_from_row(row)
        if strike_pct is not None:
            result_bruto_pct = strike_pct - 1.0
    else:
        if variation is not None:
            result_bruto_pct = variation

    return model_name, barreira_flag, result_bruto_pct


def strategy_fence_ki(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                      call_sell, put_sell, put_buy, call_buy):
    model_name = "fence_ki"
    barreira_flag = data_knock_in is not None
    result_bruto_pct = None

    put_sell_strike_px = get_strike_px_from_row(put_sell)
    call_sell_strike_pct = get_strike_pct_from_row(call_sell)

    if data_knock_in is not None:
        if call_sell_strike_pct is not None and variation is not None and preco_merc is not None:
            call_strike_px = get_strike_px_from_row(call_sell)
            if call_strike_px is not None and preco_merc > call_strike_px:
                result_bruto_pct = call_sell_strike_pct - 1.0
            else:
                result_bruto_pct = variation
    else:
        if variation is not None and preco_ref is not None and preco_merc is not None:
            if preco_merc >= preco_ref:
                result_bruto_pct = variation
            else:
                if put_sell_strike_px is None:
                    result_bruto_pct = variation
                else:
                    if preco_merc >= put_sell_strike_px:
                        result_bruto_pct = 0.0
                    else:
                        result_bruto_pct = preco_merc / put_sell_strike_px - 1.0

    return model_name, barreira_flag, result_bruto_pct


def strategy_smart_hedge(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                         call_sell, put_sell, put_buy, call_buy):
    model_name = "smart_hedge"
    barreira_flag = data_knock_in is not None
    result_bruto_pct = None

    put_buy_strike_pct = get_strike_pct_from_row(put_buy)
    protective_floor = (
        preco_ref * put_buy_strike_pct
        if (put_buy_strike_pct is not None and preco_ref is not None)
        else None
    )

    if data_knock_in is not None:
        call_sell_strike_pct = get_strike_pct_from_row(call_sell)
        if call_sell_strike_pct is not None and variation is not None and preco_merc is not None:
            call_strike_px = get_strike_px_from_row(call_sell)
            if call_strike_px is not None and preco_merc > call_strike_px:
                result_bruto_pct = call_sell_strike_pct - 1.0
            else:
                result_bruto_pct = variation
    else:
        if variation is not None and preco_ref is not None and preco_merc is not None:
            if preco_merc > preco_ref:
                result_bruto_pct = variation
            else:
                if protective_floor is None:
                    result_bruto_pct = variation
                else:
                    if preco_merc >= protective_floor:
                        result_bruto_pct = 0.0
                    else:
                        result_bruto_pct = protective_floor / preco_ref - 1.0

    return model_name, barreira_flag, result_bruto_pct


def strategy_aceleradora_ko(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                            call_sell, put_sell, put_buy, call_buy):
    model_name = "aceleradora_ko"
    # barreira = KO
    barreira_flag = data_knock_out is not None
    result_bruto_pct = None

    # KO%: maior KnockOut > 0% do grupo
    ko_pcts = []
    if "KnockOut" in g.columns:
        for v in g["KnockOut"]:
            p = parse_pct(v)
            if p is not None and p > 0:
                ko_pcts.append(p)
    ko_pct = max(ko_pcts) if ko_pcts else None
    call_sell_strike_px = get_strike_px_from_row(call_sell)

    if variation is not None and preco_ref is not None and preco_merc is not None:
        if data_knock_out is not None:
            # KO rompido: se ativo acima do strike da CALL SELL, ganho = KO - 1
            if call_sell_strike_px is not None and preco_merc > call_sell_strike_px and ko_pct is not None:
                result_bruto_pct = ko_pct - 1.0
            else:
                # se, no vencimento, estiver abaixo do strike da CALL SELL, volta a ser variação simples
                result_bruto_pct = variation
        else:
            # KO não rompido: alta em dobro, queda 1x
            if variation >= 0:
                result_bruto_pct = 2.0 * variation
            else:
                result_bruto_pct = variation

    return model_name, barreira_flag, result_bruto_pct


def strategy_twip(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                  call_sell, put_sell, put_buy, call_buy):
    model_name = "twip"
    barreira_flag = (data_knock_in is not None) or (data_knock_out is not None)
    result_bruto_pct = None

    if data_knock_in is not None or data_knock_out is not None:
        result_bruto_pct = 0.0
    else:
        if variation is not None:
            result_bruto_pct = abs(variation)

    return model_name, barreira_flag, result_bruto_pct


def strategy_smart_up(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                      call_sell, put_sell, put_buy, call_buy):
    model_name = "smart_up"
    barreira_flag = data_knock_in is not None
    result_bruto_pct = None

    put_buy_strike_pct = get_strike_pct_from_row(put_buy)
    protective_floor = (
        preco_ref * put_buy_strike_pct
        if (put_buy_strike_pct is not None and preco_ref is not None)
        else None
    )

    if data_knock_in is not None:
        call_sell_strike_pct = get_strike_pct_from_row(call_sell)
        if call_sell_strike_pct is not None and variation is not None and preco_merc is not None:
            call_strike_px = get_strike_px_from_row(call_sell)
            if call_strike_px is not None and preco_merc > call_strike_px:
                result_bruto_pct = call_sell_strike_pct - 1.0
            else:
                result_bruto_pct = variation
    else:
        if variation is not None and preco_ref is not None and preco_merc is not None:
            if preco_merc >= preco_ref:
                # Smart Up: upside em dobro
                result_bruto_pct = 2.0 * variation
            else:
                if protective_floor is None:
                    result_bruto_pct = variation
                else:
                    if preco_merc >= protective_floor:
                        result_bruto_pct = 0.0
                    else:
                        result_bruto_pct = protective_floor / preco_ref - 1.0

    return model_name, barreira_flag, result_bruto_pct

def strategy_compra_put(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                        call_sell, put_sell, put_buy, call_buy):
    """
    Compra de PUT "seca": payoff = max(K/S0 - S_T/S0, 0) - prêmio (% do nocional).
    """
    model_name = "compra_put"
    barreira_flag = False
    result_bruto_pct = None

    if preco_ref is None or preco_merc is None:
        return model_name, barreira_flag, result_bruto_pct

    # Strike relativo (K/S0)
    strike_rel = get_strike_pct_from_row(put_buy)
    if strike_rel is None:
        strike_rel = 1.0  # default 100%

    # Tentar achar prêmio (% do nocional) a partir de Perda Máxima ou Prêmio / Nocional
    nocional = get_group_common_value(g, "Nocional")
    perda_max = get_group_common_value(g, "Perda Máxima")
    premio_total = get_group_common_value(g, "Premio (R$)")

    premio_pct = 0.0
    try:
        if nocional and perda_max and float(nocional) != 0:
            premio_pct = float(perda_max) / float(nocional)
        elif nocional and premio_total and float(nocional) != 0:
            premio_pct = float(premio_total) / float(nocional)
    except Exception:
        premio_pct = 0.0

    price_ratio = preco_merc / preco_ref  # S_T / S_0
    payoff_pct = max(strike_rel - price_ratio, 0.0)
    result_bruto_pct = payoff_pct - premio_pct

    return model_name, barreira_flag, result_bruto_pct


def strategy_compra_put_spread(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                               call_sell, put_sell, put_buy, call_buy):
    """
    Long PUT Spread: compra PUT strike alto, venda PUT strike baixo.
    Payoff em % do nocional:
      - S_T >= K_high: 0
      - K_low < S_T < K_high: K_high/S0 - S_T/S0
      - S_T <= K_low: K_high/S0 - K_low/S0 (constante)
    Menos o prêmio líquido pago.
    """
    model_name = "compra_put_spread"
    barreira_flag = False
    result_bruto_pct = None

    if preco_ref is None or preco_merc is None or put_buy is None or put_sell is None:
        return model_name, barreira_flag, result_bruto_pct

    k_buy = get_strike_pct_from_row(put_buy)
    k_sell = get_strike_pct_from_row(put_sell)
    if k_buy is None or k_sell is None:
        return model_name, barreira_flag, result_bruto_pct

    k_high = max(k_buy, k_sell)
    k_low = min(k_buy, k_sell)

    # prêmio líquido (% do nocional) ~ Perda Máxima / Nocional
    nocional = get_group_common_value(g, "Nocional")
    perda_max = get_group_common_value(g, "Perda Máxima")
    premio_total = get_group_common_value(g, "Premio (R$)")

    premio_pct = 0.0
    try:
        if nocional and perda_max and float(nocional) != 0:
            premio_pct = float(perda_max) / float(nocional)
        elif nocional and premio_total and float(nocional) != 0:
            premio_pct = float(premio_total) / float(nocional)
    except Exception:
        premio_pct = 0.0

    price_ratio = preco_merc / preco_ref  # S_T / S_0

    if price_ratio >= k_high:
        payoff_pct = 0.0
    elif price_ratio <= k_low:
        payoff_pct = k_high - k_low
    else:
        payoff_pct = k_high - price_ratio

    result_bruto_pct = payoff_pct - premio_pct
    return model_name, barreira_flag, result_bruto_pct


def strategy_compra_call_spread(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                                call_sell, put_sell, put_buy, call_buy):
    """
    Long CALL Spread: compra CALL strike baixo, vende CALL strike alto.
    Payoff em % do nocional:
      - S_T <= K_low: 0
      - K_low < S_T < K_high: S_T/S0 - K_low/S0
      - S_T >= K_high: K_high/S0 - K_low/S0 (constante)
    Menos o prêmio líquido pago.
    """
    model_name = "compra_call_spread"
    barreira_flag = False
    result_bruto_pct = None

    if preco_ref is None or preco_merc is None or call_buy is None or call_sell is None:
        return model_name, barreira_flag, result_bruto_pct

    k_buy = get_strike_pct_from_row(call_buy)
    k_sell = get_strike_pct_from_row(call_sell)
    if k_buy is None or k_sell is None:
        return model_name, barreira_flag, result_bruto_pct

    k_low = min(k_buy, k_sell)
    k_high = max(k_buy, k_sell)

    nocional = get_group_common_value(g, "Nocional")
    perda_max = get_group_common_value(g, "Perda Máxima")
    premio_total = get_group_common_value(g, "Premio (R$)")

    premio_pct = 0.0
    try:
        if nocional and perda_max and float(nocional) != 0:
            premio_pct = float(perda_max) / float(nocional)
        elif nocional and premio_total and float(nocional) != 0:
            premio_pct = float(premio_total) / float(nocional)
    except Exception:
        premio_pct = 0.0

    price_ratio = preco_merc / preco_ref  # S_T / S_0

    if price_ratio <= k_low:
        payoff_pct = 0.0
    elif price_ratio >= k_high:
        payoff_pct = k_high - k_low
    else:
        payoff_pct = price_ratio - k_low

    result_bruto_pct = payoff_pct - premio_pct
    return model_name, barreira_flag, result_bruto_pct


def strategy_financiamento_ativo(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                                 call_sell, put_sell, put_buy, call_buy):
    """
    Financiamento com Ativo (covered call):
      - Compra ação (S0) + venda CALL (strike K).
      - Retorno em % do nocional ≈ retorno da ação + prêmio (%),
        limitado pelo Ganho Máximo.
    """
    model_name = "financiamento_ativo"
    barreira_flag = False
    result_bruto_pct = None

    if preco_ref is None or preco_merc is None or call_sell is None:
        return model_name, barreira_flag, result_bruto_pct

    strike_rel = get_strike_pct_from_row(call_sell)  # K/S0
    nocional = get_group_common_value(g, "Nocional")
    ganho_max = get_group_common_value(g, "Ganho Máximo")
    premio_total = get_group_common_value(g, "Premio (R$)")

    ganho_max_pct = None
    if nocional and ganho_max and float(nocional) != 0:
        try:
            ganho_max_pct = float(ganho_max) / float(nocional)
        except Exception:
            ganho_max_pct = None

    premio_pct = 0.0
    try:
        if nocional and premio_total and float(nocional) != 0:
            premio_pct = float(premio_total) / float(nocional)
        elif ganho_max_pct is not None and strike_rel is not None:
            # ganho_max = (K/S0 - 1) + prêmio
            premio_pct = ganho_max_pct - max(strike_rel - 1.0, 0.0)
    except Exception:
        premio_pct = 0.0

    price_ratio = preco_merc / preco_ref  # S_T / S_0
    ret_ativo_pct = price_ratio - 1.0

    if strike_rel is not None and price_ratio >= strike_rel and ganho_max_pct is not None:
        # Acima do strike, resultado trava no ganho máximo
        result_bruto_pct = ganho_max_pct
    else:
        # Abaixo do strike, retorno = retorno da ação + prêmio
        result = ret_ativo_pct + premio_pct
        if ganho_max_pct is not None:
            result = min(result, ganho_max_pct)
        result_bruto_pct = result

    return model_name, barreira_flag, result_bruto_pct


def strategy_cupom_alto_retorno(g, preco_ref, preco_merc, variation, data_knock_in, data_knock_out,
                                call_sell, put_sell, put_buy, call_buy):
    """
    Cupom Alto Retorno (CAR) genérico:
      - Cupom fixo + participação na alta até certa barreira de KO.
      - Se barreira de alta for tocada: fica só com o cupom.
      - Se barreira de baixa (KI) for tocada: exposto à variação do ativo.
      - Se ambas forem tocadas: também exposto à variação do ativo.
    """
    model_name = "cupom_alto_retorno"
    barreira_flag = (data_knock_in is not None) or (data_knock_out is not None)
    result_bruto_pct = None

    if preco_ref is None or preco_merc is None:
        return model_name, barreira_flag, result_bruto_pct

    # KO% e KI% a partir das colunas KnockOut / KnockIn
    ko_pcts = []
    ki_pcts = []
    if "KnockOut" in g.columns:
        for v in g["KnockOut"]:
            p = parse_pct(v)
            if p is not None and p > 0:
                ko_pcts.append(p)
    if "KnockIn" in g.columns:
        for v in g["KnockIn"]:
            p = parse_pct(v)
            if p is not None and p > 0:
                ki_pcts.append(p)

    ko_pct = max(ko_pcts) if ko_pcts else None
    ki_pct = min(ki_pcts) if ki_pcts else None  # normalmente barreira baixa

    # Ganho Máximo e nocional para inferir cupom fixo
    nocional = get_group_common_value(g, "Nocional")
    ganho_max = get_group_common_value(g, "Ganho Máximo")
    ganho_max_pct = None
    if nocional and ganho_max and float(nocional) != 0:
        try:
            ganho_max_pct = float(ganho_max) / float(nocional)
        except Exception:
            ganho_max_pct = None

    cupom_fixo_pct = 0.0
    participacao_cap_pct = 0.0

    if ganho_max_pct is not None and ko_pct is not None:
        # Suposição: ganho máximo = cupom fixo + (KO - 100%)
        participacao_cap_pct = ko_pct - 1.0
        cupom_fixo_pct = ganho_max_pct - participacao_cap_pct
        if cupom_fixo_pct < 0:
            cupom_fixo_pct = 0.0
    elif ganho_max_pct is not None:
        cupom_fixo_pct = ganho_max_pct

    price_ratio = preco_merc / preco_ref
    ret_ativo_pct = price_ratio - 1.0

    tocou_ki = data_knock_in is not None
    tocou_ko = data_knock_out is not None

    if not tocou_ki and not tocou_ko:
        # Nenhuma barreira: cupom + participação na alta limitada
        cap = participacao_cap_pct if participacao_cap_pct > 0 else ret_ativo_pct
        extra_up_pct = max(0.0, min(ret_ativo_pct, cap))
        result_bruto_pct = cupom_fixo_pct + extra_up_pct
    elif tocou_ko and not tocou_ki:
        # Só KO: fica apenas com o cupom fixo
        result_bruto_pct = cupom_fixo_pct
    else:
        # Tocou barreira baixa (com ou sem KO): exposto ao ativo
        result_bruto_pct = ret_ativo_pct

    return model_name, barreira_flag, result_bruto_pct


STRATEGY_MODELS = {
    "stock_or_coupon": strategy_stock_or_coupon,
    "fence_ki": strategy_fence_ki,
    "smart_hedge": strategy_smart_hedge,
    "aceleradora_ko": strategy_aceleradora_ko,
    "twip": strategy_twip,
    "smart_up": strategy_smart_up,
    "compra_put": strategy_compra_put,
    "compra_put_spread": strategy_compra_put_spread,
    "compra_call_spread": strategy_compra_call_spread,
    "financiamento_ativo": strategy_financiamento_ativo,
    "cupom_alto_retorno": strategy_cupom_alto_retorno,
}



NOME_MAP = {
    "stock or coupon": "stock_or_coupon",
    "fence ki com ativo": "fence_ki",
    "smart hedge": "smart_hedge",
    "aceleradora ko com ativo": "aceleradora_ko",
    "twin win protected - twip": "twip",
    "smart up": "smart_up",
    "compra de put": "compra_put",
    "compra put": "compra_put",
    "compra de put spread": "compra_put_spread",
    "compra put spread": "compra_put_spread",
    "compra de call spread": "compra_call_spread",
    "compra call spread": "compra_call_spread",
    "financiamento com ativo": "financiamento_ativo",
    "cupom alto retorno": "cupom_alto_retorno",
    "cupom de alto retorno": "cupom_alto_retorno",
}



def apply_strategy_model(nome_prod, g, preco_ref, preco_merc, variation,
                         data_knock_in, data_knock_out,
                         call_sell, put_sell, put_buy, call_buy):
    """
    Decide qual modelo usar, combinando:
    - padrão das pernas (inferência)
    - nome do produto (como apoio)
    - fallback genérico, se nada bater
    """
    warnings = []

    legs = build_legs_from_group(g)
    inferred = infer_strategy_from_legs(legs)
    nome_norm = normalize_prod_name(nome_prod)
    key = None

    if inferred is not None:
        key = inferred
        # Se o nome do produto não "bate" com o que foi inferido, só avisamos.
        mapped_from_nome = NOME_MAP.get(nome_norm, nome_norm)
        if nome_norm and key != mapped_from_nome:
            warnings.append(
                f"Nome do produto ('{nome_prod}') difere do padrão inferido pelas pernas ('{inferred}')."
            )
    else:
        key = NOME_MAP.get(nome_norm)

    model_fn = STRATEGY_MODELS.get(key)
    if model_fn is None:
        # Fallback: sem modelo específico, não chuta payoff
        barreira_flag = (data_knock_in is not None) or (data_knock_out is not None)
        model_name = None
        result_bruto_pct = None
        warnings.append(f"Modelo de cálculo não implementado para '{nome_prod}'.")
        model_found = False
    else:
        model_name, barreira_flag, result_bruto_pct = model_fn(
            g, preco_ref, preco_merc, variation,
            data_knock_in, data_knock_out,
            call_sell, put_sell, put_buy, call_buy
        )
        model_found = True

    return model_name, barreira_flag, result_bruto_pct, model_found, warnings


# ---------------------- Cálculo por operação (refatorado) ---------------------- #

def compute_result_for_group(g: pd.DataFrame) -> dict:
    """
    Versão refatorada: usa um "motor" de estratégias com detecção por padrão de pernas
    e fallback genérico caso o produto não seja reconhecido.
    """
    nome_prod = get_group_common_value(g, "Nome do Produto")
    conta = get_group_common_value(g, "Conta")
    cod_prod = get_group_common_value(g, "Código do produto")
    ativo = get_group_common_value(g, "Ativo")
    custodia = get_group_common_value(g, "Custódia")
    data_reserva = parse_date(get_group_common_value(g, "Data de Reserva"))
    data_fixing = parse_date(get_group_common_value(g, "Data de Fixing"))
    data_venc = parse_date(get_group_common_value(g, "Data de Vencimento"))
    preco_ref = get_group_common_value(g, "Preço de Referência (R$)")
    preco_merc = get_group_common_value(g, "Preço de Mercado Atual (R$)")
    data_knock_in = parse_date(get_group_common_value(g, "Data de Knock In"))
    data_knock_out = parse_date(get_group_common_value(g, "Data de Knock Out"))

    preco_ref = float(preco_ref) if preco_ref is not None and not pd.isna(preco_ref) else None
    preco_merc = float(preco_merc) if preco_merc is not None and not pd.isna(preco_merc) else None

    quant = get_group_common_value(g, "Quantidade")
    lotes = get_group_common_value(g, "Quantidade Reservada (lotes)")
    used_lotes = False
    if quant is None or pd.isna(quant) or quant == 0:
        if lotes is not None and not pd.isna(lotes) and lotes != 0:
            quant = float(lotes) * 100.0
            used_lotes = True
        else:
            quant = None
    else:
        quant = float(quant)

    warnings = []
    if preco_ref is None or preco_merc is None:
        warnings.append("Preço de Referência ou Preço de Mercado Atual ausente.")
    if quant is None:
        warnings.append("Quantidade e Quantidade Reservada (lotes) ausentes.")
    elif used_lotes:
        warnings.append("Quantidade estava vazia; foi utilizado Quantidade Reservada (lotes) * 100.")

    variation = None
    if preco_ref and preco_merc:
        variation = preco_merc / preco_ref - 1.0

    # Pernas "clássicas" por tipo/direção – ainda usamos para os modelos
    call_sell = get_leg(g, "CALL", "SELL")
    put_sell = get_leg(g, "PUT", "SELL")
    put_buy = get_leg(g, "PUT", "BUY")
    call_buy = get_leg(g, "CALL", "BUY")

    model_name, barreira_flag, result_bruto_pct, model_found, extra_warnings = apply_strategy_model(
        nome_prod=nome_prod,
        g=g,
        preco_ref=preco_ref,
        preco_merc=preco_merc,
        variation=variation,
        data_knock_in=data_knock_in,
        data_knock_out=data_knock_out,
        call_sell=call_sell,
        put_sell=put_sell,
        put_buy=put_buy,
        call_buy=call_buy,
    )
    warnings.extend(extra_warnings)

    financeiro = None
    if preco_ref is not None and quant is not None:
        financeiro = preco_ref * quant

    fee_spread = 0.0116  # 1,16%
    result_liq_pct = None
    result_bruto_val = None
    result_liq_val = None
    if result_bruto_pct is not None:
        result_liq_pct = result_bruto_pct - fee_spread
        if financeiro is not None:
            result_bruto_val = financeiro * result_bruto_pct
            result_liq_val = financeiro * result_liq_pct

    return {
        "Conta": conta,
        "Código do produto": cod_prod,
        "Nome do Produto": nome_prod,
        "Ativo": ativo,
        "Custódia": custodia or "",
        "Data de Reserva": data_reserva,
        "Data de Fixing": data_fixing,
        "Data de Vencimento": data_venc,
        "Preço de Referência (R$)": preco_ref,
        "Preço de Mercado Atual (R$)": preco_merc,
        "Quantidade": quant,
        "Financeiro de entrada (R$)": financeiro,
        "Resultado bruto (%)": result_bruto_pct,
        "Resultado líquido (%)": result_liq_pct,
        "Resultado bruto (R$)": result_bruto_val,
        "Resultado líquido (R$)": result_liq_val,
        "Data de Knock In": data_knock_in,
        "Data de Knock Out": data_knock_out,
        "Barreira atingida": "SIM" if barreira_flag else "NÃO",
        "Modelo encontrado": model_found,
        "Avisos": "; ".join(warnings) if warnings else "",
    }

# ---------------------- HTML no formato do exemplo ---------------------- #

def build_multiop_table(df_ops: pd.DataFrame) -> str:
    """
    Monta UMA tabela azul, com a 1ª coluna de rótulos (Início, Vencimento, etc.)
    e as demais colunas com os dados de cada operação (Op 1, Op 2, ...),
    SEM linha de título nas colunas.
    """
    ops = df_ops.reset_index(drop=True)

    html = []
    html.append(
        '<table style="border-collapse:collapse;'
        'font-family:Calibri,Arial,sans-serif;font-size:16px;'
        'margin-bottom:16px;">'
    )

    def get_barreira(row):
        val = row.get("Barreira atingida", "")
        if not val:
            return "NÃO"
        return str(val)


    def add_row(label, getter, money=False, pct=False, yellow=False):
        """Linha com 1 rótulo + 1 célula por operação."""
        bg_label = "#002b5c" if not yellow else "#ffff00"
        color_label = "#ffffff" if not yellow else "#000000"

        html.append("<tr>")
        html.append(
            f'<td style="border:1px solid #000000;padding:4px 6px;'
            f'background-color:{bg_label};color:{color_label};'
            f'font-weight:bold;">{label}</td>'
        )

        for _, row in ops.iterrows():
            val = getter(row)
            if money:
                txt = fmt_money(val) if val not in ("", None) else ""
                align = "right"
            elif pct:
                txt = fmt_pct(val) if val not in ("", None) else ""
                align = "right"
            else:
                if isinstance(val, (datetime, date)):
                    txt = fmt_date(val)
                else:
                    if val is None or (isinstance(val, float) and math.isnan(val)):
                        txt = ""
                    else:
                        if label == "Quantidade" and isinstance(val, (int, float, np.floating)):
                            txt = f"{int(val)}"
                        else:
                            txt = str(val)
                align = "left"

            bg_val = "#ffff00" if yellow else "#ffffff"
            html.append(
                f'<td style="border:1px solid #000000;padding:4px 6px;'
                f'background-color:{bg_val};text-align:{align};">{txt}</td>'
            )

        html.append("</tr>")

    # Linhas principais
    add_row("Início", lambda r: r.get("Data de Reserva") or r.get("Data de Registro"))
    add_row("Vencimento", lambda r: r.get("Data de Vencimento"))
    add_row("Conta", lambda r: r.get("Conta"))
    add_row("Ativo", lambda r: r.get("Ativo"))
    add_row("Quantidade", lambda r: r.get("Quantidade"))
    add_row("Estrutura", lambda r: r.get("Nome do Produto"))
    add_row("Preço de Referência", lambda r: r.get("Preço de Referência (R$)"), money=True)
    add_row("Financeiro de Entrada", lambda r: r.get("Financeiro de entrada (R$)"), money=True)
    add_row("Barreira atingida", get_barreira)
    add_row("Resultado Bruto", lambda r: r.get("Resultado bruto (%)"), pct=True)
    add_row("Res. Líquido (corretagem e custo B3 estimado, caso venda)",
            lambda r: r.get("Resultado líquido (%)"), pct=True)
    add_row("Res. Líquido (corretagem e custo B3 estimado)",
            lambda r: r.get("Resultado líquido (R$)"), money=True)
    add_row("Montado com compra ou sob custódia?", lambda r: r.get("Custódia"))

    # Linha amarela final
    add_row("Vender ativo? Responder", lambda r: "Sim ou Não?", yellow=True)

    html.append("</table>")
    return "".join(html)



def build_email_body(assessor_nome: str | None,
                     df_ops: pd.DataFrame,
                     hoje_str: str,
                     signature_cid: str | None = None) -> str:
    """Corpo do e-mail: texto + UMA tabela multi-operação + assinatura opcional."""
    table_html = build_multiop_table(df_ops)
    saudacao_nome = f" {assessor_nome}" if assessor_nome else ""

    sig_html = ""
    if signature_cid:
        sig_html = f'<p><img src="cid:{signature_cid}"></p>'

    body = f"""
<html>
  <body style="font-family:Calibri,Arial,sans-serif;font-size:16px;color:#000000;">
    <p>Prezado(a){saudacao_nome}, tudo bem?</p>
    <p>Temos a seguinte(s) Operação(ões) Estruturada(s) fixando hoje (<b>{hoje_str}</b>):</p>
    {table_html}
    <p>Solicitamos que responda quanto à venda ou não das ações.</p>
    <p>Ficamos no aguardo da resposta para atuação. Lembrando que por padrão as ações serão
       <b style="color:red;">VENDIDAS</b>, caso o cliente opte por não vender, favor informar neste fluxo até as 15:00 horas.</p>
    <p>Dúvidas, à disposição.</p>
    <p><i>*Obs – O cálculo de resultado inclui dividendos recebidos no período.</i></p>
    {sig_html}
  </body>
</html>
"""
    return body



# ---------------------- Envio via Outlook ---------------------- #

def enviar_emails(df_final: pd.DataFrame):
    """Cria e envia 1 e-mail por assessor, com 1 cartão por operação dentro."""
    if win32 is None:
        print("Biblioteca 'pywin32' não encontrada. Instale com: pip install pywin32")
        pause()
        return
    
    base_dir = get_base_dir()
    sig_path = find_signature_image(base_dir)
    PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

    email_assessor_col = (
        find_col(df_final, ["email", "assessor"])
        or find_col(df_final, ["email", "assessores"])
        or "E-mail Assessor"
    )
    email_lider_col = (
        find_col(df_final, ["email", "lider"])
        or find_col(df_final, ["email", "líder"])
    )
    assessor_nome_col = find_col(df_final, ["assessor"])

    print(f"Coluna identificada para e-mail do assessor: {email_assessor_col}")
    if email_lider_col:
        print(f"Coluna identificada para e-mail do líder: {email_lider_col}")

    if email_assessor_col not in df_final.columns:
        print("Não encontrei a coluna de e-mail do assessor (ex.: 'E-mail Assessor') após varredura.")
        print("Colunas disponíveis são:")
        for c in df_final.columns:
            print(" -", c)
        pause()
        return

    total_ops = len(df_final)
    print(f"Total de operações com fixing hoje (após merge com Advisor Base): {total_ops}")

    col_series = df_final[email_assessor_col].astype(str).str.strip()
    df_nonempty_email = df_final[~col_series.isin(["", "nan", "none", "NaN", "None"])]

    if df_nonempty_email.empty:
        print(
            f"Encontrei {total_ops} operações com fixing hoje, "
            f"mas nenhuma com e-mail de assessor preenchido na coluna '{email_assessor_col}'."
        )
        try:
            contas = df_final["Conta"].astype(str).unique().tolist()
            print("Contas com operações hoje:", ", ".join(contas[:20]))
        except Exception:
            pass
        pause()
        return

    outlook = win32.Dispatch("Outlook.Application")
    hoje_str = date.today().strftime("%d/%m/%Y")

    enviados = 0
    for assessor_email, grupo in df_nonempty_email.groupby(email_assessor_col):
        if assessor_email is None:
            continue
        assessor_email = str(assessor_email).strip()
        if not assessor_email:
            continue

        email_lider = None
        if email_lider_col and email_lider_col in grupo.columns:
            vals = grupo[email_lider_col].dropna().astype(str).unique()
            if len(vals) > 0:
                email_lider = vals[0].strip()

        assessor_nome = None
        if assessor_nome_col and assessor_nome_col in grupo.columns:
            vals = grupo[assessor_nome_col].dropna().astype(str).unique()
            if len(vals) > 0:
                assessor_nome = vals[0].strip()

        subject = f"Fixing de operações estruturadas - {hoje_str}"

        mail = outlook.CreateItem(0)

        # Assinatura inline (se existir arquivo de assinatura)
        signature_cid = None
        if sig_path and os.path.exists(sig_path):
            attach = mail.Attachments.Add(sig_path)
            attach.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, "assinatura_inline")
            signature_cid = "assinatura_inline"

        html_body = build_email_body(assessor_nome, grupo, hoje_str, signature_cid)

        mail.To = assessor_email
        if email_lider:
            mail.CC = email_lider
        mail.CC = os.getenv("FIXED_CC_EMAILS", "risk-team@example.com")
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Display()
        enviados += 1
        print(f"Email enviado para {assessor_email} (operações: {len(grupo)})")

    if enviados == 0:
        print("Nenhum e-mail foi enviado, apesar de existirem operações com fixing hoje.")
        print("Verifique se os e-mails dos assessores estão corretamente preenchidos na Advisor Base.")
    else:
        print(f"Envio concluído. Total de e-mails enviados: {enviados}")
    pause()


# ---------------------- Entrada / saída de arquivos ---------------------- #

def find_excel_file(base_dir: str, name_keywords) -> str | None:
    """Procura um arquivo Excel na pasta cujo nome contenha TODOS os termos de name_keywords."""
    files = os.listdir(base_dir)
    for f in files:
        lower = f.lower()
        if not lower.endswith(('.xlsx', '.xlsm', '.xls')):
            continue
        if all(kw.lower() in lower for kw in name_keywords):
            return os.path.join(base_dir, f)
    return None


def main():
    base_dir = get_base_dir()
    print(f"Pasta base: {base_dir}")

    # Arquivo de operações
    ops_path = find_excel_file(base_dir, ["operações", "produtos", "estruturados"]) \
               or find_excel_file(base_dir, ["operacoes", "produtos", "estruturados"])
    if not ops_path:
        print("Não encontrei o arquivo 'Structured_Products.xlsx' na mesma pasta do script.")
        pause()
        return

    # Advisor Base
    advisor_base_path = find_excel_file(base_dir, ["base", "btg"])
    if not advisor_base_path:
        print("Não encontrei o arquivo da Advisor Base (ex.: 'Advisor Base.xlsx') na mesma pasta do script.")
        pause()
        return

    print(f"Usando operações em: {ops_path}")
    print(f"Usando Advisor Base em: {advisor_base_path}")

    # Carrega operações
    df_ops = pd.read_excel(ops_path, sheet_name="Operações de Produtos Estrutura")

    # Filtra fixing = hoje
    df_ops["__fixing_date__"] = pd.to_datetime(df_ops["Data de Fixing"], dayfirst=True, errors="coerce").dt.date
    hoje = date.today()
    df_today = df_ops[df_ops["__fixing_date__"] == hoje].copy()

    if df_today.empty:
        print(f"Não há operações com Data de Fixing = {hoje.strftime('%d/%m/%Y')}.")
        pause()
        return

    # Calcula resultado por (Conta, Código do produto)
    results = []
    for (conta, cod), g in df_today.groupby(["Conta", "Código do produto"]):
        results.append(compute_result_for_group(g))

    df_res = pd.DataFrame(results)
    if df_res.empty:
        print("Nenhuma operação válida encontrada para hoje.")
        pause()
        return

    # Carrega Advisor Base e faz merge pelas contas
    df_advisor_base = pd.read_excel(advisor_base_path)
    conta_col_btg = find_col(df_advisor_base, ["conta"]) or "Conta"
    if conta_col_btg not in df_advisor_base.columns:
        print("Não encontrei coluna 'Conta' na Advisor Base.")
        print("Colunas disponíveis:")
        for c in df_advisor_base.columns:
            print(" -", c)
        pause()
        return

    df_advisor_base["Conta_norm"] = normalize_conta_series(df_advisor_base[conta_col_btg])
    df_res["Conta_norm"] = normalize_conta_series(df_res["Conta"])

    df_final = pd.merge(
        df_res,
        df_advisor_base,
        on="Conta_norm",
        how="left",
        suffixes=("", "_btg")
    )
    df_final["Conta"] = df_res["Conta"].values

    print("Algumas contas e e-mails depois do merge:")
    cols_debug = [c for c in df_final.columns if "Conta" in c or "mail" in c.lower()]
    print(df_final[cols_debug].head(10))

    # Aviso de estruturas não mapeadas
    sem_modelo = df_final[~df_final["Modelo encontrado"]]
    if not sem_modelo.empty:
        print("Atenção: existem operações cujo modelo de cálculo não foi implementado.")
        for _, row in sem_modelo.iterrows():
            print(f"  Conta {row['Conta']} - Produto {row['Nome do Produto']} (cód {row['Código do produto']})")

    # Envia os e-mails
    enviar_emails(df_final)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        _log_unhandled_exception("run_err")
        raise
