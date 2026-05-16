# -*- coding: utf-8 -*-

# PUBLIC-SAFE PORTFOLIO VERSION
# Real credentials, client data, private endpoints and internal paths were removed or replaced by placeholders.
import argparse
import zipfile
from typing import Dict, Optional, Tuple

import pandas as pd

# -----------------------------
# PARSERS (B3 + BR)
# -----------------------------
def venc_to_yyyymmdd(venc: str) -> str:
    """
    Aceita '16/01/2026' ou '2026-01-16' ou '20260116' e retorna '20260116'
    """
    s = venc.strip()
    if s.isdigit() and len(s) == 8:
        return s
    s = s.replace("-", "/")
    parts = s.split("/")
    if len(parts) == 3:
        dd, mm, yyyy = parts
        dd = dd.zfill(2)
        mm = mm.zfill(2)
        yyyy = yyyy.zfill(4)
        return f"{yyyy}{mm}{dd}"
    raise ValueError(f"Vencimento inválido: {venc!r}. Use dd/mm/aaaa ou yyyymmdd.")

import re
from typing import Optional

def suffix_from_shareclass(shareclass: str) -> Optional[str]:
    """
    Retorna sufixo B3 (3/4/5/6/7/8/11) com base em:
    - ON/PN/PNA/PNB/PNC/PND/UNT/UNIT
    - OU fallback: pega número no final de tokens como EJ11 / ED3 / EJ4
    Só aceita números válidos.
    """
    if not shareclass:
        return None

    sc = str(shareclass).upper().strip()

    # pega primeiro token (antes de espaços ou "/")
    tok = re.split(r"[\s/]+", sc)[0].strip()

    mapping = {
        "ON": "3",
        "PN": "4",
        "PNA": "5",
        "PNB": "6",
        "PNC": "7",
        "PND": "8",
        "UNT": "11",
        "UNIT": "11",
    }
    if tok in mapping:
        return mapping[tok]

    # fallback: EJ11 / ED3 / EJ4 -> pega número final
    m = re.search(r"(\d{1,2})$", tok)
    if m:
        num = m.group(1)
        if num in {"3", "4", "5", "6", "7", "8", "11"}:
            return num

    return None

def build_base_ticker(mnemonico: str, shareclass: str) -> str:
    m = (mnemonico or "").strip().upper()
    suf = suffix_from_shareclass(shareclass or "")
    return f"{m}{suf}" if suf else m

def load_b3_mapping_from_zip(zip_path: str, venc_yyyymmdd: str) -> Dict[str, str]:
    """
    Lê SI_D_SEDE.txt dentro do zip e retorna dict {COD_OPCAO: ATIVO_BASE_TICKER}
    filtrando pelo vencimento (campo vencimento no arquivo).
    """
    with zipfile.ZipFile(zip_path) as zf:
        # normalmente é esse nome, mas vamos ser tolerantes
        txt_name = None
        for n in zf.namelist():
            if n.lower().endswith(".txt"):
                txt_name = n
                break
        if not txt_name:
            raise FileNotFoundError("Não achei .txt dentro do zip.")

        mapping: Dict[str, str] = {}

        with zf.open(txt_name) as f:
            for raw in f:
                # o arquivo costuma ser latin1/ansi
                line = raw.decode("latin1", errors="ignore").rstrip("\r\n")
                if not line.startswith("02|"):
                    continue

                parts = line.split("|")
                # no seu arquivo, as linhas 02 têm 19 campos
                if len(parts) < 18:
                    continue

                # índices observados no seu arquivo:
                # parts[6]  = mnemônico do ativo base (ex.: 'ABEV', 'BEEF')
                # parts[7]  = classe (ex.: 'ON      NM')
                # parts[13] = código da opção (ex.: 'ABEVA128')
                # parts[17] = vencimento (ex.: '20260116')
                mnemonico = parts[6].strip().upper()
                shareclass = parts[7].strip()
                cod_opcao = parts[13].strip().upper().replace(" ", "")  # remove espaços internos
                venc = parts[17].strip()

                if not cod_opcao:
                    continue
                if venc != venc_yyyymmdd:
                    continue

                # ✅ VALIDAÇÃO: opção normalmente começa com o mnemônico (ABEVxxxx, PRIOxxxx)
                # Se não bater, ignora pra não contaminar mapping (evita "ABEV" em "PRIO...")
                mn4 = mnemonico.replace(" ", "")
                if mn4 and not cod_opcao.startswith(mn4):
                    # Se quiser depurar, descomente:
                    # print("SKIP mismatch:", cod_opcao, "mnemonico:", mnemonico, "shareclass:", shareclass)
                    continue

                ativo_base = build_base_ticker(mnemonico, shareclass)

                # ✅ Se já existir (duplicata), não deixa sobrescrever por outro valor
                if cod_opcao in mapping and mapping[cod_opcao] != ativo_base:
                    # Se quiser depurar, descomente:
                    # print("DUP conflito:", cod_opcao, mapping[cod_opcao], "->", ativo_base)
                    continue

                mapping[cod_opcao] = ativo_base

        return mapping

# -----------------------------
# EXCEL (xlwings)
# -----------------------------
def find_listobject_by_name(book, table_name: str):
    """
    Procura ListObject (tabela) em todas as planilhas.
    Retorna (sheet, listobject_api) ou (None, None).
    """
    for sht in book.sheets:
        try:
            lo = sht.api.ListObjects(table_name)
            return sht, lo
        except Exception:
            pass
    return None, None

def read_table_to_df(lo_api) -> pd.DataFrame:
    """
    Lê a ListObject inteira (inclui header) e devolve DataFrame.
    """
    rng = lo_api.Range  # inclui header
    values = rng.Value
    if not values or len(values) < 2:
        return pd.DataFrame()

    header = list(values[0])
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)
    return df

def write_column(lo_api, col_name: str, values_1d):
    """
    Escreve uma coluna (DataBodyRange) na ListObject.
    """
    col = lo_api.ListColumns(col_name)
    body = col.DataBodyRange
    # xlwings/win32 espera lista de listas (vertical)
    body.Value = [[v] for v in values_1d]

def norm_op_code(v) -> str:
    s = "" if v is None else str(v)
    # remove NBSP (caractere invisível comum do Excel)
    s = s.replace("\xa0", " ")
    # remove todos os espaços
    s = s.strip().upper().replace(" ", "")
    return s



def fill_ativo_base_in_open_excel(
    venc_ddmmyyyy: str,
    zip_path: str,
    workbook_name_hint: Optional[str] = None,
    table_name: str = "posições",
    col_opcao: str = "Produto",
    col_out: str = "Ativo base",
    col_venc: str = "Data Exercício",
):
    venc_yyyymmdd = venc_to_yyyymmdd(venc_ddmmyyyy)
    mapping = load_b3_mapping_from_zip(zip_path, venc_yyyymmdd)

    if not mapping:
        raise RuntimeError(f"Mapping vazio para vencimento {venc_yyyymmdd}. Verifique o arquivo da B3.")

    import xlwings as xw

    app = xw.apps.active
    if app is None:
        raise RuntimeError("Não encontrei Excel aberto. Abra o arquivo 'Operations Workbook' e rode novamente.")

    # tenta pegar workbook ativo ou pelo hint
    book = app.books.active
    if workbook_name_hint:
        for b in app.books:
            if workbook_name_hint.lower() in b.name.lower():
                book = b
                break

    sht, lo = find_listobject_by_name(book, table_name)
    if lo is None:
        raise RuntimeError(f"Não encontrei a tabela '{table_name}' no workbook aberto.")

    # ---------- helpers locais ----------
    def col_values(colname: str):
        """Lê uma coluna da tabela (DataBodyRange) e devolve lista 1D."""
        try:
            rng = lo.ListColumns(colname).DataBodyRange
        except Exception:
            raise RuntimeError(f"Coluna '{colname}' não existe na tabela '{table_name}'.")

        if rng is None:
            return []

        v = rng.Value  # pode ser: tuple(tuple()), tuple(), ou valor único
        if v is None:
            return []
        if isinstance(v, tuple):
            # geralmente vem como ((x,), (y,), ...)
            if len(v) > 0 and isinstance(v[0], tuple):
                return [row[0] for row in v]
            return list(v)
        return [v]

    def any_to_yyyymmdd(v) -> Optional[str]:
        """Converte valores vindos do Excel (datetime, texto, serial) em YYYYMMDD."""
        if v is None:
            return None

        # datetime/date (Excel geralmente entrega datetime)
        try:
            import datetime as _dt
            if isinstance(v, (_dt.datetime, _dt.date)):
                return v.strftime("%Y%m%d")
        except Exception:
            pass

        # serial excel (às vezes vem float/int)
        if isinstance(v, (int, float)):
            try:
                # Excel base date: 1899-12-30
                import datetime as _dt
                base = _dt.datetime(1899, 12, 30)
                dt = base + _dt.timedelta(days=float(v))
                return dt.strftime("%Y%m%d")
            except Exception:
                return None

        # string
        s = str(v).strip()
        if not s:
            return None

        # corta hora se existir
        if " " in s:
            s = s.split(" ")[0]

        s = s.replace("-", "/")

        if s.isdigit() and len(s) == 8:
            return s

        parts = s.split("/")
        if len(parts) == 3:
            dd, mm, yyyy = parts
            if len(yyyy) == 2:
                yyyy = "20" + yyyy
            return f"{yyyy.zfill(4)}{mm.zfill(2)}{dd.zfill(2)}"

        return None
    # ---------- fim helpers ----------

    # lê colunas necessárias direto do Excel (sem pandas)
    venc_list = col_values(col_venc)
    prod_list = col_values(col_opcao)
    out_list = col_values(col_out)

    if not venc_list:
        raise RuntimeError(f"A tabela '{table_name}' parece não ter linhas (DataBodyRange vazio).")

    # normaliza tamanhos (em caso de coluna vazia)
    n = len(venc_list)
    if len(prod_list) != n:
        raise RuntimeError(f"Tamanho da coluna '{col_opcao}' ({len(prod_list)}) diferente de '{col_venc}' ({n}).")
    if len(out_list) != n:
        # se a coluna "Ativo base" estiver vazia e vier menor, ajusta
        out_list = (out_list + [None] * n)[:n]

    linhas_venc = 0
    mapeados = 0

    for i in range(n):
        v_yyyymmdd = any_to_yyyymmdd(venc_list[i])
        if v_yyyymmdd != venc_yyyymmdd:
            continue

        linhas_venc += 1

        cod_op = str(prod_list[i]).strip().upper()
        if not cod_op:
            continue

        base = mapping.get(cod_op)
        if base:
            out_list[i] = base
            mapeados += 1

    # escreve de volta na coluna Ativo base
    write_column(lo, col_out, out_list)

    return {
        "workbook": book.name,
        "table": table_name,
        "vencimento": venc_ddmmyyyy,
        "linhas_total": n,
        "linhas_venc": linhas_venc,
        "mapeados": mapeados,
        "mapping_size": len(mapping),
    }

# -----------------------------
# EXPORT LOOKUP (excel bonitinho)
# -----------------------------
def export_lookup_xlsx(zip_path: str, out_xlsx: str):
    """
    Exporta todas as opções do arquivo para um XLSX com colunas úteis:
    COD_OPCAO, ATIVO_BASE, VENCIMENTO, STRIKE, TIPO
    """
    rows = []
    with zipfile.ZipFile(zip_path) as zf:
        txt_name = next((n for n in zf.namelist() if n.lower().endswith(".txt")), None)
        if not txt_name:
            raise FileNotFoundError("Não achei .txt dentro do zip.")
        with zf.open(txt_name) as f:
            for raw in f:
                line = raw.decode("latin1", errors="ignore").rstrip("\r\n")
                if not line.startswith("02|"):
                    continue
                parts = line.split("|")
                if len(parts) < 18:
                    continue
                tipo_desc = parts[3].strip().upper()  # "OPCOES COMPRA"/"OPCOES VENDA"
                mnemonico = parts[6].strip()
                shareclass = parts[7].strip()
                cod_opcao = parts[13].strip().upper()
                strike = parts[16].strip()
                venc = parts[17].strip()
                ativo_base = build_base_ticker(mnemonico, shareclass)
                rows.append((cod_opcao, ativo_base, venc, strike, tipo_desc))

    df = pd.DataFrame(rows, columns=["COD_OPCAO", "ATIVO_BASE", "VENCIMENTO", "STRIKE", "TIPO"])
    df.drop_duplicates(subset=["COD_OPCAO"], inplace=True)
    df.sort_values(["VENCIMENTO", "COD_OPCAO"], inplace=True)
    df.to_excel(out_xlsx, index=False)

# -----------------------------
# CLI
# -----------------------------
def main():
    ap = argparse.ArgumentParser()

    ap.add_argument(
        "--zip",
        default=r"data\SI_D_SEDE_sample.zip",
        help="Caminho do .zip da B3 (SI_D_SEDE.zip)"
    )
    ap.add_argument(
        "--venc",
        default="16/01/2026",
        help="Vencimento dd/mm/aaaa (ex: 16/01/2026)"
    )
    ap.add_argument(
        "--workbook",
        default="Operations Workbook",
        help="Parte do nome do workbook aberto no Excel"
    )
    ap.add_argument(
        "--export",
        default=None,
        help="Se quiser gerar lookup em XLSX: informe o caminho do arquivo de saída"
    )

    args = ap.parse_args()

    # resto igual...


    if args.export:
        export_lookup_xlsx(args.zip, args.export)
        print(f"OK: Lookup exportado para: {args.export}")

    if args.venc:
        info = fill_ativo_base_in_open_excel(
            venc_ddmmyyyy=args.venc,
            zip_path=args.zip,
            workbook_name_hint=args.workbook
        )
        print("OK:", info)

if __name__ == "__main__":
    main()
