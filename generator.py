# generator.py
#
# CAMBIOS IMPORTANTES:
# - Moneda en DGI se toma tal cual (ej: UYU, USD, BRL...) y se mapea con /Datos/Monedas.xlsx.
# - En el TXT, el campo Moneda es el "Digito" (0-9) definido por el usuario.
# - Caja por moneda sale de Parámetros.xlsx: "Caja <DGI>".
# - Haber por moneda sale de Proveedores.xlsx: "Haber <DGI>".
#
# FIX CRÍTICO:
# - Nota de crédito NO se hace con "montos negativos" en el mismo lado,
#   porque rompe asientos en muchos importadores.
# - Para Nota de crédito se INVIERTE el asiento:
#     Debe <-> Haber (y montos positivos).
#
# NUEVO:
# - precheck_missing_currencies(): detecta monedas del DGI que no están activas/configuradas.
# - Si faltan monedas configuradas -> RuntimeError("FALTAN_MONEDAS_EN_CONFIG")

import unicodedata
from difflib import SequenceMatcher
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple
import shutil
import pandas as pd


STOPWORDS_ABREV = {
    "de", "del", "la", "el", "los", "las", "y", "a", "por", "para", "en",
    "un", "una", "unos", "unas",
}

PROV_BASE_COLS = ["RUT", "Nombre", "Debe", "IVA Fijo", "Libro", "Cont/Cred"]
PARAM_BASE_COLS = ["IVA 10", "IVA 22", "IVA GEN", "REDONDEOS", "RETENCIONES"]


def _norm_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ").strip().lower()
    s = " ".join(s.split())
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return s


def _clean_excel_str(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.strip()


def _norm_dgi_code(s: str) -> str:
    return _clean_excel_str(s).upper()


def _auto_abbrev_from_tipo(tipo_cfe: str, max_len: int = 6) -> str:
    t = _norm_text(tipo_cfe)
    if not t:
        return "DOC"
    parts = []
    for chunk in t.replace("-", " ").split():
        chunk = chunk.strip()
        if not chunk or chunk in STOPWORDS_ABREV:
            continue
        parts.append(chunk)
    if not parts:
        return "DOC"
    initials = "".join(p[0].upper() for p in parts if p)
    return initials[:max_len] if initials else "DOC"


def load_proveedores_xlsx_robusto(path: Path) -> pd.DataFrame:
    for h in range(0, 30):
        try:
            df = pd.read_excel(path, dtype=str, header=h).fillna("")
            df.columns = [str(c).strip() for c in df.columns]
            if "RUT" in df.columns and "Nombre" in df.columns:
                for c in PROV_BASE_COLS:
                    if c not in df.columns:
                        df[c] = ""
                df["RUT"] = df["RUT"].astype(str).str.strip()
                df["Nombre"] = df["Nombre"].astype(str).str.strip()
                df = df[(df["RUT"] != "") | (df["Nombre"] != "")].copy()
                df.to_excel(path, index=False)
                return df
        except Exception:
            pass

    cols = PROV_BASE_COLS
    df = pd.read_excel(path, dtype=str, header=None).fillna("")
    while df.shape[1] < len(cols):
        df[df.shape[1]] = ""
    df = df.iloc[:, :len(cols)].copy()
    df.columns = cols
    df["RUT"] = df["RUT"].astype(str).str.strip()
    df["Nombre"] = df["Nombre"].astype(str).str.strip()
    df = df[(df["RUT"] != "") | (df["Nombre"] != "")].copy()
    df.to_excel(path, index=False)
    return df


class ContyGenerator:
    def __init__(self, company_dir, dgi_xls):
        self.company_dir = Path(company_dir)
        self.dgi_xls = Path(dgi_xls)

        self.data_dir = self.company_dir / "Datos"
        self.proveedores_xlsx = self.data_dir / "Proveedores.xlsx"
        self.parametros_xlsx = self.data_dir / "Parámetros.xlsx"
        self.abreviaturas_xlsx = self.data_dir / "Abreviaturas.xlsx"
        self.monedas_xlsx = self.data_dir / "Monedas.xlsx"

        today = datetime.now().strftime("%Y%m%d")
        self.output_dir = self.data_dir / f"TXTCFE_{today}"

    def prepare_output_dir(self, clear_if_exists: bool = True) -> None:
        if clear_if_exists and self.output_dir.exists():
            shutil.rmtree(self.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def _read_monedas_map(self) -> Dict[str, str]:
        df = pd.read_excel(self.monedas_xlsx, dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        for c in ["DGI", "Digito", "Activa"]:
            if c not in df.columns:
                raise ValueError("Monedas.xlsx inválido: faltan columnas DGI/Digito/Activa.")

        df["DGI"] = df["DGI"].astype(str).apply(_norm_dgi_code)
        df["Digito"] = df["Digito"].astype(str).str.strip()
        df["Activa"] = df["Activa"].astype(str).str.strip().str.lower()

        active = df[df["Activa"].isin(["si", "sí", "s", "y", "yes", "1", "true"])].copy()

        digits = active["Digito"].tolist()
        for d in digits:
            if not (len(d) == 1 and d.isdigit()):
                raise ValueError("Monedas.xlsx: hay monedas activas con dígito TXT inválido (debe ser 0-9).")
        if len(set(digits)) != len(digits):
            raise ValueError("Monedas.xlsx: hay dígitos TXT repetidos entre monedas activas.")

        out = {r["DGI"]: r["Digito"] for _, r in active.iterrows() if r["DGI"]}
        if "UYU" not in out or "USD" not in out:
            raise ValueError("Monedas.xlsx: UYU y USD deben estar activas.")
        return out

    def precheck_missing_currencies(self) -> List[str]:
        df_dgi = self._read_dgi()
        monedas = sorted(set(_norm_dgi_code(m) for m in df_dgi["Moneda"].tolist()))
        monedas = [m for m in monedas if m]
        mapa = self._read_monedas_map()
        missing = [m for m in monedas if m not in mapa]
        return missing

    def _write_proveedores_no_encontrados(self, missing: List[str]) -> Path:
        out = self.data_dir / "Proveedores no encontrados.txt"
        self.data_dir.mkdir(parents=True, exist_ok=True)
        content = "\n".join(sorted(set(missing))) + ("\n" if missing else "")
        out.write_text(content, encoding="utf-8")
        return out

    def precheck_missing_ruts(self, write_file: bool = True, clear_output: bool = True) -> List[str]:
        self.prepare_output_dir(clear_if_exists=clear_output)
        df_dgi = self._read_dgi()
        df_prov = self._read_proveedores()

        missing = []
        for _, r in df_dgi.iterrows():
            rut = _clean_excel_str(r.get("RUT Emisor", "")).replace(" ", "").replace("\u00A0", "")
            rut_lookup = rut.replace(".", "")
            if not rut:
                continue
            if rut in df_prov.index or rut_lookup in df_prov.index:
                continue
            missing.append(rut)

        missing = sorted(set(missing))

        out_path = self.output_dir / "ruts_no_en_proveedores.txt"
        if write_file and missing:
            out_path.write_text("\n".join(missing) + "\n", encoding="utf-8")
        else:
            try:
                if out_path.exists():
                    out_path.unlink()
            except Exception:
                pass

        return missing

    def run(self, allow_missing: bool = False, skip_prepare: bool = False) -> Path:
        if not skip_prepare:
            self.prepare_output_dir(clear_if_exists=True)

        missing_cur = self.precheck_missing_currencies()
        if missing_cur:
            raise RuntimeError("FALTAN_MONEDAS_EN_CONFIG")

        df_dgi = self._read_dgi()
        df_prov = self._read_proveedores()
        params = self._read_parametros()
        abrev_map = self._read_abreviaturas()
        moneda_map = self._read_monedas_map()

        year, month = self._validate_single_period(df_dgi)

        missing = self.precheck_missing_ruts(write_file=True, clear_output=False)
        if missing:
            if allow_missing:
                self._write_proveedores_no_encontrados(missing)
            else:
                raise RuntimeError("FALTAN_RUTS_EN_PROVEEDORES")

        txt_path = self.company_dir / f"MG{year}{month}01.txt"
        self._generate_files(df_dgi, df_prov, params, abrev_map, moneda_map, txt_path)
        return txt_path

    def _read_dgi(self) -> pd.DataFrame:
        df = pd.read_excel(self.dgi_xls, header=8)
        df = df.iloc[1:].copy()
        df.columns = [
            "Fecha", "Fecha valor", "Tipo CFE", "Serie", "Número",
            "RUT Emisor", "Moneda", "Monto Neto", "Monto IVA",
            "Monto Total", "Monto Ret/Per", "Monto Cred. Fiscal"
        ]
        return df

    def _read_proveedores(self) -> pd.DataFrame:
        df = load_proveedores_xlsx_robusto(self.proveedores_xlsx).fillna("")
        df["RUT"] = df["RUT"].apply(_clean_excel_str).str.replace(" ", "", regex=False)
        df["Nombre"] = df["Nombre"].astype(str).str.strip()
        for c in PROV_BASE_COLS:
            if c not in df.columns:
                df[c] = ""
        return df.set_index("RUT")

    def _read_parametros(self) -> Dict[str, str]:
        df = pd.read_excel(self.parametros_xlsx, dtype=str).fillna("")
        df.columns = df.columns.astype(str).str.strip()

        missing_base = [c for c in PARAM_BASE_COLS if c not in df.columns]
        if missing_base:
            raise ValueError(f"En Parámetros faltan columnas base: {missing_base}")

        def val(c): return str(df.iloc[0][c]).strip() if len(df) else ""

        params: Dict[str, str] = {
            "IVA10": val("IVA 10"),
            "IVA22": val("IVA 22"),
            "IVAGEN": val("IVA GEN"),
            "REDONDEOS": val("REDONDEOS"),
            "RETENCIONES": val("RETENCIONES"),
        }

        for c in df.columns:
            cstr = str(c).strip()
            if cstr.upper().startswith("CAJA "):
                code = cstr.split(" ", 1)[1].strip().upper()
                if code:
                    params[f"CAJA_{code}"] = val(cstr)

        return params

    def _read_abreviaturas(self) -> Dict[str, str]:
        df = pd.read_excel(self.abreviaturas_xlsx, dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        cols_norm = {c.lower().replace(" ", ""): c for c in df.columns}

        def pick(*cands):
            for cand in cands:
                k = cand.lower().replace(" ", "")
                if k in cols_norm:
                    return cols_norm[k]
            return None

        tipo_col = pick("Tipo CFE", "Tipo")
        abrv_col = pick("Abreviado", "Abrev", "Abreviacion", "Abreviación")

        if tipo_col is None or abrv_col is None:
            raise ValueError("Abreviaturas.xlsx debe tener columnas: Tipo CFE y Abreviado.")

        mapa = {}
        for _, r in df.iterrows():
            k = _norm_text(r.get(tipo_col))
            v = str(r.get(abrv_col) or "").strip()
            if k and v:
                mapa[k] = v
        return mapa

    def _validate_single_period(self, df: pd.DataFrame) -> Tuple[str, str]:
        fechas = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=True)
        meses = fechas.dt.to_period("M").dropna().unique()
        if len(meses) != 1:
            raise ValueError("El archivo DGI contiene más de un período")
        p = meses[0]
        return f"{p.year % 100:02d}", f"{p.month:02d}"

    def _get_abbrev(self, tipo_cfe: str, abrev_map: Dict[str, str]) -> Tuple[str, str]:
        key = _norm_text(tipo_cfe)
        if key in abrev_map:
            return abrev_map[key], "MAP"

        best_key = None
        best_score = 0.0
        for k in abrev_map.keys():
            s = SequenceMatcher(None, key, k).ratio()
            if s > best_score:
                best_score = s
                best_key = k

        if best_key is not None and best_score >= 0.85:
            return abrev_map[best_key], "FUZZY"

        return _auto_abbrev_from_tipo(tipo_cfe), "AUTO"

    def _is_credit_note(self, tipo_cfe_raw: str) -> bool:
        t = _norm_text(tipo_cfe_raw)
        return ("nota de credito" in t) or ("nota de crédito" in t)

    def _generate_files(self, df_dgi, df_prov, params, abrev_map, moneda_map: Dict[str, str], txt_path: Path):
        LIMITE_INCONGRUENTE = 1.00

        for c in ["Monto Neto", "Monto IVA", "Monto Total", "Monto Ret/Per"]:
            df_dgi[c] = pd.to_numeric(df_dgi[c], errors="coerce").fillna(0.0)

        def d2(x): return f"{float(x):.2f}"
        def cot7(): return f"{0.0:.7f}"

        def day(v):
            t = pd.to_datetime(v, errors="coerce", dayfirst=True)
            return "00" if pd.isna(t) else f"{t.day:02d}"

        def join_concept(*parts):
            parts = [str(p).strip() for p in parts if str(p).strip()]
            return " ".join(parts)

        lines = []
        lines.append(";BorraExistentes = Si")
        lines.append("Dia,Debe,Haber,Concepto,Numero,RUC,Moneda,Total,CodigoIVA,IVA,Cotizacion,Libro")

        iva_raro_lines = []
        incongruentes_detalle = []

        def emit(dia_s, debe, haber, concepto, numero, ruc, moneda_digit, monto, libro_char):
            # monto siempre numérico, se escribe en "Total"
            lines.append(",".join([
                dia_s,
                debe or "",
                haber or "",
                concepto,
                numero,
                ruc or "",
                str(moneda_digit),
                d2(monto),
                "0",
                d2(0),
                cot7(),
                libro_char,
            ]))

        def post(dia_s, debe, haber, concepto, numero, ruc, moneda_digit, monto, libro_char, invert: bool):
            """
            Posteo contable:
            - En factura (invert=False): Debe/Haber normal
            - En NC (invert=True): se invierte Debe <-> Haber
            - Monto SIEMPRE positivo para máxima compatibilidad
            """
            m = abs(float(monto))
            if m < 0.000001:
                return
            if invert:
                emit(dia_s, haber, debe, concepto, numero, ruc, moneda_digit, m, libro_char)
            else:
                emit(dia_s, debe, haber, concepto, numero, ruc, moneda_digit, m, libro_char)

        for _, r in df_dgi.iterrows():
            dia = day(r.get("Fecha"))
            tipo_cfe_raw = _clean_excel_str(r.get("Tipo CFE", ""))
            serie = _clean_excel_str(r.get("Serie", ""))
            numero_doc = _clean_excel_str(r.get("Número", ""))

            moneda_dgi = _norm_dgi_code(r.get("Moneda", ""))
            if moneda_dgi not in moneda_map:
                raise RuntimeError("FALTAN_MONEDAS_EN_CONFIG")
            moneda_digit = moneda_map[moneda_dgi]

            rut_dgi = _clean_excel_str(r.get("RUT Emisor", "")).replace(" ", "").replace("\u00A0", "")
            rut_lookup = rut_dgi.replace(".", "")

            neto = float(r.get("Monto Neto", 0.0))
            iva = float(r.get("Monto IVA", 0.0))
            total = float(r.get("Monto Total", 0.0))
            retper = float(r.get("Monto Ret/Per", 0.0) or 0.0)

            is_nc = self._is_credit_note(tipo_cfe_raw)

            prov_row = None
            if rut_dgi in df_prov.index:
                prov_row = df_prov.loc[rut_dgi]
            elif rut_lookup in df_prov.index:
                prov_row = df_prov.loc[rut_lookup]

            if prov_row is None:
                continue

            nombre_prov = str(prov_row.get("Nombre", "")).strip().upper()

            libro_raw = str(prov_row.get("Libro", "")).strip().upper()
            libro_char = "E" if libro_raw == "E" else "C"

            contcred_raw = str(prov_row.get("Cont/Cred", "")).strip().lower()
            es_contado = (contcred_raw == "contado")

            debe_cta = str(prov_row.get("Debe", "")).strip()

            haber_col = f"Haber {moneda_dgi}"
            haber_prov_cta = str(prov_row.get(haber_col, "")).strip()
            if not haber_prov_cta:
                haber_prov_cta = str(prov_row.get("Haber USD", "")).strip()

            caja_cta = (params.get(f"CAJA_{moneda_dgi}") or "").strip()
            if not caja_cta:
                caja_cta = (params.get("CAJA_USD") or "").strip()

            abrv, _ = self._get_abbrev(tipo_cfe_raw, abrev_map)

            concepto = join_concept(abrv, serie, nombre_prov)

            # tasa IVA
            tasa_base = None
            if iva != 0 and neto != 0:
                tasa_base = round(abs(iva) / abs(neto), 2)

            diff = round(total - neto - iva - retper, 2)
            if abs(diff) > LIMITE_INCONGRUENTE:
                total_calc = round(neto + iva + retper, 2)
                incongruentes_detalle.append(
                    "\n".join([
                        f"RUT: {rut_dgi}",
                        f"Comprobante: {concepto}",
                        f"Número (DGI): {numero_doc}",
                        f"Fecha (día): {dia}",
                        f"Moneda (DGI): {moneda_dgi} -> {moneda_digit}",
                        f"Es NC: {'Si' if is_nc else 'No'}",
                        "",
                        f"Neto:        {d2(neto)}",
                        f"IVA:         {d2(iva)}",
                        f"Retención:   {d2(retper)}",
                        f"Total (DGI): {d2(total)}",
                        f"Total calc:  {d2(total_calc)}",
                        f"Diferencia:  {d2(diff)}",
                        "----------------------------------------",
                    ])
                )

            if tasa_base is not None and tasa_base not in (0.10, 0.22):
                iva_raro_lines.append(
                    f"RUT={rut_dgi} | Dia={dia} | Concepto='{concepto}' | Num='{numero_doc}' | tasa_base={tasa_base} "
                    f"| Neto={d2(neto)} | IVA={d2(iva)} | Total={d2(total)} | Moneda={moneda_dgi}->{moneda_digit}"
                )

            iva_fijo_raw = str(prov_row.get("IVA Fijo", "")).strip()
            iva_forzado = iva_fijo_raw in ("10", "22")

            def cta_iva():
                if iva_forzado:
                    return params["IVA10"] if iva_fijo_raw == "10" else params["IVA22"]
                if tasa_base == 0.10:
                    return params["IVA10"]
                if tasa_base == 0.22:
                    return params["IVA22"]
                return params["IVAGEN"]

            # Posteo por componentes (Factura normal o NC invertida)
            post(dia, debe_cta, "", concepto, numero_doc, rut_dgi, moneda_digit, neto, libro_char, invert=is_nc)
            post(dia, cta_iva(), "", concepto, numero_doc, rut_dgi, moneda_digit, iva, libro_char, invert=is_nc)
            post(dia, params["RETENCIONES"], "", concepto, numero_doc, rut_dgi, moneda_digit, retper, libro_char, invert=is_nc)
            post(dia, params["REDONDEOS"], "", concepto, numero_doc, rut_dgi, moneda_digit, diff, libro_char, invert=is_nc)

            # Contrapartida final (Caja si contado, Haber proveedor si crédito)
            total_contable = round(neto + iva + retper + diff, 2)
            haber_final = caja_cta if es_contado else haber_prov_cta

            # En factura: Debe vacío, Haber cuenta (crédito)
            # En NC: se invierte (debitamos proveedor/caja)
            post(dia, "", haber_final, concepto, numero_doc, rut_dgi, moneda_digit, total_contable, libro_char, invert=is_nc)

        txt_content = "\r\n".join(lines) + "\r\n"
        txt_path.write_bytes(txt_content.encode("cp1252", errors="replace"))

        (self.output_dir / "iva_raro.txt").write_text(
            "\n".join(iva_raro_lines) + ("\n" if iva_raro_lines else ""),
            encoding="utf-8"
        )
        (self.output_dir / "incongruentes.txt").write_text(
            ("\n\n".join(incongruentes_detalle).rstrip() + "\n") if incongruentes_detalle else "",
            encoding="utf-8"
        )