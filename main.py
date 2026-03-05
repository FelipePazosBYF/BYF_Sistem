# main.py
# BYFSistem – Menú + Empresas + Proveedores + Generar TXT (por empresa)
#
# Incluye mejoras PRO:
# - APP_VERSION
# - Logs automáticos
# - Panel soporte (CTRL+SHIFT+B)
# - Splash screen
# - Auto-update 1 vez por día (con confirmación)
#
# NOTA: Pandas se importa "lazy" (cuando se usa) para acelerar arranque.
#       Esto ayuda mucho en PCs lentas.

import sys
import json
import csv
import shutil
import traceback
import zipfile
import threading
import subprocess
from datetime import datetime, date
from pathlib import Path
from typing import Optional, Dict, Any, Tuple, List

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import requests

from generator import ContyGenerator

APP_NAME = "BYFSistem"
APP_VERSION = "1.0.0"

LICENSE_URL = "https://script.google.com/macros/s/AKfycbyI4EWZrWTcQx9EHWEvL1oXJ9Clx5wFTmEOoQ-SKwNifYi1MeCMxe81iYgWNGSkl3csfw/exec"

# 🔧 Auto-update: poné acá el JSON público (GitHub/Drive/CDN)
# Ejemplo JSON:
# {
#   "version": "1.0.1",
#   "url": "https://....../BYFSistem_1.0.1.zip",
#   "notes": "- Fix ...\n- Mejora ..."
# }
UPDATE_URL = "https://raw.githubusercontent.com/FelipePazosBYF/BYF_Sistem/main/update.json" 

# Proveedores base
PROV_BASE_COLS = ["RUT", "Nombre", "Debe", "IVA Fijo", "Libro", "Cont/Cred"]

# Parámetros base (sin monedas)
PARAM_BASE_COLS = ["IVA 10", "IVA 22", "IVA GEN", "REDONDEOS", "RETENCIONES"]

# Monedas default sugeridas (usuario puede modificar/eliminar extras, base no)
DEFAULT_CURRENCIES = [
    {"name": "Pesos uruguayos", "dgi": "UYU", "digit": "0", "locked": True},
    {"name": "Dólares", "dgi": "USD", "digit": "1", "locked": True},
    {"name": "Euros", "dgi": "EUR", "digit": "2", "locked": False},
    {"name": "Pesos argentinos", "dgi": "ARS", "digit": "3", "locked": False},
    {"name": "Pesos chilenos", "dgi": "CLP", "digit": "4", "locked": False},
    {"name": "Reales", "dgi": "BRL", "digit": "5", "locked": False},
]

# Tabla fija abreviaturas
DEFAULT_ABREVIATURAS_ROWS = [
    {"Tipo CFE": "e-Factura", "Abreviado": "EFAC"},
    {"Tipo CFE": "Nota de crédito de e-Factura", "Abreviado": "N/C EFAC"},
    {"Tipo CFE": "Nota de débito de e-Factura", "Abreviado": "N/D EFAC"},
    {"Tipo CFE": "e-Ticket", "Abreviado": "ETIQ"},
    {"Tipo CFE": "Nota de crédito de e-Ticket", "Abreviado": "N/C ETIK"},
    {"Tipo CFE": "Nota de débito de e-Ticket", "Abreviado": "N/D ETIK"},
    {"Tipo CFE": "e-Factura de Exportación", "Abreviado": "EFACEXP"},
    {"Tipo CFE": "Nota de crédito de e-Factura de Exportación", "Abreviado": "N/C EFEXP"},
    {"Tipo CFE": "Nota de débito de e-Factura de Exportación", "Abreviado": "N/D EFEXP"},
    {"Tipo CFE": "e-Remito de Exportación", "Abreviado": "EREMEXP"},
    {"Tipo CFE": "e-Ticket Venta por Cuenta Ajena", "Abreviado": "ETKVCTAAJ"},
    {"Tipo CFE": "Nota de crédito de e-Ticket Venta por Cuenta Ajena", "Abreviado": "N/CTKCTAJ"},
    {"Tipo CFE": "Nota de débito de e-Ticket Venta por Cuenta Ajena", "Abreviado": "N/DTKCTAJ"},
    {"Tipo CFE": "e-Factura Venta por Cuenta Ajena", "Abreviado": "EFACVTAAJ"},
    {"Tipo CFE": "Nota de crédito de e-Factura Venta por Cuenta Ajena", "Abreviado": "N/CFACVAJ"},
    {"Tipo CFE": "Nota de débito de e-Factura Venta por Cuenta Ajena", "Abreviado": "N/DFACVAJ"},
    {"Tipo CFE": "e -Remito", "Abreviado": "EREMITO"},
    {"Tipo CFE": "e-Resguardo", "Abreviado": "ERESG"},
]

MANUAL_TEXTO = """
BYFSistem — Manual rápido

- Monedas:
  Se definen en /Datos/Monedas.xlsx (código DGI + dígito TXT).
  Ej: UYU -> 0, USD -> 1, BRL -> 2.
  El TXT usa el dígito, NO el texto.

- Parámetros:
  Las cajas van por moneda: Caja UYU, Caja USD, etc.

- Proveedores:
  Los haberes van por moneda: Haber UYU, Haber USD, etc.

- Configuraciones > Proveedores (masivo):
  Permite setear Libro/Cont-Cred y Haberes por cada moneda activa.

- Soporte:
  CTRL+SHIFT+B abre panel con info de soporte.
""".strip()


# =========================
#   LOGS (PRO)
# =========================

def get_logs_dir() -> Path:
    p = Path.home() / "Documents" / "BYFSistem" / "logs"
    p.mkdir(parents=True, exist_ok=True)
    return p

def log_exception(context: str, e: Exception):
    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fp = get_logs_dir() / f"error_{ts}.log"
        fp.write_text(
            f"[{ts}] {context}\n\n{repr(e)}\n\n{traceback.format_exc()}",
            encoding="utf-8"
        )
    except Exception:
        pass


# =========================
#   RUTAS (PyInstaller safe)
# =========================

def resource_path(relative: str) -> Path:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / relative
    return Path(__file__).resolve().parent / relative

def get_icon_paths() -> Tuple[Path, Path]:
    return resource_path("byf.ico"), resource_path("Logo reducido.png")

def apply_window_icon(win: tk.Tk | tk.Toplevel) -> None:
    ico_path, png_path = get_icon_paths()
    try:
        if ico_path.exists():
            win.iconbitmap(str(ico_path))
    except Exception:
        pass
    try:
        if png_path.exists():
            img = tk.PhotoImage(file=str(png_path))
            win._byf_icon_ref = img
            win.iconphoto(True, img)
    except Exception:
        pass


# =========================
#   CONFIG (persistente)
# =========================

def get_app_dir() -> Path:
    base = Path.home() / ".byfsistem"
    base.mkdir(parents=True, exist_ok=True)
    return base

def get_config_path() -> Path:
    return get_app_dir() / "config.json"

def load_app_config() -> Dict[str, Any]:
    p = get_config_path()
    if not p.exists():
        return {"last_company_dir": "", "last_browse_dir": "", "last_update_check": ""}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if "last_update_check" not in data:
            data["last_update_check"] = ""
        return data
    except Exception:
        return {"last_company_dir": "", "last_browse_dir": "", "last_update_check": ""}

def save_app_config(cfg: Dict[str, Any]) -> None:
    get_config_path().write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# =========================
#   TK ROOT (UNO SOLO)
# =========================

def make_root() -> tk.Tk:
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    apply_window_icon(root)
    return root

def info(parent, title: str, msg: str) -> None:
    messagebox.showinfo(title, msg, parent=parent)

def error(parent, title: str, msg: str) -> None:
    messagebox.showerror(title, msg, parent=parent)

def ask_text(parent, title: str, prompt: str) -> Optional[str]:
    val = simpledialog.askstring(title, prompt, parent=parent)
    if val is None:
        return None
    val = val.strip()
    return val if val else ""

def pick_folder(parent, title: str, initialdir: Optional[str] = None) -> str:
    return filedialog.askdirectory(title=title, parent=parent, initialdir=initialdir or "")

def pick_file(parent, title: str, initialdir: Optional[str] = None) -> str:
    return filedialog.askopenfilename(
        title=title,
        parent=parent,
        initialdir=initialdir or "",
        filetypes=[("Excel", "*.xls *.xlsx"), ("Todos los archivos", "*.*")],
    )


# =========================
#   SPLASH (PRO)
# =========================

def show_splash(root: tk.Tk) -> tk.Toplevel:
    win = tk.Toplevel(root)
    apply_window_icon(win)
    win.overrideredirect(True)
    win.attributes("-topmost", True)

    w, h = 420, 160
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    frm = tk.Frame(win, padx=18, pady=18)
    frm.pack(fill="both", expand=True)

    tk.Label(frm, text=APP_NAME, font=("Segoe UI", 16, "bold")).pack(anchor="w")
    tk.Label(frm, text=f"Versión {APP_VERSION}", font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 10))
    tk.Label(frm, text="Iniciando…", font=("Segoe UI", 11)).pack(anchor="w")

    win.update_idletasks()
    return win


# =========================
#   LICENCIA
# =========================

def check_license_or_exit(root: tk.Tk) -> None:
    license_key = ask_text(root, APP_NAME, "Ingresá tu clave de licencia:")
    if license_key is None:
        sys.exit(0)

    if not license_key:
        error(root, "Licencia", "No ingresaste ninguna clave de licencia.")
        sys.exit(1)

    try:
        resp = requests.get(LICENSE_URL, params={"key": license_key}, timeout=12)
        data = resp.json()
    except Exception as e:
        error(root, "Licencia", f"No se pudo conectar al servidor.\n\nDetalle:\n{e}")
        sys.exit(1)

    if not data.get("ok"):
        error(root, "Licencia", f"Error del servidor:\n{data.get('error','UNKNOWN')}")
        sys.exit(1)

    if not data.get("valid"):
        msg = "Licencia inválida"
        if data.get("reason"):
            msg += f": {data['reason']}"
        if data.get("cliente"):
            msg += f"\nCliente: {data['cliente']}"
        if data.get("expira"):
            msg += f"\nExpira: {data['expira']}"
        error(root, "Licencia", msg)
        sys.exit(1)

    ok_msg = "Licencia válida ✅"
    if data.get("cliente"):
        ok_msg += f"\nCliente: {data['cliente']}"
    if data.get("expira"):
        ok_msg += f"\nExpira: {data['expira']}"
    info(root, "Licencia", ok_msg)


# =========================
#   EMPRESA / DATOS
# =========================

def company_data_dir(company_dir: Path) -> Path:
    return company_dir / "Datos"

def proveedores_xlsx_path(company_dir: Path) -> Path:
    return company_data_dir(company_dir) / "Proveedores.xlsx"

def parametros_xlsx_path(company_dir: Path) -> Path:
    return company_data_dir(company_dir) / "Parámetros.xlsx"

def abreviaturas_xlsx_path(company_dir: Path) -> Path:
    return company_data_dir(company_dir) / "Abreviaturas.xlsx"

def monedas_xlsx_path(company_dir: Path) -> Path:
    return company_data_dir(company_dir) / "Monedas.xlsx"

def read_parametros_df(path: Path):
    import pandas as pd
    df = pd.read_excel(path, dtype=str).fillna("")
    df.columns = df.columns.astype(str).str.strip()
    if len(df) == 0:
        df = pd.DataFrame([[""] * len(df.columns)], columns=df.columns)
    return df

def write_parametros_df(path: Path, df) -> None:
    df.to_excel(path, index=False)

def read_proveedores_df(path: Path):
    import pandas as pd
    df = pd.read_excel(path, dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    if len(df) == 0:
        df = pd.DataFrame([[""] * len(df.columns)], columns=df.columns)
    return df

def write_proveedores_df(path: Path, df) -> None:
    df.to_excel(path, index=False)

def _norm_dgi_code(s: str) -> str:
    s = "" if s is None else str(s)
    return s.strip().upper()

def read_monedas_df(path: Path):
    import pandas as pd
    df = pd.read_excel(path, dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    for c in ["DGI", "Nombre", "Digito", "Activa", "Locked"]:
        if c not in df.columns:
            df[c] = ""
    df["DGI"] = df["DGI"].apply(_norm_dgi_code)
    df["Digito"] = df["Digito"].astype(str).str.strip()
    df["Activa"] = df["Activa"].astype(str).str.strip()
    df["Locked"] = df["Locked"].astype(str).str.strip()
    df = df[df["DGI"] != ""].copy()
    return df

def write_monedas_df(path: Path, df) -> None:
    df.to_excel(path, index=False)

def get_active_currency_codes(company_dir: Path) -> List[str]:
    mp = monedas_xlsx_path(company_dir)
    dfm = read_monedas_df(mp)
    act = dfm[dfm["Activa"].str.lower().isin(["si", "sí", "s", "y", "yes", "1", "true"])]
    codes = act["DGI"].tolist()
    order = ["UYU", "USD", "EUR", "ARS", "CLP", "BRL"]
    codes = sorted(set(codes), key=lambda x: order.index(x) if x in order else 999)
    return codes

def ensure_templates(company_dir: Path) -> None:
    import pandas as pd

    ddir = company_data_dir(company_dir)
    ddir.mkdir(parents=True, exist_ok=True)

    # ABREVIATURAS: tabla fija
    ap = abreviaturas_xlsx_path(company_dir)
    needs_write = False
    if not ap.exists():
        needs_write = True
    else:
        try:
            df_ab = pd.read_excel(ap, dtype=str).fillna("")
            df_ab.columns = [str(c).strip() for c in df_ab.columns]
            if df_ab.shape[0] == 0 or "Tipo CFE" not in df_ab.columns or "Abreviado" not in df_ab.columns:
                needs_write = True
        except Exception:
            needs_write = True

    if needs_write:
        pd.DataFrame(DEFAULT_ABREVIATURAS_ROWS, columns=["Tipo CFE", "Abreviado"]).to_excel(ap, index=False)

    # MONEDAS: fuente de verdad
    mp = monedas_xlsx_path(company_dir)
    if not mp.exists():
        rows = []
        for cur in DEFAULT_CURRENCIES:
            rows.append({
                "DGI": cur["dgi"],
                "Nombre": cur["name"],
                "Digito": str(cur["digit"]),
                "Activa": "Si" if cur["locked"] else "No",
                "Locked": "Si" if cur["locked"] else "No",
            })
        pd.DataFrame(rows, columns=["DGI", "Nombre", "Digito", "Activa", "Locked"]).to_excel(mp, index=False)
    else:
        # normalizar y asegurar UYU/USD existan
        try:
            dfm = read_monedas_df(mp)
            existing = set(dfm["DGI"].tolist())
            changed = False
            for cur in DEFAULT_CURRENCIES[:2]:  # UYU/USD
                if cur["dgi"] not in existing:
                    dfm = pd.concat([dfm, pd.DataFrame([{
                        "DGI": cur["dgi"],
                        "Nombre": cur["name"],
                        "Digito": str(cur["digit"]),
                        "Activa": "Si",
                        "Locked": "Si",
                    }])], ignore_index=True)
                    changed = True
            if changed:
                write_monedas_df(mp, dfm)
        except Exception:
            pass

    # PARAMETROS: base + Caja por monedas activas
    pp = parametros_xlsx_path(company_dir)
    if not pp.exists():
        cols = PARAM_BASE_COLS.copy()
        cols += ["Caja UYU", "Caja USD"]
        pd.DataFrame([[""] * len(cols)], columns=cols).to_excel(pp, index=False)
    else:
        try:
            dfp = read_parametros_df(pp)
            changed = False
            for c in PARAM_BASE_COLS:
                if c not in dfp.columns:
                    dfp[c] = ""
                    changed = True
            active = get_active_currency_codes(company_dir)
            for code in active:
                col = f"Caja {code}"
                if col not in dfp.columns:
                    dfp[col] = ""
                    changed = True
            if changed:
                write_parametros_df(pp, dfp)
        except Exception:
            pass

    # PROVEEDORES: base + Haber por monedas activas
    pr = proveedores_xlsx_path(company_dir)
    if not pr.exists():
        cols = PROV_BASE_COLS + ["Haber UYU", "Haber USD"]
        pd.DataFrame(columns=cols).to_excel(pr, index=False)
    else:
        try:
            dfv = read_proveedores_df(pr)
            changed = False
            for c in PROV_BASE_COLS:
                if c not in dfv.columns:
                    dfv[c] = ""
                    changed = True
            active = get_active_currency_codes(company_dir)
            for code in active:
                col = f"Haber {code}"
                if col not in dfv.columns:
                    dfv[col] = ""
                    changed = True
            if changed:
                write_proveedores_df(pr, dfv)
        except Exception:
            pass


# =========================
#   PROVEEDORES: util
# =========================

def _norm_rut(r: str) -> str:
    r = "" if r is None else str(r)
    r = r.replace("\u00A0", "").replace(" ", "").strip()
    return r


# =========================
#   IMPORT Proveedores.txt
# =========================

def proveedores_txt_path(company_dir: Path) -> Path:
    return company_data_dir(company_dir) / "Proveedores.txt"

def read_text_any_encoding(p: Path) -> str:
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return p.read_text(encoding=enc, errors="strict")
        except Exception:
            continue
    return p.read_text(encoding="utf-8", errors="ignore")

def sniff_delimiter(sample_line: str) -> str:
    if sample_line.count(";") >= 1:
        return ";"
    if sample_line.count(",") >= 1:
        return ","
    if "\t" in sample_line:
        return "\t"
    return ","

def read_proveedores_txt_rut_nombre(txt_path: Path) -> Dict[str, str]:
    text = read_text_any_encoding(txt_path)
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return {}
    delim = sniff_delimiter(lines[0])
    out: Dict[str, str] = {}
    reader = csv.reader(lines, delimiter=delim, quotechar='"', skipinitialspace=True)
    for row in reader:
        if not row or len(row) < 2:
            continue
        rut = str(row[0]).strip().strip('"').strip()
        nombre = str(row[1]).strip().strip('"').strip()
        if rut and nombre:
            out[rut] = nombre
    return out

def import_proveedores_txt_if_any(company_dir: Path) -> Tuple[int, Optional[str]]:
    txt = proveedores_txt_path(company_dir)
    if not txt.exists():
        return 0, None

    ensure_templates(company_dir)
    xlsx = proveedores_xlsx_path(company_dir)

    df = read_proveedores_df(xlsx).fillna("")
    if "RUT" not in df.columns or "Nombre" not in df.columns:
        import pandas as pd
        df = pd.DataFrame(columns=PROV_BASE_COLS + ["Haber UYU", "Haber USD"])

    existing = set(df["RUT"].astype(str).map(_norm_rut).tolist()) if "RUT" in df.columns else set()

    nuevos = read_proveedores_txt_rut_nombre(txt)
    added_rows = []
    added = 0
    for rut, nombre in nuevos.items():
        rut_n = _norm_rut(rut)
        if not rut_n or rut_n in existing:
            continue
        row = {c: "" for c in df.columns}
        for c in PROV_BASE_COLS:
            if c not in row:
                row[c] = ""
        row["RUT"] = rut_n
        row["Nombre"] = nombre.strip()
        row["Libro"] = "C"
        row["Cont/Cred"] = "Crédito"
        added_rows.append(row)
        existing.add(rut_n)
        added += 1

    if added_rows:
        import pandas as pd
        df2 = pd.concat([df, pd.DataFrame(added_rows)], ignore_index=True)
        write_proveedores_df(xlsx, df2)

    try:
        txt.unlink()
    except Exception:
        try:
            backup = txt.with_name("Proveedores_importado.txt")
            try:
                if backup.exists():
                    backup.unlink()
            except Exception:
                pass
            shutil.move(str(txt), str(backup))
        except Exception:
            pass

    return added, txt.name


# =========================
#   UI STATE
# =========================

class AppState:
    def __init__(self):
        self.cfg = load_app_config()
        self.company_dir: Optional[Path] = None
        last = (self.cfg.get("last_company_dir") or "").strip()
        if last:
            p = Path(last)
            if p.exists() and p.is_dir():
                self.company_dir = p

    def set_company(self, company_dir: Path):
        self.company_dir = company_dir
        self.cfg["last_company_dir"] = str(company_dir)
        self.cfg["last_browse_dir"] = str(company_dir)
        save_app_config(self.cfg)


# =========================
#   Soporte PRO: update
# =========================

def _parse_version(v: str) -> Tuple[int, int, int]:
    v = (v or "").strip()
    parts = v.split(".")
    nums = []
    for i in range(3):
        try:
            nums.append(int(parts[i]) if i < len(parts) else 0)
        except Exception:
            nums.append(0)
    return (nums[0], nums[1], nums[2])

def _is_newer(remote: str, local: str) -> bool:
    return _parse_version(remote) > _parse_version(local)

def _today_str() -> str:
    return date.today().isoformat()

def _app_folder() -> Path:
    # En onedir, sys.executable apunta al exe dentro de la carpeta dist\BYFSistem\
    return Path(sys.executable).resolve().parent

def _updates_dir() -> Path:
    p = Path.home() / "Documents" / "BYFSistem" / "updates"
    p.mkdir(parents=True, exist_ok=True)
    return p

def check_updates_daily(parent, state: AppState):
    # 1 vez por día
    try:
        if not UPDATE_URL:
            return
        last = (state.cfg.get("last_update_check") or "").strip()
        if last == _today_str():
            return
        state.cfg["last_update_check"] = _today_str()
        save_app_config(state.cfg)

        # thread para no congelar UI
        def worker():
            try:
                r = requests.get(UPDATE_URL, timeout=12)
                data = r.json()
                remote_v = str(data.get("version", "")).strip()
                url = str(data.get("url", "")).strip()
                notes = str(data.get("notes", "")).strip()

                if not remote_v or not url:
                    return
                if not _is_newer(remote_v, APP_VERSION):
                    return

                def ask_on_ui():
                    msg = f"Hay una nueva versión disponible.\n\n" \
                          f"Actual: {APP_VERSION}\nNueva: {remote_v}\n\n"
                    if notes:
                        msg += f"Cambios:\n{notes}\n\n"
                    msg += "¿Querés actualizar ahora?"
                    if messagebox.askyesno("Actualización", msg, parent=parent):
                        run_update_flow(parent, remote_v, url)
                parent.after(0, ask_on_ui)

            except Exception as e:
                log_exception("Update check", e)

        threading.Thread(target=worker, daemon=True).start()

    except Exception as e:
        log_exception("check_updates_daily", e)

def run_update_flow(parent, remote_version: str, zip_url: str):
    """
    Descarga ZIP (onedir) y reemplaza carpeta actual con un .bat.
    """
    try:
        upd_dir = _updates_dir()
        zip_path = upd_dir / f"BYFSistem_{remote_version}.zip"
        extract_dir = upd_dir / f"BYFSistem_{remote_version}"

        # limpiar si existe
        if extract_dir.exists():
            shutil.rmtree(extract_dir, ignore_errors=True)

        # descargar
        info(parent, "Actualización", "Descargando actualización…")
        with requests.get(zip_url, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(zip_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 256):
                    if chunk:
                        f.write(chunk)

        # extraer
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(extract_dir)

        # detectar carpeta raíz del zip:
        # esperamos que el zip contenga una carpeta "BYFSistem" o directamente los archivos.
        new_root = extract_dir
        candidates = [p for p in extract_dir.iterdir() if p.is_dir()]
        if len(candidates) == 1 and (candidates[0] / "BYFSistem.exe").exists():
            new_root = candidates[0]
        elif (extract_dir / "BYFSistem.exe").exists():
            new_root = extract_dir
        else:
            # buscar un exe dentro
            found = list(extract_dir.rglob("BYFSistem.exe"))
            if found:
                new_root = found[0].parent
            else:
                raise RuntimeError("El ZIP no contiene BYFSistem.exe en una estructura válida.")

        cur_dir = _app_folder()
        cur_exe = cur_dir / "BYFSistem.exe"

        # BAT de reemplazo
        bat = upd_dir / "update_byfsistem.bat"
        backup_dir = cur_dir.with_name(cur_dir.name + "_backup")

        # Usamos robocopy para copiar reemplazando (más robusto en Windows)
        # 1) esperar 1s
        # 2) matar proceso
        # 3) renombrar carpeta actual a backup
        # 4) crear carpeta nueva y copiar
        # 5) abrir exe nuevo
        bat_content = f"""@echo off
setlocal
timeout /t 1 /nobreak >nul

REM Cerrar el programa (por si sigue abierto)
taskkill /IM BYFSistem.exe /F >nul 2>nul

REM Esperar un poco
timeout /t 1 /nobreak >nul

REM Borrar backup anterior
if exist "{backup_dir}" rmdir /s /q "{backup_dir}"

REM Renombrar carpeta actual a backup
cd /d "{cur_dir.parent}"
rename "{cur_dir.name}" "{backup_dir.name}"

REM Crear carpeta nueva
mkdir "{cur_dir}"

REM Copiar archivos nuevos
robocopy "{new_root}" "{cur_dir}" /E /NFL /NDL /NJH /NJS /nc /ns /np >nul

REM Lanzar nueva versión
start "" "{cur_exe}"

endlocal
"""
        bat.write_text(bat_content, encoding="utf-8")

        info(parent, "Actualización", "Se va a cerrar y actualizar el programa.\n\nDale Aceptar para continuar.")
        # ejecutar bat y cerrar app
        subprocess.Popen(["cmd", "/c", str(bat)], creationflags=0x08000000)  # CREATE_NO_WINDOW
        parent.winfo_toplevel().destroy()

    except Exception as e:
        log_exception("run_update_flow", e)
        error(parent, "Actualización", f"No se pudo actualizar:\n\n{e}")


# =========================
#   Missing providers + helpers
# =========================

def _ensure_haber_cols_for_active(company_dir: Path) -> None:
    prov_path = proveedores_xlsx_path(company_dir)
    dfv = read_proveedores_df(prov_path).fillna("")
    active = get_active_currency_codes(company_dir)
    changed = False
    for base in PROV_BASE_COLS:
        if base not in dfv.columns:
            dfv[base] = ""
            changed = True
    for code in active:
        col = f"Haber {code}"
        if col not in dfv.columns:
            dfv[col] = ""
            changed = True
    if changed:
        write_proveedores_df(prov_path, dfv)

def add_missing_providers_to_xlsx(company_dir: Path, missing_ruts: List[str]) -> int:
    import pandas as pd

    ensure_templates(company_dir)
    _ensure_haber_cols_for_active(company_dir)

    prov_path = proveedores_xlsx_path(company_dir)
    df = read_proveedores_df(prov_path).fillna("")
    existing = set(df["RUT"].astype(str).map(_norm_rut).tolist()) if "RUT" in df.columns else set()

    added_rows = []
    added = 0
    for rut in missing_ruts:
        rut_n = _norm_rut(rut)
        if not rut_n or rut_n in existing:
            continue
        row = {c: "" for c in df.columns}
        if "RUT" in row:
            row["RUT"] = rut_n
        if "Nombre" in row:
            row["Nombre"] = ""
        if "Libro" in row:
            row["Libro"] = "C"
        if "Cont/Cred" in row:
            row["Cont/Cred"] = "Crédito"
        if "IVA Fijo" in row:
            row["IVA Fijo"] = ""
        added_rows.append(row)
        existing.add(rut_n)
        added += 1

    if added_rows:
        df2 = pd.concat([df, pd.DataFrame(added_rows)], ignore_index=True)
        write_proveedores_df(prov_path, df2)
    return added


class MissingProvidersWindow(tk.Toplevel):
    def __init__(self, parent, state: AppState, missing_ruts: List[str]):
        super().__init__(parent)
        apply_window_icon(self)
        self.parent = parent
        self.state = state
        self.missing_ruts = sorted(set(missing_ruts))
        self.action = "cancel"
        self.title("Faltan proveedores")
        self.geometry("640x460")
        self.resizable(False, False)

        tk.Label(self, text="Faltan proveedores en Proveedores.xlsx", font=("Segoe UI", 13, "bold"))\
            .pack(anchor="w", padx=14, pady=(12, 4))

        tk.Label(
            self,
            text="Estos RUTs aparecen en el Excel DGI pero no están en la planilla de proveedores.\n"
                 "Podés agregarlos automáticamente y completarlos, o continuar igual.",
            justify="left"
        ).pack(anchor="w", padx=14, pady=(0, 10))

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        self.listbox = tk.Listbox(frame, height=14, font=("Segoe UI", 10))
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=sb.set)

        for rut in self.missing_ruts:
            self.listbox.insert(tk.END, rut)

        bottom = tk.Frame(self)
        bottom.pack(fill="x", padx=14, pady=(0, 14))

        tk.Button(bottom, text="Cancelar", height=2, command=self.on_cancel).pack(side="left")
        tk.Button(bottom, text="Continuar de todos modos", height=2, command=self.on_continue)\
            .pack(side="right")
        tk.Button(bottom, text="Agregar proveedores", height=2, command=self.on_add)\
            .pack(side="right", padx=10)

        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def on_cancel(self):
        self.action = "cancel"
        self.destroy()

    def on_continue(self):
        self.action = "continue"
        self.destroy()

    def on_add(self):
        if not self.state.company_dir:
            error(self, "Proveedores", "Primero seleccione la empresa.")
            return
        try:
            added = add_missing_providers_to_xlsx(self.state.company_dir, self.missing_ruts)
            first = _norm_rut(self.missing_ruts[0]) if self.missing_ruts else None
            info(
                self,
                "Proveedores",
                f"Se agregaron {added} proveedores nuevos.\n\n"
                f"Ahora se abrirá Proveedores en el primer agregado."
            )
            ProveedoresWindow(self.parent, self.state, start_rut=first)
            self.action = "add"
            self.destroy()
        except Exception as e:
            log_exception("Agregar proveedores faltantes", e)
            error(self, "Proveedores", f"No se pudieron agregar proveedores:\n\n{e}")


# =========================
#   Config Hub / Bulk / Param / Monedas
# =========================

class ConfigHubWindow(tk.Toplevel):
    def __init__(self, parent, state: AppState):
        super().__init__(parent)
        apply_window_icon(self)
        self.parent = parent
        self.state = state

        self.title("Configuraciones")
        self.geometry("520x260")
        self.resizable(False, False)

        if not self.state.company_dir:
            info(self, "Configuraciones", "Primero seleccione la empresa.")
            self.destroy()
            return

        ensure_templates(self.state.company_dir)

        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(frm, text="Configuraciones", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 10))
        tk.Button(frm, text="Proveedores (masivo)", height=2, command=self.open_bulk).pack(fill="x", pady=6)
        tk.Button(frm, text="Parámetros", height=2, command=self.open_parametros).pack(fill="x", pady=6)
        tk.Button(frm, text="Monedas", height=2, command=self.open_monedas).pack(fill="x", pady=6)
        tk.Button(frm, text="Cerrar", height=2, command=self.destroy).pack(fill="x", pady=(12, 0))

    def open_bulk(self):
        BulkProveedorConfigWindow(self, self.state)

    def open_parametros(self):
        ParametrosWindow(self, self.state)

    def open_monedas(self):
        MonedasWindow(self, self.state)


class BulkProveedorConfigWindow(tk.Toplevel):
    def __init__(self, parent, state: AppState):
        super().__init__(parent)
        apply_window_icon(self)
        self.parent = parent
        self.state = state

        self.title("Configuraciones — Proveedores (masivo)")
        self.geometry("720x520")
        self.minsize(720, 520)

        if not self.state.company_dir:
            info(self, "Configuraciones", "Primero seleccione la empresa.")
            self.destroy()
            return

        ensure_templates(self.state.company_dir)
        self.codes = get_active_currency_codes(self.state.company_dir)
        _ensure_haber_cols_for_active(self.state.company_dir)

        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(frm, text="Aplicar a TODOS los proveedores:", font=("Segoe UI", 11, "bold"))\
            .grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        tk.Label(frm, text="Libro").grid(row=1, column=0, sticky="e", padx=8, pady=8)
        self.libro_var = tk.StringVar(value="C")
        ttk.Combobox(frm, textvariable=self.libro_var, values=["C", "E"], width=10, state="readonly")\
            .grid(row=1, column=1, sticky="w", padx=8, pady=8)

        tk.Label(frm, text="Cont/Cred").grid(row=2, column=0, sticky="e", padx=8, pady=8)
        self.contcred_var = tk.StringVar(value="Crédito")
        ttk.Combobox(frm, textvariable=self.contcred_var, values=["Crédito", "Contado"], width=12, state="readonly")\
            .grid(row=2, column=1, sticky="w", padx=8, pady=8)

        sep = ttk.Separator(frm, orient="horizontal")
        sep.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(14, 12))

        tk.Label(frm, text="Cuentas Haber (por moneda DGI activa)", font=("Segoe UI", 10, "bold"))\
            .grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 8))

        self.haber_vars: Dict[str, tk.StringVar] = {}
        row = 5
        for code in self.codes:
            col = f"Haber {code}"
            tk.Label(frm, text=col).grid(row=row, column=0, sticky="e", padx=8, pady=6)
            v = tk.StringVar(value="")
            self.haber_vars[code] = v
            tk.Entry(frm, textvariable=v, width=32).grid(row=row, column=1, sticky="w", padx=8, pady=6)
            row += 1

        tk.Label(
            frm,
            text="Nota: Esto pisa Libro, Cont/Cred y los Haberes en todas las filas con RUT/Nombre.",
            justify="left"
        ).grid(row=row, column=0, columnspan=3, sticky="w", padx=2, pady=(12, 0))

        btns = tk.Frame(self)
        btns.pack(fill="x", padx=16, pady=(0, 16))

        tk.Button(btns, text="Aplicar a todos", height=2, command=self.apply_all).pack(side="right")
        tk.Button(btns, text="Cerrar", height=2, command=self.destroy).pack(side="right", padx=10)

    def apply_all(self):
        try:
            path = proveedores_xlsx_path(self.state.company_dir)
            df = read_proveedores_df(path).fillna("")
            if "RUT" not in df.columns or "Nombre" not in df.columns:
                raise ValueError("Proveedores.xlsx no tiene columnas RUT/Nombre.")

            for code in self.codes:
                c = f"Haber {code}"
                if c not in df.columns:
                    df[c] = ""

            mask = (df["RUT"].astype(str).str.strip() != "") | (df["Nombre"].astype(str).str.strip() != "")
            df.loc[mask, "Libro"] = self.libro_var.get().strip()
            df.loc[mask, "Cont/Cred"] = self.contcred_var.get().strip()

            for code, var in self.haber_vars.items():
                df.loc[mask, f"Haber {code}"] = var.get().strip()

            write_proveedores_df(path, df)
            info(self, "Configuraciones", "Cambios aplicados ✅")
        except Exception as e:
            log_exception("BulkProveedorConfigWindow.apply_all", e)
            error(self, "Configuraciones", f"No se pudieron aplicar los cambios:\n\n{e}")


class ParametrosWindow(tk.Toplevel):
    def __init__(self, parent, state: AppState):
        super().__init__(parent)
        apply_window_icon(self)
        self.parent = parent
        self.state = state

        self.title("Configuraciones — Parámetros")
        self.geometry("760x520")
        self.minsize(760, 520)

        if not self.state.company_dir:
            info(self, "Parámetros", "Primero seleccione la empresa.")
            self.destroy()
            return

        ensure_templates(self.state.company_dir)
        self.path = parametros_xlsx_path(self.state.company_dir)
        self.df = read_parametros_df(self.path)

        tk.Label(self, text=f"Editando: {self.path}", font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(10, 0))
        tk.Label(self, text="Parámetros", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=14, pady=(6, 10))

        container = tk.Frame(self)
        container.pack(fill="both", expand=True, padx=14, pady=10)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.inner = tk.Frame(canvas)

        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.vars: Dict[str, tk.StringVar] = {}

        base_cols = [c for c in PARAM_BASE_COLS if c in self.df.columns]
        caja_cols = [c for c in self.df.columns if str(c).strip().upper().startswith("CAJA ")]
        active = get_active_currency_codes(self.state.company_dir)
        caja_order = [f"Caja {c}" for c in active]
        caja_cols = sorted(set(caja_cols), key=lambda x: caja_order.index(x) if x in caja_order else 999)
        ordered = base_cols + caja_cols + [c for c in self.df.columns if c not in base_cols + caja_cols]

        row = 0
        for col in ordered:
            tk.Label(self.inner, text=col, width=18, anchor="e").grid(row=row, column=0, sticky="e", padx=8, pady=6)
            v = tk.StringVar(value=str(self.df.iloc[0].get(col, "")).strip())
            self.vars[col] = v
            tk.Entry(self.inner, textvariable=v, width=45).grid(row=row, column=1, sticky="w", padx=8, pady=6)
            row += 1

        btns = tk.Frame(self)
        btns.pack(fill="x", padx=14, pady=(0, 14))

        tk.Button(btns, text="Guardar", height=2, command=self.save).pack(side="right")
        tk.Button(btns, text="Cerrar", height=2, command=self.destroy).pack(side="right", padx=10)

    def save(self):
        try:
            import pandas as pd
            if len(self.df) == 0:
                self.df = pd.DataFrame([[""] * len(self.df.columns)], columns=self.df.columns)

            # asegurar texto SIEMPRE
            self.df = self.df.astype(str).fillna("")

            for col, var in self.vars.items():
                if col not in self.df.columns:
                    self.df[col] = ""
                self.df.at[0, col] = (var.get() or "").strip()

            out = self.df.copy()
            out.columns = out.columns.astype(str).str.strip()
            for c in out.columns:
                out[c] = out[c].astype(str).fillna("").map(lambda x: x.strip())

            out.to_excel(self.path, index=False)
            info(self, "Parámetros", "Guardado ✅")
        except Exception as e:
            log_exception("ParametrosWindow.save", e)
            error(self, "Parámetros", f"No se pudo guardar:\n\n{e}")


class MonedasWindow(tk.Toplevel):
    def __init__(self, parent, state: AppState):
        super().__init__(parent)
        apply_window_icon(self)
        self.parent = parent
        self.state = state

        self.title("Configuraciones — Monedas")
        self.geometry("860x650")
        self.minsize(860, 650)
        self.resizable(True, True)

        if not self.state.company_dir:
            info(self, "Monedas", "Primero seleccione la empresa.")
            self.destroy()
            return

        ensure_templates(self.state.company_dir)
        self.monedas_path = monedas_xlsx_path(self.state.company_dir)
        self.param_path = parametros_xlsx_path(self.state.company_dir)
        self.prov_path = proveedores_xlsx_path(self.state.company_dir)

        top = tk.Frame(self)
        top.pack(fill="x", padx=14, pady=(12, 6))

        tk.Label(top, text="Monedas (código DGI + dígito TXT)", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(
            top,
            text="Regla: el dígito (0–9) es lo que va en el TXT en la columna Moneda.\n"
                 "Ej: UYU -> 0, USD -> 1, BRL -> 2.",
            justify="left"
        ).pack(anchor="w", pady=(6, 0))

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        cols = ("DGI", "Nombre", "Digito", "Activa", "Locked")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=11)
        for c in cols:
            self.tree.heading(c, text=c)
            w = 110 if c in ("DGI", "Digito", "Activa", "Locked") else 330
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        controls = tk.LabelFrame(self, text="Acciones", padx=10, pady=10)
        controls.pack(fill="x", padx=14, pady=(0, 12))

        row1 = tk.Frame(controls)
        row1.pack(fill="x", pady=6)
        tk.Label(row1, text="Agregar moneda (código DGI):").pack(side="left")

        self.add_dgi = tk.StringVar()
        tk.Entry(row1, textvariable=self.add_dgi, width=12).pack(side="left", padx=8)

        tk.Label(row1, text="Nombre:").pack(side="left", padx=(10, 0))
        self.add_name = tk.StringVar()
        tk.Entry(row1, textvariable=self.add_name, width=30).pack(side="left", padx=8)

        tk.Label(row1, text="Dígito TXT:").pack(side="left", padx=(10, 0))
        self.add_digit = tk.StringVar()
        tk.Entry(row1, textvariable=self.add_digit, width=5).pack(side="left", padx=8)

        tk.Button(row1, text="Agregar", height=1, command=self.add_currency).pack(side="left", padx=10)

        row2 = tk.Frame(controls)
        row2.pack(fill="x", pady=6)

        tk.Button(row2, text="Toggle Activa", command=self.toggle_active).pack(side="left")
        tk.Button(row2, text="Editar dígito TXT", command=self.edit_digit).pack(side="left", padx=8)
        tk.Button(row2, text="Quitar (solo no locked)", command=self.remove_selected).pack(side="left", padx=8)

        tk.Button(controls, text="Guardar y sincronizar columnas", height=2, command=self.save_and_sync).pack(side="right")

        self.refresh()

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        df = read_monedas_df(self.monedas_path)
        df = df.sort_values(by=["Locked", "DGI"], ascending=[False, True])
        for _, r in df.iterrows():
            self.tree.insert("", "end", values=(r["DGI"], r["Nombre"], r["Digito"], r["Activa"], r["Locked"]))

    def _selected_dgi(self) -> Optional[str]:
        sel = self.tree.selection()
        if not sel:
            return None
        values = self.tree.item(sel[0], "values")
        return str(values[0]).strip().upper() if values else None

    def _validate_digit(self, d: str) -> bool:
        d = (d or "").strip()
        return len(d) == 1 and d.isdigit()

    def add_currency(self):
        import pandas as pd

        dgi = _norm_dgi_code(self.add_dgi.get())
        name = (self.add_name.get() or "").strip()
        digit = (self.add_digit.get() or "").strip()

        if not dgi:
            info(self, "Monedas", "Poné el código DGI (ej: BRL).")
            return
        if not name:
            info(self, "Monedas", "Poné un nombre (ej: Reales).")
            return
        if not self._validate_digit(digit):
            info(self, "Monedas", "El dígito TXT debe ser un solo número (0–9).")
            return

        df = read_monedas_df(self.monedas_path)
        if dgi in set(df["DGI"].tolist()):
            info(self, "Monedas", "Esa moneda ya existe. Usá editar/toggle.")
            return

        used = set(df["Digito"].astype(str).str.strip().tolist())
        if digit in used:
            info(self, "Monedas", f"El dígito {digit} ya está usado por otra moneda. Elegí otro.")
            return

        df = pd.concat([df, pd.DataFrame([{
            "DGI": dgi,
            "Nombre": name,
            "Digito": digit,
            "Activa": "Si",
            "Locked": "No",
        }])], ignore_index=True)
        write_monedas_df(self.monedas_path, df)
        self.add_dgi.set(""); self.add_name.set(""); self.add_digit.set("")
        self.refresh()

    def toggle_active(self):
        dgi = self._selected_dgi()
        if not dgi:
            info(self, "Monedas", "Seleccioná una moneda.")
            return
        df = read_monedas_df(self.monedas_path)
        idx = df.index[df["DGI"] == dgi]
        if len(idx) == 0:
            return
        i = idx[0]
        locked = str(df.at[i, "Locked"]).strip().lower() in ("si", "sí", "1", "true", "y", "yes")
        if locked:
            info(self, "Monedas", "UYU/USD son base (locked) y no se desactivan.")
            return
        cur = str(df.at[i, "Activa"]).strip().lower()
        df.at[i, "Activa"] = "No" if cur in ("si", "sí", "1", "true", "y", "yes") else "Si"
        write_monedas_df(self.monedas_path, df)
        self.refresh()

    def edit_digit(self):
        dgi = self._selected_dgi()
        if not dgi:
            info(self, "Monedas", "Seleccioná una moneda.")
            return
        new_d = ask_text(self, "Editar dígito TXT", f"Nuevo dígito TXT (0–9) para {dgi}:")
        if new_d is None:
            return
        new_d = new_d.strip()
        if not self._validate_digit(new_d):
            info(self, "Monedas", "El dígito TXT debe ser un solo número (0–9).")
            return
        df = read_monedas_df(self.monedas_path)
        used = set(df[df["DGI"] != dgi]["Digito"].astype(str).str.strip().tolist())
        if new_d in used:
            info(self, "Monedas", f"El dígito {new_d} ya está usado por otra moneda.")
            return
        idx = df.index[df["DGI"] == dgi]
        if len(idx) == 0:
            return
        df.at[idx[0], "Digito"] = new_d
        write_monedas_df(self.monedas_path, df)
        self.refresh()

    def remove_selected(self):
        dgi = self._selected_dgi()
        if not dgi:
            info(self, "Monedas", "Seleccioná una moneda.")
            return
        df = read_monedas_df(self.monedas_path)
        idx = df.index[df["DGI"] == dgi]
        if len(idx) == 0:
            return
        i = idx[0]
        locked = str(df.at[i, "Locked"]).strip().lower() in ("si", "sí", "1", "true", "y", "yes")
        if locked:
            info(self, "Monedas", "UYU/USD son base y no se pueden quitar.")
            return
        if not messagebox.askyesno(
            "Quitar moneda",
            f"¿Seguro que querés quitar {dgi}?\n\nSe eliminarán columnas:\n- Parámetros: Caja {dgi}\n- Proveedores: Haber {dgi}",
            parent=self
        ):
            return

        df = df[df["DGI"] != dgi].copy()
        write_monedas_df(self.monedas_path, df)

        try:
            dfp = read_parametros_df(self.param_path)
            colp = f"Caja {dgi}"
            if colp in dfp.columns:
                dfp = dfp.drop(columns=[colp])
                write_parametros_df(self.param_path, dfp)
        except Exception:
            pass
        try:
            dfv = read_proveedores_df(self.prov_path)
            colv = f"Haber {dgi}"
            if colv in dfv.columns:
                dfv = dfv.drop(columns=[colv])
                write_proveedores_df(self.prov_path, dfv)
        except Exception:
            pass

        self.refresh()

    def save_and_sync(self):
        try:
            dfm = read_monedas_df(self.monedas_path)

            digits = dfm["Digito"].astype(str).str.strip().tolist()
            for d in digits:
                if not (len(d) == 1 and d.isdigit()):
                    raise ValueError("Hay monedas con dígito TXT inválido. Debe ser 1 dígito (0–9).")
            if len(set(digits)) != len(digits):
                raise ValueError("Hay dígitos TXT repetidos. Deben ser únicos por moneda.")

            active = dfm[dfm["Activa"].str.lower().isin(["si", "sí", "s", "y", "yes", "1", "true"])]["DGI"].tolist()
            if "UYU" not in active or "USD" not in active:
                raise ValueError("UYU y USD deben estar activas.")

            dfp = read_parametros_df(self.param_path)
            changed_p = False
            for base in PARAM_BASE_COLS:
                if base not in dfp.columns:
                    dfp[base] = ""
                    changed_p = True
            for c in active:
                col = f"Caja {c}"
                if col not in dfp.columns:
                    dfp[col] = ""
                    changed_p = True
            if changed_p:
                write_parametros_df(self.param_path, dfp)

            dfv = read_proveedores_df(self.prov_path).fillna("")
            changed_v = False
            for base in PROV_BASE_COLS:
                if base not in dfv.columns:
                    dfv[base] = ""
                    changed_v = True
            for c in active:
                col = f"Haber {c}"
                if col not in dfv.columns:
                    dfv[col] = ""
                    changed_v = True
            if changed_v:
                write_proveedores_df(self.prov_path, dfv)

            info(self, "Monedas", "Guardado y sincronizado ✅\n\nSe aseguraron columnas Caja/Haber para monedas activas.")
        except Exception as e:
            log_exception("MonedasWindow.save_and_sync", e)
            error(self, "Monedas", f"No se pudo guardar/sincronizar:\n\n{e}")


# =========================
#   Menú principal
# =========================

class MainMenu(tk.Toplevel):
    def __init__(self, root: tk.Tk, state: AppState):
        super().__init__(root)
        apply_window_icon(self)

        self.root = root
        self.state = state
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("520x380")
        self.resizable(False, False)

        # Panel soporte (PRO)
        self.bind_all("<Control-Shift-b>", lambda e: self.open_support_panel())

        menubar = tk.Menu(self)
        self.config(menu=menubar)

        m_ayuda = tk.Menu(menubar, tearoff=0)
        m_ayuda.add_command(label="Manual de usuario", command=self.open_manual)
        menubar.add_cascade(label="Ayuda", menu=m_ayuda)

        m_cfg = tk.Menu(menubar, tearoff=0)
        m_cfg.add_command(label="Configuraciones", command=self.open_config)
        menubar.add_cascade(label="Configuraciones", menu=m_cfg)

        self.label_company = tk.Label(self, text=self._company_label(), font=("Segoe UI", 10, "bold"))
        self.label_company.pack(pady=(14, 10))

        frm = tk.Frame(self)
        frm.pack(pady=10, padx=18, fill="x")

        tk.Button(frm, text="Empresas", height=2, command=self.open_empresas).pack(fill="x", pady=6)
        tk.Button(frm, text="Proveedores", height=2, command=self.open_proveedores).pack(fill="x", pady=6)
        tk.Button(frm, text="Generar TXT", height=2, command=self.run_generate).pack(fill="x", pady=6)

        bottom = tk.Frame(self)
        bottom.pack(fill="x", padx=18, pady=(8, 0))
        tk.Button(bottom, text="Salir", height=2, command=self.on_exit).pack(fill="x")

        self.after(100, self._bootstrap_company)

        # Auto-update diario (PRO)
        self.after(600, lambda: check_updates_daily(self, self.state))

    def open_support_panel(self):
        win = tk.Toplevel(self)
        apply_window_icon(win)
        win.title("Soporte - BYFSistem")
        win.geometry("740x460")

        info_txt = []
        info_txt.append(f"App: {APP_NAME}")
        info_txt.append(f"Versión: {APP_VERSION}")
        info_txt.append(f"Empresa: {self.state.company_dir or '(ninguna)'}")
        info_txt.append(f"Config: {get_config_path()}")
        info_txt.append(f"Logs: {get_logs_dir()}")
        info_txt.append(f"Executable: {sys.executable}")
        info_txt.append(f"Update URL: {UPDATE_URL or '(vacío)'}")

        txt = tk.Text(win, wrap="word", font=("Consolas", 10))
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", "\n".join(map(str, info_txt)))
        txt.configure(state="disabled")

        def copy():
            win.clipboard_clear()
            win.clipboard_append("\n".join(map(str, info_txt)))
            messagebox.showinfo("Soporte", "Info copiada ✅", parent=win)

        btns = tk.Frame(win)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(btns, text="Copiar info", height=2, command=copy).pack(side="right")
        tk.Button(btns, text="Cerrar", height=2, command=win.destroy).pack(side="right", padx=10)

    def open_manual(self):
        win = tk.Toplevel(self)
        apply_window_icon(win)
        win.title("Ayuda - Manual de usuario")
        win.geometry("820x600")
        win.minsize(820, 600)

        tk.Label(win, text="Manual de usuario", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=14, pady=(12, 6))

        frm = tk.Frame(win)
        frm.pack(fill="both", expand=True, padx=14, pady=12)

        scroll = tk.Scrollbar(frm, orient="vertical")
        scroll.pack(side="right", fill="y")

        txt = tk.Text(frm, wrap="word", yscrollcommand=scroll.set, font=("Segoe UI", 10))
        txt.pack(fill="both", expand=True)
        scroll.config(command=txt.yview)

        txt.insert("1.0", MANUAL_TEXTO)
        txt.configure(state="disabled")

        tk.Button(win, text="Cerrar", height=2, command=win.destroy).pack(pady=(0, 14))

    def open_config(self):
        ConfigHubWindow(self, self.state)

    def _company_label(self) -> str:
        if self.state.company_dir:
            return f"Empresa seleccionada: {self.state.company_dir}"
        return "Empresa seleccionada: (ninguna)"

    def _refresh_company_label(self):
        self.label_company.config(text=self._company_label())

    def _bootstrap_company(self):
        if not self.state.company_dir:
            return
        try:
            ensure_templates(self.state.company_dir)
            added, fname = import_proveedores_txt_if_any(self.state.company_dir)
            if fname:
                info(self, "Proveedores", f"Se importó '{fname}' y se agregaron {added} proveedores nuevos.")
        except Exception as e:
            log_exception("bootstrap_company", e)
            error(self, "Empresa", f"No se pudo preparar la carpeta de la empresa:\n\n{e}")

    def open_empresas(self):
        EmpresasWindow(self, self.state, on_changed=self._refresh_company_label)

    def open_proveedores(self):
        if not self.state.company_dir:
            info(self, "Proveedores", "Primero seleccione la empresa.")
            return
        try:
            ensure_templates(self.state.company_dir)
            added, fname = import_proveedores_txt_if_any(self.state.company_dir)
            if fname:
                info(self, "Proveedores", f"Se importó '{fname}' y se agregaron {added} proveedores nuevos.")
            ProveedoresWindow(self, self.state)
        except Exception as e:
            log_exception("open_proveedores", e)
            error(self, "Proveedores", f"No se pudo abrir Proveedores:\n\n{e}")

    def run_generate(self):
        if not self.state.company_dir:
            info(self, "Generar", "Primero seleccione la empresa.")
            return

        initial = self.state.cfg.get("last_browse_dir") or str(self.state.company_dir)
        dgi_path = pick_file(self, "Seleccioná el archivo DGI (CFE)", initialdir=initial)
        if not dgi_path:
            return
        self.state.cfg["last_browse_dir"] = str(Path(dgi_path).parent)
        save_app_config(self.state.cfg)

        try:
            ensure_templates(self.state.company_dir)
            added, fname = import_proveedores_txt_if_any(self.state.company_dir)
            if fname:
                info(self, "Proveedores", f"Se importó '{fname}' y se agregaron {added} proveedores nuevos.")

            gen = ContyGenerator(company_dir=self.state.company_dir, dgi_xls=dgi_path)

            missing_cur = gen.precheck_missing_currencies()
            if missing_cur:
                info(
                    self,
                    "Monedas faltantes",
                    "El Excel DGI contiene monedas que no están configuradas en /Datos/Monedas.xlsx.\n\n"
                    "Abrí Configuraciones > Monedas y agregalas con su dígito TXT.\n\n"
                    "Faltan:\n- " + "\n- ".join(missing_cur)
                )
                return

            missing = gen.precheck_missing_ruts(write_file=True, clear_output=True)
            if missing:
                dlg = MissingProvidersWindow(self, self.state, missing)
                self.wait_window(dlg)

                if dlg.action == "cancel":
                    return
                if dlg.action == "add":
                    info(self, "Generar TXT", "Cuando termines de completar los proveedores, volvé a 'Generar TXT' ✅")
                    return
                if dlg.action == "continue":
                    txt_path = gen.run(allow_missing=True, skip_prepare=True)
                    info(
                        self,
                        "Generar TXT",
                        f"Proceso finalizado ✅\n\nTXT principal:\n{txt_path}\n\n"
                        f"Auxiliares en:\n{gen.output_dir}"
                    )
                    return

            txt_path = gen.run(allow_missing=False, skip_prepare=True)
            info(self, "Generar TXT", f"Proceso finalizado ✅\n\nTXT principal:\n{txt_path}\n\nAuxiliares en:\n{gen.output_dir}")

        except RuntimeError as re:
            if str(re) == "FALTAN_MONEDAS_EN_CONFIG":
                info(self, "Monedas", "Faltan monedas en Configuraciones > Monedas. Agregalas y probá de nuevo.")
                return
            if str(re) == "FALTAN_RUTS_EN_PROVEEDORES":
                info(self, "Proveedores", "Faltan RUTs en Proveedores.xlsx. Agregalos y probá de nuevo.")
                return
            log_exception("run_generate RuntimeError", re)
            error(self, "Error", f"Ocurrió un error:\n\n{re}")
        except Exception as e:
            log_exception("run_generate", e)
            error(self, "Error", f"Ocurrió un error:\n\n{e}")

    def on_exit(self):
        self.destroy()
        self.root.quit()


# =========================
#   Empresas
# =========================

class EmpresasWindow(tk.Toplevel):
    def __init__(self, parent, state: AppState, on_changed=None):
        super().__init__(parent)
        apply_window_icon(self)
        self.parent = parent
        self.state = state
        self.on_changed = on_changed

        self.title("Empresas")
        self.geometry("650x230")
        self.resizable(False, False)

        tk.Label(self, text="Seleccioná la carpeta de la empresa:", font=("Segoe UI", 10, "bold")).pack(pady=(12, 6))

        frm = tk.Frame(self)
        frm.pack(fill="x", padx=14)

        self.entry = tk.Entry(frm)
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        tk.Button(frm, text="Buscar...", command=self.browse).pack(side="left")

        frm2 = tk.Frame(self)
        frm2.pack(fill="x", padx=14, pady=(10, 0))

        tk.Button(frm2, text="Seleccionar", height=2, command=self.select).pack(side="left")
        tk.Button(frm2, text="Cerrar", height=2, command=self.destroy).pack(side="right")

        if self.state.company_dir:
            self.entry.insert(0, str(self.state.company_dir))

        tk.Label(
            self,
            text="Al seleccionar: se crea /Datos con Proveedores.xlsx, Parámetros.xlsx, Abreviaturas.xlsx, Monedas.xlsx.\n"
                 "Si aparece /Datos/Proveedores.txt se importa y se borra.",
            justify="left"
        ).pack(pady=10, padx=14, anchor="w")

    def browse(self):
        initial = self.state.cfg.get("last_browse_dir") or ""
        folder = pick_folder(self, "Elegí la carpeta de la empresa", initialdir=initial)
        if folder:
            self.entry.delete(0, tk.END)
            self.entry.insert(0, folder)

    def select(self):
        folder = self.entry.get().strip()
        if not folder:
            info(self, "Empresas", "Seleccione una carpeta válida.")
            return
        p = Path(folder)
        if not p.exists() or not p.is_dir():
            error(self, "Empresas", "La carpeta no existe o no es válida.")
            return

        try:
            ensure_templates(p)
            added, fname = import_proveedores_txt_if_any(p)
            self.state.set_company(p)
            if self.on_changed:
                self.on_changed()

            msg = f"Empresa seleccionada:\n{p}\n\nDatos:\n{company_data_dir(p)}"
            if fname:
                msg += f"\n\nImportado: {fname} (agregados {added})"
            info(self, "Empresas", msg)
            self.destroy()
        except Exception as e:
            log_exception("EmpresasWindow.select", e)
            error(self, "Empresas", f"No se pudo preparar la empresa:\n\n{e}")


# =========================
#   Proveedores UI
# =========================

class ProveedoresWindow(tk.Toplevel):
    def __init__(self, parent, state: AppState, start_rut: Optional[str] = None):
        super().__init__(parent)
        apply_window_icon(self)
        self.parent = parent
        self.state = state

        self.title("Proveedores")
        self.geometry("1040x560")
        self.resizable(False, False)

        ensure_templates(self.state.company_dir)
        self.path = proveedores_xlsx_path(self.state.company_dir)

        self.codes = get_active_currency_codes(self.state.company_dir)
        _ensure_haber_cols_for_active(self.state.company_dir)

        self.df = read_proveedores_df(self.path).fillna("")
        if "RUT" not in self.df.columns or "Nombre" not in self.df.columns:
            raise ValueError("Proveedores.xlsx está corrupto (no tiene RUT/Nombre).")

        changed = False
        for c in PROV_BASE_COLS:
            if c not in self.df.columns:
                self.df[c] = ""
                changed = True
        for code in self.codes:
            col = f"Haber {code}"
            if col not in self.df.columns:
                self.df[col] = ""
                changed = True
        if changed:
            write_proveedores_df(self.path, self.df)

        self.idx = 0 if len(self.df) > 0 else -1
        if start_rut:
            target = _norm_rut(start_rut)
            for i in range(len(self.df)):
                if _norm_rut(self.df.at[i, "RUT"]) == target:
                    self.idx = i
                    break

        self.dirty = False
        self._suspend_dirty = False

        top = tk.Frame(self)
        top.pack(fill="x", padx=12, pady=(10, 0))

        tk.Label(top, text="Buscar (RUT o Nombre):").pack(side="left")
        self.search_var = tk.StringVar()
        tk.Entry(top, textvariable=self.search_var).pack(side="left", fill="x", expand=True, padx=8)
        tk.Button(top, text="Buscar", command=self.on_search).pack(side="left")

        main = tk.Frame(self)
        main.pack(fill="both", expand=True, padx=12, pady=10)

        left = tk.LabelFrame(main, text="Datos", padx=10, pady=10)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right = tk.LabelFrame(main, text="Cuentas", padx=10, pady=10)
        right.pack(side="left", fill="both", expand=True)

        self.vars: Dict[str, tk.StringVar] = {}
        for c in ["RUT", "Nombre", "Debe"]:
            self.vars[c] = tk.StringVar()
        for code in self.codes:
            self.vars[f"Haber {code}"] = tk.StringVar()

        r = 0
        self._add_entry(left, "Nombre", "Nombre", r); r += 1
        self._add_entry(left, "RUT", "RUT", r); r += 1

        tk.Label(left, text="IVA Fijo").grid(row=r, column=0, sticky="e", padx=8, pady=6)
        self.iva_var = tk.StringVar()
        ttk.Combobox(left, textvariable=self.iva_var, values=["", "10", "22"], width=10, state="readonly")\
            .grid(row=r, column=1, sticky="w", padx=8, pady=6)
        r += 1

        tk.Label(left, text="Libro").grid(row=r, column=0, sticky="e", padx=8, pady=6)
        self.libro_var = tk.StringVar()
        ttk.Combobox(left, textvariable=self.libro_var, values=["C", "E"], width=10, state="readonly")\
            .grid(row=r, column=1, sticky="w", padx=8, pady=6)
        r += 1

        tk.Label(left, text="Cont/Cred").grid(row=r, column=0, sticky="e", padx=8, pady=6)
        self.contcred_var = tk.StringVar()
        ttk.Combobox(left, textvariable=self.contcred_var, values=["Crédito", "Contado"], width=12, state="readonly")\
            .grid(row=r, column=1, sticky="w", padx=8, pady=6)

        rr = 0
        self._add_entry(right, "Debe", "Debe", rr); rr += 1
        for code in self.codes:
            self._add_entry(right, f"Haber {code}", f"Haber {code}", rr)
            rr += 1

        bottom = tk.Frame(self)
        bottom.pack(fill="x", padx=12, pady=(0, 12))

        self.lbl_pos = tk.Label(bottom, text="")
        self.lbl_pos.pack(side="left")

        tk.Button(bottom, text="Agregar proveedor +", command=self.add_provider).pack(side="left", padx=10)

        tk.Button(bottom, text="Guardar", command=self.save_current).pack(side="right")
        tk.Button(bottom, text="Siguiente ▶", command=self.next).pack(side="right", padx=8)
        tk.Button(bottom, text="◀ Anterior", command=self.prev).pack(side="right", padx=8)

        for v in self.vars.values():
            v.trace_add("write", lambda *args: self._mark_dirty())
        self.iva_var.trace_add("write", lambda *args: self._mark_dirty())
        self.libro_var.trace_add("write", lambda *args: self._mark_dirty())
        self.contcred_var.trace_add("write", lambda *args: self._mark_dirty())

        self._render_current()

    def _add_entry(self, parent, label, col, row):
        tk.Label(parent, text=label).grid(row=row, column=0, sticky="e", padx=8, pady=6)
        tk.Entry(parent, textvariable=self.vars[col], width=40).grid(row=row, column=1, sticky="w", padx=8, pady=6)

    def _mark_dirty(self):
        if self._suspend_dirty:
            return
        self.dirty = True

    def _render_current(self):
        self._suspend_dirty = True
        try:
            if self.idx < 0 or len(self.df) == 0:
                for v in self.vars.values():
                    v.set("")
                self.iva_var.set("")
                self.libro_var.set("C")
                self.contcred_var.set("Crédito")
                self.lbl_pos.config(text="0 / 0")
                self.dirty = False
                return

            row = self.df.iloc[self.idx].to_dict()
            for c in self.vars.keys():
                self.vars[c].set(str(row.get(c, "")))

            iva = str(row.get("IVA Fijo", "")).strip()
            self.iva_var.set(iva if iva in ("10", "22") else "")

            libro = str(row.get("Libro", "")).strip().upper()
            self.libro_var.set(libro if libro in ("C", "E") else "C")

            cc = str(row.get("Cont/Cred", "")).strip()
            self.contcred_var.set(cc if cc in ("Crédito", "Contado") else "Crédito")

            self.lbl_pos.config(text=f"{self.idx + 1} / {len(self.df)}")
            self.dirty = False
        finally:
            self._suspend_dirty = False

    def _confirm_unsaved(self) -> str:
        if not self.dirty:
            return "discard"
        res = messagebox.askyesnocancel(
            "Cambios sin guardar",
            "Tenés cambios sin guardar.\n\n¿Querés guardar antes de continuar?",
            parent=self
        )
        if res is None:
            return "cancel"
        if res is True:
            return "save"
        return "discard"

    def _handle_unsaved_before_nav(self) -> bool:
        decision = self._confirm_unsaved()
        if decision == "cancel":
            return False
        if decision == "save":
            return bool(self.save_current(silent=True))
        self._render_current()
        return True

    def add_provider(self):
        import pandas as pd
        if len(self.df) > 0 and self.dirty:
            if not self._handle_unsaved_before_nav():
                return

        new_row = {c: "" for c in self.df.columns}
        for c in PROV_BASE_COLS:
            if c not in new_row:
                new_row[c] = ""
        for code in self.codes:
            hc = f"Haber {code}"
            if hc not in new_row:
                new_row[hc] = ""

        new_row["Libro"] = "C"
        new_row["Cont/Cred"] = "Crédito"
        new_row["IVA Fijo"] = ""

        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.idx = len(self.df) - 1
        self._render_current()
        self.dirty = True
        info(self, "Proveedores", "Proveedor nuevo creado.\nCompletá RUT y Nombre y tocá 'Guardar' ✅")

    def save_current(self, silent: bool = False) -> bool:
        if self.idx < 0 or len(self.df) == 0:
            return True

        for c in self.vars.keys():
            self.df.at[self.idx, c] = self.vars[c].get().strip()

        self.df.at[self.idx, "IVA Fijo"] = self.iva_var.get().strip()
        self.df.at[self.idx, "Libro"] = self.libro_var.get().strip()
        self.df.at[self.idx, "Cont/Cred"] = self.contcred_var.get().strip()

        rut = _norm_rut(self.df.at[self.idx, "RUT"])
        nom = str(self.df.at[self.idx, "Nombre"]).strip()

        if not rut or not nom:
            if not silent:
                error(self, "Proveedores", "RUT y Nombre son obligatorios.")
            return False

        ruts = self.df["RUT"].astype(str).map(_norm_rut).tolist()
        duplicates = [i for i, r in enumerate(ruts) if r == rut and i != self.idx]
        if duplicates:
            if not silent:
                error(self, "Proveedores", f"Ya existe un proveedor con ese RUT:\n\n{rut}")
            return False

        write_proveedores_df(self.path, self.df)
        self.dirty = False
        if not silent:
            info(self, "Proveedores", "Guardado ✅")
        return True

    def prev(self):
        if len(self.df) == 0:
            return
        if not self._handle_unsaved_before_nav():
            return
        self.idx = max(0, self.idx - 1)
        self._render_current()

    def next(self):
        if len(self.df) == 0:
            return
        if not self._handle_unsaved_before_nav():
            return
        self.idx = min(len(self.df) - 1, self.idx + 1)
        self._render_current()

    def on_search(self):
        q = self.search_var.get().strip().lower()
        if not q:
            return
        for i in range(len(self.df)):
            rut = str(self.df.at[i, "RUT"]).lower()
            nom = str(self.df.at[i, "Nombre"]).lower()
            if q in rut or q in nom:
                if self.dirty and not self._handle_unsaved_before_nav():
                    return
                self.idx = i
                self._render_current()
                return
        info(self, "Buscar", "No se encontraron coincidencias.")


# =========================
#   MAIN
# =========================

def main():
    root = make_root()
    splash = show_splash(root)

    try:
        check_license_or_exit(root)
        state = AppState()
        win = MainMenu(root, state)
        win.focus_force()
    finally:
        try:
            splash.destroy()
        except Exception:
            pass

    root.mainloop()

if __name__ == "__main__":
    main()