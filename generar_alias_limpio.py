import pandas as pd
import re
import unicodedata
from pathlib import Path

INPUT_FILE = "alias.xlsx"
OUTPUT_FILE = "alias_limpio.xlsx"

# =========================================================
# HELPERS
# =========================================================
def clean_str(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.strip()

def norm_text(s):
    s = clean_str(s).upper()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def norm_col(s):
    s = norm_text(s)
    s = s.replace(".", "")
    s = s.replace("/", "_")
    s = s.replace("-", "_")
    s = re.sub(r"[^A-Z0-9_ ]+", "", s)
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_")

def clean_cuit(s):
    s = clean_str(s)
    s = re.sub(r"\D", "", s)
    return s

def clean_code(s):
    s = norm_text(s)
    s = s.replace("CONSORCIO", "").strip()
    s = re.sub(r"[^A-Z0-9]+", "", s)
    return s

def make_unique(cols):
    out = []
    used = {}
    for c in cols:
        base = c if c else "COL"
        if base not in used:
            used[base] = 1
            out.append(base)
        else:
            used[base] += 1
            out.append(f"{base}_{used[base]}")
    return out

def first_nonempty(*vals):
    for v in vals:
        v = clean_str(v)
        if v and norm_text(v) not in ["NAN", "NONE", "NULL"]:
            return v
    return ""

def add_alias(alias_rows, codigo, alias, tipo="MANUAL", prioridad=10):
    alias = clean_str(alias)
    codigo = clean_code(codigo)
    if not codigo or not alias:
        return
    key = (codigo, norm_text(alias))
    if key in add_alias.seen:
        return
    add_alias.seen.add(key)
    alias_rows.append({
        "CODIGO_CONSORCIO": codigo,
        "ALIAS": alias,
        "TIPO": tipo,
        "PRIORIDAD": prioridad,
        "ACTIVO": "SI",
        "OBSERVACIONES": ""
    })
add_alias.seen = set()

# =========================================================
# LEER ARCHIVO
# =========================================================
if not Path(INPUT_FILE).exists():
    raise FileNotFoundError(f"No existe el archivo {INPUT_FILE}")

xls = pd.ExcelFile(INPUT_FILE)

# =========================================================
# 1) CONSORCIOS
# =========================================================
df_cons = pd.read_excel(xls, sheet_name="CONSORCIOS", dtype=str)
df_cons.columns = [norm_col(c) for c in df_cons.columns]

# corregir nombre mal escrito
if "DOMILICIO" in df_cons.columns and "DOMICILIO" not in df_cons.columns:
    df_cons = df_cons.rename(columns={"DOMILICIO": "DOMICILIO"})

# limpiar todas las celdas
for col in df_cons.columns:
    df_cons[col] = df_cons[col].apply(clean_str)

# detectar columnas si vienen con nombres parecidos
col_alias = "ALIAS_CONSORCIO" if "ALIAS_CONSORCIO" in df_cons.columns else ("ALIAS CONSORCIO" if "ALIAS CONSORCIO" in df_cons.columns else None)
col_consorcio = "CONSORCIO" if "CONSORCIO" in df_cons.columns else None
col_razon = "RAZON_SOCIAL" if "RAZON_SOCIAL" in df_cons.columns else ("RAZON SOCIAL" if "RAZON SOCIAL" in df_cons.columns else None)
col_domicilio = "DOMICILIO" if "DOMICILIO" in df_cons.columns else None
col_cuit = "CUIT" if "CUIT" in df_cons.columns else None

consorcios_rows = []
alias_rows = []
servicios_rows = []

for _, row in df_cons.iterrows():
    alias_val = first_nonempty(
        row.get(col_alias, "") if col_alias else "",
        row.get(col_consorcio, "") if col_consorcio else "",
        row.get(col_razon, "") if col_razon else "",
        row.get(col_domicilio, "") if col_domicilio else ""
    )

    if not alias_val:
        continue

    codigo = clean_code(row.get(col_alias, "")) if col_alias else ""
    if not codigo:
        codigo = clean_code(alias_val)
    if not codigo:
        # fallback simple con las palabras importantes
        codigo = re.sub(r"[^A-Z0-9]+", "", norm_text(alias_val))[:12]

    nombre_canonico = first_nonempty(
        row.get(col_consorcio, "") if col_consorcio else "",
        row.get(col_razon, "") if col_razon else "",
        row.get(col_domicilio, "") if col_domicilio else ""
    )

    cuit = clean_cuit(row.get(col_cuit, "")) if col_cuit else ""

    domicilio = row.get(col_domicilio, "") if col_domicilio else ""
    razon_social = row.get(col_razon, "") if col_razon else ""

    consorcios_rows.append({
        "CODIGO_CONSORCIO": codigo,
        "NOMBRE_CANONICO": nombre_canonico,
        "CUIT": cuit,
        "RAZON_SOCIAL": razon_social,
        "DOMICILIO": domicilio,
        "LOCALIDAD": "",
        "CP": "",
        "ACTIVO": "SI",
        "OBSERVACIONES": ""
    })

    # aliases automáticos desde los campos de la fila
    add_alias(alias_rows, codigo, nombre_canonico, tipo="MANUAL", prioridad=10)
    add_alias(alias_rows, codigo, row.get(col_alias, "") if col_alias else "", tipo="MANUAL", prioridad=10)
    add_alias(alias_rows, codigo, row.get(col_consorcio, "") if col_consorcio else "", tipo="MANUAL", prioridad=10)
    add_alias(alias_rows, codigo, row.get(col_razon, "") if col_razon else "", tipo="MANUAL", prioridad=8)
    add_alias(alias_rows, codigo, row.get(col_domicilio, "") if col_domicilio else "", tipo="MANUAL", prioridad=7)
    if "ESPECIAL" in df_cons.columns:
        add_alias(alias_rows, codigo, row.get("ESPECIAL", ""), tipo="MANUAL", prioridad=6)

    # servicios
    for serv in ["AYSA", "METROGAS", "EDENOR"]:
        if serv in df_cons.columns:
            nro = clean_cuit(row.get(serv, ""))
            if nro:
                servicios_rows.append({
                    "CODIGO_CONSORCIO": codigo,
                    "SERVICIO": serv,
                    "NRO_CLIENTE": nro,
                    "CUIT_EMPRESA": "",
                    "ACTIVO": "SI",
                    "OBSERVACIONES": ""
                })

df_cons_limpio = pd.DataFrame(consorcios_rows).drop_duplicates(subset=["CODIGO_CONSORCIO"]).reset_index(drop=True)
df_alias_limpio = pd.DataFrame(alias_rows).drop_duplicates(subset=["CODIGO_CONSORCIO", "ALIAS"]).reset_index(drop=True)
df_servicios_limpio = pd.DataFrame(servicios_rows).drop_duplicates(subset=["CODIGO_CONSORCIO", "SERVICIO", "NRO_CLIENTE"]).reset_index(drop=True)

# =========================================================
# 2) PROVEEDORES
# =========================================================
df_prov = pd.read_excel(xls, sheet_name="PROVEEDORES", dtype=str)
df_prov.columns = [norm_col(c) for c in df_prov.columns]

for col in df_prov.columns:
    df_prov[col] = df_prov[col].apply(clean_str)

prov_rows = []
for _, row in df_prov.iterrows():
    nombre_real = first_nonempty(row.get("NOMBRE_REAL", ""), row.get("NOMBRE REAL", ""))
    alias_prov = first_nonempty(row.get("ALIAS_PROV", ""), row.get("ALIAS PROV", ""))
    if not nombre_real and not alias_prov:
        continue

    if not alias_prov:
        alias_prov = nombre_real[:20].upper()

    # palabras clave sugeridas simples
    palabras = []
    if alias_prov:
        palabras.append(norm_text(alias_prov).lower())
    if nombre_real:
        palabras.append(norm_text(nombre_real).lower())

    prov_rows.append({
        "NOMBRE_REAL": nombre_real,
        "ALIAS_PROV": alias_prov,
        "CUIT": clean_cuit(row.get("CUIT", "")) if "CUIT" in df_prov.columns else "",
        "PALABRAS_CLAVE": "; ".join([p for p in palabras if p]),
        "ACTIVO": "SI",
        "OBSERVACIONES": ""
    })

df_prov_limpio = pd.DataFrame(prov_rows).drop_duplicates(subset=["NOMBRE_REAL", "ALIAS_PROV"]).reset_index(drop=True)

# =========================================================
# 3) PLANILLA DE PAGOS RAW
# =========================================================
# Leemos sin header para preservar exactamente lo que hay
raw_plan = pd.read_excel(xls, sheet_name="PLANILLA DE PAGOS", header=None, dtype=str)
raw_plan = raw_plan.fillna("")

# Tomamos la primera fila como encabezados "base"
base_headers = [norm_col(x) if clean_str(x) else "" for x in raw_plan.iloc[0].tolist()]
base_headers = make_unique(base_headers)

plan_raw = raw_plan.iloc[1:].reset_index(drop=True).copy()
plan_raw.columns = base_headers

# limpiar celdas
for col in plan_raw.columns:
    plan_raw[col] = plan_raw[col].apply(clean_str)

# quitar filas totalmente vacías
plan_raw = plan_raw.loc[~(plan_raw.apply(lambda r: all(clean_str(v) == "" for v in r), axis=1))].copy()

# =========================================================
# 4) PLANILLA PAGO LIMPIA / NORMALIZADA
# =========================================================
# intentamos usar el formato de tu hoja actual pero con nombres estables
# si no existen algunos campos, quedan vacíos
def get_col(df, wanted):
    wanted_norm = norm_col(wanted)
    for c in df.columns:
        if norm_col(c) == wanted_norm:
            return c
    return None

col_cons = get_col(plan_raw, "CONSORCIO")
if not col_cons:
    # si el excel está raro, usamos la primera columna con datos como código
    col_cons = plan_raw.columns[0]

normalized = pd.DataFrame()
normalized["CODIGO_CONSORCIO"] = plan_raw[col_cons].apply(lambda x: clean_code(x) if x else clean_str(x))

# Mapeo de columnas más comunes en tu planilla actual
mapeos = {
    "AYSA": "AYSA",
    "METROGAS": "METROGAS",
    "EDENOR": "EDENOR",
    "ASCENSOR": "ASCENSOR",
    "DESIN.": "DESINFECCION",
    "DESINFECCION": "DESINFECCION",
    "LIMP TAN.": "LIMPIEZA_TANQUES",
    "LIMPIEZA": "LIMPIEZA",
    "CERT CALD.": "CERT_CALD",
    "CERT CALD": "CERT_CALD",
    "ABONO VAR.": "ABONO_VAR",
    "ABONO VAR": "ABONO_VAR",
    "JARDIN.": "JARDIN",
    "JARDIN": "JARDIN",
    "CERT INCE.": "CERT_INCE",
    "CERT INCE": "CERT_INCE",
    "ABL": "ABL",
    "FIBERCORP": "FIBERCORP"
}

for src, dst in mapeos.items():
    c = get_col(plan_raw, src)
    if c:
        normalized[dst] = plan_raw[c].apply(clean_str)
    else:
        normalized[dst] = ""

normalized["OBSERVACIONES"] = ""
normalized = normalized.drop_duplicates(subset=["CODIGO_CONSORCIO"]).reset_index(drop=True)

# =========================================================
# 5) EXPORTAR EXCEL
# =========================================================
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_cons_limpio.to_excel(writer, sheet_name="CONSORCIOS", index=False)
    df_alias_limpio.to_excel(writer, sheet_name="CONSORCIO_ALIAS", index=False)
    df_servicios_limpio.to_excel(writer, sheet_name="SERVICIOS_CLIENTES", index=False)
    df_prov_limpio.to_excel(writer, sheet_name="PROVEEDORES", index=False)
    normalized.to_excel(writer, sheet_name="PLANILLA_PAGOS", index=False)
    plan_raw.to_excel(writer, sheet_name="PLANILLA_PAGOS_RAW", index=False)

print(f"OK -> archivo generado: {OUTPUT_FILE}")
print("Hojas creadas: CONSORCIOS, CONSORCIO_ALIAS, SERVICIOS_CLIENTES, PROVEEDORES, PLANILLA_PAGOS, PLANILLA_PAGOS_RAW")