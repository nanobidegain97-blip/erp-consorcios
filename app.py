
import streamlit as st
import fitz
import pdfplumber
from openai import OpenAI
import json
import pandas as pd
from io import BytesIO
import os
import base64
import time
import sqlite3
import datetime
import re
import unicodedata
import zipfile  # Módulo para comprimir archivos
from difflib import SequenceMatcher

st.set_page_config(page_title="Sistema ERP Consorcios", layout="wide")

# =========================================================
# CONFIG OPENAI
# =========================================================
OPENAI_API_KEY = "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"

try:
    if "OPENAI_API_KEY" in st.secrets:
        OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]
except Exception:
    pass

if not OPENAI_API_KEY:
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")

client = OpenAI(api_key="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") if OPENAI_API_KEY else None

# =========================================================
# RUTAS
# =========================================================
RUTA_EXCEL = "alias_limpio.xlsx"
DB_FILE = "consorcios.db"

# =========================================================
# HELPERS
# =========================================================
def limpiar_str(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    if x.endswith(".0"):
        x = x[:-2]
    return x.strip()

def normalizar_texto(x):
    if pd.isna(x):
        return ""
    x = limpiar_str(x).upper()
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    x = re.sub(r"[^A-Z0-9]+", " ", x)
    x = re.sub(r"\s+", " ", x).strip()
    return x

def normalizar_codigo(x):
    x = normalizar_texto(x)
    x = x.replace("CONSORCIO", "").strip()
    x = re.sub(r"[^A-Z0-9]", "", x)
    return x

def normalizar_cuit(x):
    return re.sub(r"\D", "", limpiar_str(x))

def similitud(a, b):
    a = normalizar_texto(a)
    b = normalizar_texto(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

def es_activo(x):
    return normalizar_texto(x) in ["SI", "S", "1", "TRUE", "YES", "ACTIVO"]

def parse_codigo_desde_display(texto):
    texto = limpiar_str(texto)
    if not texto or texto == "SIN IDENTIFICAR":
        return "SIN IDENTIFICAR"
    if " - " in texto:
        posible = normalizar_codigo(texto.split(" - ", 1)[0])
        if posible:
            return posible
    return normalizar_codigo(texto) or "SIN IDENTIFICAR"

# =========================================================
# BASE DE DATOS Y MIGRACIONES
# =========================================================
def agregar_columna_si_no_existe(cursor, tabla, columna, definicion):
    try:
        cursor.execute(f"ALTER TABLE {tabla} ADD COLUMN {columna} {definicion}")
    except sqlite3.OperationalError:
        pass

def inicializar_bd():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS facturas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            consorcio_codigo TEXT,
            consorcio TEXT,
            proveedor TEXT,
            monto REAL,
            fecha_emision TEXT,
            fecha_vencimiento TEXT,
            numero_factura TEXT,
            nombre_archivo TEXT,
            estado TEXT DEFAULT 'Pendiente',
            fecha_carga TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS consorcios (
            codigo TEXT PRIMARY KEY,
            nombre_canonico TEXT,
            cuit TEXT,
            domicilio TEXT,
            activo INTEGER DEFAULT 1
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS consorcio_aliases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            consorcio_codigo TEXT,
            alias TEXT UNIQUE,
            tipo TEXT DEFAULT 'MANUAL',
            prioridad INTEGER DEFAULT 10,
            fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (consorcio_codigo) REFERENCES consorcios(codigo)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS servicios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            consorcio_codigo TEXT,
            servicio TEXT,
            nro_cliente TEXT,
            FOREIGN KEY (consorcio_codigo) REFERENCES consorcios(codigo)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS proveedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_real TEXT UNIQUE,
            alias_prov TEXT,
            palabras_clave TEXT
        )
    """)

    agregar_columna_si_no_existe(cursor, "facturas", "periodo", "TEXT DEFAULT 'Historial Viejo'")
    agregar_columna_si_no_existe(cursor, "facturas", "texto_detectado", "TEXT")
    agregar_columna_si_no_existe(cursor, "facturas", "nombre_leido", "TEXT")
    agregar_columna_si_no_existe(cursor, "facturas", "direccion_leida", "TEXT")
    agregar_columna_si_no_existe(cursor, "facturas", "cuit_leido", "TEXT")
    agregar_columna_si_no_existe(cursor, "facturas", "cuenta_leida", "TEXT")
    agregar_columna_si_no_existe(cursor, "facturas", "score_match", "REAL DEFAULT 0")
    agregar_columna_si_no_existe(cursor, "facturas", "motivo_match", "TEXT")
    agregar_columna_si_no_existe(cursor, "facturas", "consorcio_sugerido", "TEXT")

    conn.commit()
    conn.close()

inicializar_bd()

# =========================================================
# CARGA Y SINCRONIZACION EXCEL -> SQLITE
# =========================================================
def cargar_diccionarios_excel():
    df_cons = pd.DataFrame()
    df_alias = pd.DataFrame()
    df_serv = pd.DataFrame()
    df_prov = pd.DataFrame()
    df_planilla = pd.DataFrame()
    df_planilla_raw = pd.DataFrame()

    if os.path.exists(RUTA_EXCEL):
        try:
            df_cons = pd.read_excel(RUTA_EXCEL, sheet_name="CONSORCIOS", dtype=str).fillna("")
            df_alias = pd.read_excel(RUTA_EXCEL, sheet_name="CONSORCIO_ALIAS", dtype=str).fillna("")
            df_serv = pd.read_excel(RUTA_EXCEL, sheet_name="SERVICIOS_CLIENTES", dtype=str).fillna("")
            df_prov = pd.read_excel(RUTA_EXCEL, sheet_name="PROVEEDORES", dtype=str).fillna("")
            df_planilla = pd.read_excel(RUTA_EXCEL, sheet_name="PLANILLA_PAGOS", dtype=str).fillna("")
            df_planilla_raw = pd.read_excel(RUTA_EXCEL, sheet_name="PLANILLA_PAGOS_RAW", dtype=str).fillna("")
        except Exception as e:
            st.error(f"Error leyendo {RUTA_EXCEL}: {e}")

    return df_cons, df_alias, df_serv, df_prov, df_planilla, df_planilla_raw

df_cons, df_alias, df_serv, df_prov, df_planilla, df_planilla_raw = cargar_diccionarios_excel()

col_cons_codigo = "CODIGO_CONSORCIO"
col_cons_nombre = "NOMBRE_CANONICO"

def sincronizar_sqlite_desde_excel():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM consorcios")
    if cursor.fetchone()[0] == 0 and not df_cons.empty:
        for _, row in df_cons.iterrows():
            codigo = normalizar_codigo(row.get(col_cons_codigo, ""))
            nombre = limpiar_str(row.get(col_cons_nombre, ""))
            cuit = normalizar_cuit(row.get("CUIT", ""))
            domicilio = limpiar_str(row.get("DOMICILIO", ""))
            activo = 1 if es_activo(row.get("ACTIVO", "SI")) else 0
            if codigo:
                cursor.execute("""
                    INSERT OR REPLACE INTO consorcios (codigo, nombre_canonico, cuit, domicilio, activo)
                    VALUES (?, ?, ?, ?, ?)
                """, (codigo, nombre, cuit, domicilio, activo))

    cursor.execute("SELECT COUNT(*) FROM consorcio_aliases")
    if cursor.fetchone()[0] == 0 and not df_alias.empty:
        for _, row in df_alias.iterrows():
            codigo = normalizar_codigo(row.get("CODIGO_CONSORCIO", ""))
            alias = limpiar_str(row.get("ALIAS", ""))
            tipo = limpiar_str(row.get("TIPO", "MANUAL")) or "MANUAL"
            prioridad = 10
            try:
                prioridad = int(float(row.get("PRIORIDAD", 10)))
            except:
                pass
            if codigo and alias:
                cursor.execute("""
                    INSERT OR IGNORE INTO consorcio_aliases (consorcio_codigo, alias, tipo, prioridad)
                    VALUES (?, ?, ?, ?)
                """, (codigo, alias, tipo, prioridad))

    cursor.execute("SELECT COUNT(*) FROM servicios")
    if cursor.fetchone()[0] == 0 and not df_serv.empty:
        for _, row in df_serv.iterrows():
            codigo = normalizar_codigo(row.get("CODIGO_CONSORCIO", ""))
            servicio = limpiar_str(row.get("SERVICIO", ""))
            nro = normalizar_cuit(row.get("NRO_CLIENTE", ""))
            if codigo and servicio:
                cursor.execute("""
                    INSERT INTO servicios (consorcio_codigo, servicio, nro_cliente)
                    VALUES (?, ?, ?)
                """, (codigo, servicio, nro))

    cursor.execute("SELECT COUNT(*) FROM proveedores")
    if cursor.fetchone()[0] == 0 and not df_prov.empty:
        for _, row in df_prov.iterrows():
            nombre_real = limpiar_str(row.get("NOMBRE_REAL", ""))
            alias_prov = limpiar_str(row.get("ALIAS_PROV", ""))
            palabras_clave = limpiar_str(row.get("PALABRAS_CLAVE", ""))
            if nombre_real:
                cursor.execute("""
                    INSERT OR IGNORE INTO proveedores (nombre_real, alias_prov, palabras_clave)
                    VALUES (?, ?, ?)
                """, (nombre_real, alias_prov, palabras_clave))

    conn.commit()
    conn.close()

sincronizar_sqlite_desde_excel()

# =========================================================
# MAPAS DESDE SQLITE (DINÁMICOS)
# =========================================================
def construir_mapas_desde_sqlite():
    mapa_codigo_nombre = {}
    mapa_nombre_codigo = {}
    mapa_alias_codigo = {}
    mapa_servicios = {}

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("SELECT codigo, nombre_canonico, cuit FROM consorcios WHERE activo = 1")
    for codigo, nombre, cuit in cursor.fetchall():
        codigo = normalizar_codigo(codigo)
        if codigo and nombre:
            mapa_codigo_nombre[codigo] = nombre
            mapa_nombre_codigo[normalizar_texto(nombre)] = codigo
        if codigo and cuit:
            mapa_servicios[("CUIT", cuit)] = codigo

    cursor.execute("SELECT consorcio_codigo, alias FROM consorcio_aliases")
    for codigo, alias in cursor.fetchall():
        codigo = normalizar_codigo(codigo)
        alias_norm = normalizar_texto(alias)
        if codigo and alias_norm:
            mapa_alias_codigo[alias_norm] = codigo

    cursor.execute("SELECT consorcio_codigo, servicio, nro_cliente FROM servicios")
    for codigo, servicio, nro in cursor.fetchall():
        codigo = normalizar_codigo(codigo)
        servicio_norm = normalizar_texto(servicio)
        nro_norm = normalizar_cuit(nro)
        if codigo and servicio_norm and nro_norm:
            mapa_servicios[(servicio_norm, nro_norm)] = codigo

    conn.close()
    return mapa_codigo_nombre, mapa_nombre_codigo, mapa_alias_codigo, mapa_servicios

mapa_codigo_nombre, mapa_nombre_codigo, mapa_alias_codigo, mapa_servicios = construir_mapas_desde_sqlite()

def formato_consorcio(codigo):
    codigo = normalizar_codigo(codigo)
    nombre = mapa_codigo_nombre.get(codigo, "")
    if codigo == "SIN IDENTIFICAR":
        return "SIN IDENTIFICAR"
    if nombre:
        return f"{codigo} - {nombre}"
    return codigo

def resolver_codigo_consorcio(texto):
    texto = limpiar_str(texto)
    if not texto or texto == "SIN IDENTIFICAR":
        return "SIN IDENTIFICAR"

    if " - " in texto:
        candidato = normalizar_codigo(texto.split(" - ", 1)[0])
        if candidato in mapa_codigo_nombre:
            return candidato

    codigo = normalizar_codigo(texto)
    if codigo in mapa_codigo_nombre:
        return codigo

    texto_norm = normalizar_texto(texto)

    if texto_norm in mapa_nombre_codigo:
        return mapa_nombre_codigo[texto_norm]

    if texto_norm in mapa_alias_codigo:
        return mapa_alias_codigo[texto_norm]

    mejor = ("SIN IDENTIFICAR", 0.0)
    for norm_nombre, codigo_posible in mapa_nombre_codigo.items():
        s = similitud(texto_norm, norm_nombre)
        if s > mejor[1]:
            mejor = (codigo_posible, s)

    if mejor[1] >= 0.88:
        return mejor[0]

    for alias_norm, codigo_posible in mapa_alias_codigo.items():
        s = similitud(texto_norm, alias_norm)
        if s > mejor[1]:
            mejor = (codigo_posible, s)

    if mejor[1] >= 0.88:
        return mejor[0]

    return "SIN IDENTIFICAR"

def resolver_nombre_consorcio(texto):
    codigo = resolver_codigo_consorcio(texto)
    return formato_consorcio(codigo) if codigo != "SIN IDENTIFICAR" else "SIN IDENTIFICAR"

def guardar_alias_detectado(alias, codigo, tipo="OCR", prioridad=8, conn=None):
    alias = limpiar_str(alias)
    codigo = normalizar_codigo(codigo)

    if not alias or not codigo or codigo == "SIN IDENTIFICAR":
        return

    if normalizar_texto(alias) == normalizar_texto(codigo):
        return

    close_conn = False
    if conn is None:
        conn = sqlite3.connect(DB_FILE, timeout=30)
        conn.execute("PRAGMA busy_timeout = 30000;")
        close_conn = True

    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT OR IGNORE INTO consorcio_aliases (consorcio_codigo, alias, tipo, prioridad)
            VALUES (?, ?, ?, ?)
        """, (codigo, alias, tipo, prioridad))

        if close_conn:
            conn.commit()
    except Exception:
        pass
    finally:
        if close_conn:
            conn.close()

def resolver_proveedor_alias(nombre_prov):
    nombre_prov = limpiar_str(nombre_prov)
    if not nombre_prov:
        return ""

    conn = sqlite3.connect(DB_FILE)
    df_prov_db = pd.read_sql_query("SELECT * FROM proveedores", conn)
    conn.close()

    if df_prov_db.empty:
        return nombre_prov.upper()

    prov_norm = normalizar_texto(nombre_prov)

    for _, row in df_prov_db.iterrows():
        nombre_real = limpiar_str(row.get("nombre_real", ""))
        alias_prov = limpiar_str(row.get("alias_prov", ""))
        palabras = limpiar_str(row.get("palabras_clave", "")).split(";")

        opciones = [nombre_real, alias_prov] + palabras
        for op in opciones:
            op_norm = normalizar_texto(op)
            if not op_norm:
                continue
            if op_norm in prov_norm or prov_norm in op_norm:
                return alias_prov.upper() if alias_prov else nombre_real.upper()

    return nombre_prov.upper()

def obtener_contexto_ia(df_c, df_p):
    texto_contexto = "\n\nAYUDA DE CONTEXTO (Bases de datos de la empresa):"

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT nombre_canonico FROM consorcios WHERE activo = 1")
    consorcios_db = [r[0] for r in cursor.fetchall()]

    cursor.execute("SELECT alias_prov, nombre_real FROM proveedores")
    provs_db = []
    for alias_p, nombre_r in cursor.fetchall():
        provs_db.append(alias_p if alias_p else nombre_r)
    conn.close()

    if consorcios_db:
        nombres_cons = ", ".join(consorcios_db)
        texto_contexto += f"\n- Consorcios/Edificios válidos: {nombres_cons}"

    if provs_db:
        nombres_prov = ", ".join(list(filter(None, provs_db)))
        texto_contexto += f"\n- Proveedores frecuentes: {nombres_prov}"

    texto_contexto += "\n(Si reconoces estos nombres o direcciones en la factura, devuélvelos exactamente igual)."
    return texto_contexto

def obtener_candidatos_consorcio(info, df_alias):
    conn = sqlite3.connect(DB_FILE)
    df_cons_db = pd.read_sql_query("SELECT * FROM consorcios WHERE activo = 1", conn)
    df_serv_db = pd.read_sql_query("SELECT * FROM servicios", conn)
    df_alias_db = pd.read_sql_query("SELECT * FROM consorcio_aliases", conn)
    conn.close()

    if df_cons_db.empty:
        return []

    cuit_ia = normalizar_cuit(info.get("cuit_consorcio", ""))
    cuenta_ia = normalizar_cuit(info.get("nro_cuenta_cliente", ""))
    nombre_ia = normalizar_texto(info.get("nombre_consorcio_leido", ""))
    direccion_ia = normalizar_texto(info.get("direccion_consorcio_leida", ""))

    candidatos = []

    for _, row in df_cons_db.iterrows():
        codigo = normalizar_codigo(row.get("codigo", ""))
        if not codigo:
            continue

        nombre_canonico = limpiar_str(row.get("nombre_canonico", ""))
        cuit_row = normalizar_cuit(row.get("cuit", ""))
        domicilio_row = normalizar_texto(row.get("domicilio", ""))

        score = 0
        motivos = []

        if cuit_ia and cuit_row and cuit_ia == cuit_row:
            score = 100
            motivos.append("CUIT exacto")

        if cuenta_ia and not df_serv_db.empty:
            match_serv = df_serv_db[
                (df_serv_db["consorcio_codigo"].apply(normalizar_codigo) == codigo) &
                (df_serv_db["nro_cliente"].apply(normalizar_cuit) == cuenta_ia)
            ]
            if not match_serv.empty:
                score = max(score, 95)
                motivos.append("Cuenta cliente exacta")

        texto_ref = " ".join([nombre_canonico, domicilio_row]).strip()
        texto_ref_norm = normalizar_texto(texto_ref)

        if nombre_ia:
            if nombre_ia in texto_ref_norm or texto_ref_norm in nombre_ia:
                score = max(score, 90)
                motivos.append("Nombre coincide")
            else:
                s = similitud(nombre_ia, texto_ref_norm)
                if s >= 0.82:
                    score = max(score, int(s * 100))
                    motivos.append("Nombre similar")

        if direccion_ia:
            if direccion_ia in texto_ref_norm or texto_ref_norm in direccion_ia:
                score = max(score, 85)
                motivos.append("Dirección coincide")
            else:
                s = similitud(direccion_ia, texto_ref_norm)
                if s >= 0.80:
                    score = max(score, int(s * 100))
                    motivos.append("Dirección similar")

        if not df_alias_db.empty:
            aliases = df_alias_db[df_alias_db["consorcio_codigo"].apply(normalizar_codigo) == codigo]
            for _, arow in aliases.iterrows():
                alias_norm = normalizar_texto(arow.get("alias", ""))
                if alias_norm and nombre_ia:
                    if alias_norm in nombre_ia or nombre_ia in alias_norm:
                        score = max(score, 88)
                        motivos.append("Alias coincide")
                        break
                    s = similitud(nombre_ia, alias_norm)
                    if s >= 0.82:
                        score = max(score, int(s * 100))
                        motivos.append("Alias similar")
                        break

        if score > 0:
            candidatos.append({
                "codigo": codigo,
                "nombre": nombre_canonico,
                "score": score,
                "motivo": " | ".join(list(dict.fromkeys(motivos)))
            })

    candidatos = sorted(candidatos, key=lambda x: x["score"], reverse=True)

    vistos = set()
    out = []
    for c in candidatos:
        if c["codigo"] not in vistos:
            vistos.add(c["codigo"])
            out.append(c)

    return out[:5]

# =========================================================
# IA
# =========================================================
def extraer_datos_ia_vision(pdf_bytes, contexto):
    if client is None:
        return {"error": "Falta configurar OPENAI_API_KEY"}

    base64_images = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for i in range(min(len(doc), 2)):
            pix = doc.load_page(i).get_pixmap(matrix=fitz.Matrix(2, 2))
            base64_images.append(base64.b64encode(pix.tobytes("png")).decode("utf-8"))
    except Exception as e:
        return {"error": f"Error: {e}"}

    prompt = (
        "Analiza las IMÁGENES y extrae en JSON. "
        "Claves: fecha_emision, vencimiento_completo, año_emision, mes_emision, monto_a_pagar, numero_factura, "
        "cuit_consorcio, nro_cuenta_cliente, proveedor, nombre_consorcio_leido, direccion_consorcio_leida."
        + contexto
    )

    mensajes = [{"role": "user", "content": [{"type": "text", "text": prompt}]}]
    for b64_img in base64_images:
        mensajes[0]["content"].append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{b64_img}", "detail": "high"}
        })

    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=mensajes,
            response_format={"type": "json_object"},
            temperature=0.0
        )
        return json.loads(resp.choices[0].message.content)
    except Exception as e:
        return {"error": str(e)}

def extraer_datos_ia_texto(texto, contexto):
    if client is None:
        return {"error": "Falta configurar OPENAI_API_KEY"}

    prompt = (
        "Analiza este TEXTO y extrae en JSON. "
        "Claves: fecha_emision, vencimiento_completo, año_emision, mes_emision, monto_a_pagar, numero_factura, "
        "cuit_consorcio, nro_cuenta_cliente, proveedor, nombre_consorcio_leido, direccion_consorcio_leida.\n\n"
        f"TEXTO: {texto}"
        + contexto
    )

    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0.0
        )
        return json.loads(resp.choices[0].message.content)
    except Exception as e:
        return {"error": str(e)}

# =========================================================
# DB OPERACIONES
# =========================================================
def guardar_cambios_consorcio(df_editado, codigo_consorcio, periodo_actual):
    conn = sqlite3.connect(DB_FILE, timeout=30)
    cursor = conn.cursor()

    for _, row in df_editado.iterrows():
        id_db = row.get("id_db", None)
        nuevo_estado = row.get("Estado", "Pendiente")
        nuevo_monto = row.get("Monto", 0)
        nuevo_numero = row.get("Factura", "-")

        if pd.notna(id_db) and str(id_db).strip() != "":
            cursor.execute("""
                UPDATE facturas
                SET estado = ?, monto = ?, numero_factura = ?
                WHERE id = ?
            """, (nuevo_estado, nuevo_monto, nuevo_numero, int(id_db)))

    conn.commit()
    conn.close()

def reasignar_facturas_sin_identificar(df_editado):
    conn = sqlite3.connect(DB_FILE, timeout=30)
    cursor = conn.cursor()

    for _, row in df_editado.iterrows():
        id_db = row.get("id_db", None)
        consorcio_display = limpiar_str(row.get("Consorcio", ""))
        codigo_nuevo = parse_codigo_desde_display(consorcio_display)
        texto_detectado = limpiar_str(row.get("Texto Detectado", ""))
        nombre_leido = limpiar_str(row.get("Nombre Leido", ""))

        if pd.notna(id_db) and str(id_db).strip() != "" and codigo_nuevo != "SIN IDENTIFICAR":
            cursor.execute("""
                UPDATE facturas
                SET consorcio_codigo = ?, consorcio = ?
                WHERE id = ?
            """, (codigo_nuevo, formato_consorcio(codigo_nuevo), int(id_db)))

            if nombre_leido:
                guardar_alias_detectado(nombre_leido, codigo_nuevo, tipo="MANUAL", prioridad=9)
            elif texto_detectado:
                guardar_alias_detectado(texto_detectado, codigo_nuevo, tipo="MANUAL", prioridad=9)

    conn.commit()
    conn.close()

def eliminar_factura_db(id_factura):
    conn = sqlite3.connect(DB_FILE, timeout=30)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM facturas WHERE id = ?", (int(id_factura),))
    conn.commit()
    conn.close()

# =========================================================
# MATRIZ DE CONTROL DINÁMICA
# =========================================================
def generar_matriz_control(periodo_actual):
    conn = sqlite3.connect(DB_FILE)
    db_facturas = pd.read_sql_query(
        "SELECT * FROM facturas WHERE periodo = ?",
        conn,
        params=(periodo_actual,)
    )
    df_serv_db = pd.read_sql_query("SELECT * FROM servicios", conn)
    df_cons_db = pd.read_sql_query("SELECT * FROM consorcios WHERE activo = 1", conn)
    conn.close()

    if df_serv_db.empty:
        return pd.DataFrame(columns=[
            "id_db", "Codigo Consorcio", "Consorcio", "Servicio", "Estado",
            "Monto", "Factura", "Texto Detectado", "Nombre Leido"
        ])

    df_serv_db["consorcio_codigo"] = df_serv_db["consorcio_codigo"].apply(normalizar_codigo)
    db_facturas["consorcio_codigo"] = db_facturas["consorcio_codigo"].apply(normalizar_codigo)
    db_facturas["proveedor_norm"] = db_facturas["proveedor"].astype(str).apply(normalizar_texto)

    codigos_validos = set(df_cons_db["codigo"].tolist())

    matriz = []
    ids_procesados = set()
    claves_procesadas = set()

    for _, row in df_serv_db.iterrows():
        codigo = row["consorcio_codigo"]
        if codigo not in codigos_validos:
            continue

        servicio = row["servicio"]
        nro_cliente = row["nro_cliente"]

        clave = (codigo, normalizar_texto(servicio), normalizar_cuit(nro_cliente))
        if clave in claves_procesadas:
            continue
        claves_procesadas.add(clave)

        servicio_norm = normalizar_texto(servicio)
        alias_servicio = normalizar_texto(resolver_proveedor_alias(servicio)) if servicio else ""

        if servicio_norm:
            match = db_facturas[
                (db_facturas["consorcio_codigo"] == codigo) &
                (
                    db_facturas["proveedor_norm"].apply(
                        lambda x: (
                            servicio_norm in x or x in servicio_norm or
                            alias_servicio in x or x in alias_servicio
                        )
                    )
                )
            ]
        else:
            match = db_facturas[db_facturas["consorcio_codigo"] == codigo]

        if not match.empty:
            for _, f in match.iterrows():
                if f["id"] not in ids_procesados:
                    ids_procesados.add(f["id"])
                    matriz.append({
                        "id_db": f["id"],
                        "Codigo Consorcio": f["consorcio_codigo"],
                        "Consorcio": f.get("consorcio", formato_consorcio(f["consorcio_codigo"])),
                        "Servicio": f["proveedor"],
                        "Estado": f["estado"],
                        "Monto": float(f["monto"]),
                        "Factura": f["numero_factura"],
                        "Texto Detectado": limpiar_str(f.get("texto_detectado", "")),
                        "Nombre Leido": limpiar_str(f.get("nombre_leido", "")),
                    })
        else:
            matriz.append({
                "id_db": None,
                "Codigo Consorcio": codigo,
                "Consorcio": formato_consorcio(codigo),
                "Servicio": servicio if servicio else "SIN SERVICIO",
                "Estado": "Sin factura cargada",
                "Monto": 0.0,
                "Factura": "-",
                "Texto Detectado": "",
                "Nombre Leido": "",
            })

    pendientes_extra = db_facturas[~db_facturas["id"].isin(ids_procesados)].copy()
    for _, f in pendientes_extra.iterrows():
        matriz.append({
            "id_db": f["id"],
            "Codigo Consorcio": f.get("consorcio_codigo", "SIN IDENTIFICAR"),
            "Consorcio": f.get("consorcio", formato_consorcio(f.get("consorcio_codigo", "SIN IDENTIFICAR"))),
            "Servicio": f["proveedor"],
            "Estado": f["estado"],
            "Monto": float(f["monto"]),
            "Factura": f["numero_factura"],
            "Texto Detectado": limpiar_str(f.get("texto_detectado", "")),
            "Nombre Leido": limpiar_str(f.get("nombre_leido", "")),
        })

    return pd.DataFrame(matriz)

# =========================================================
# HELPER DE GENERACIÓN DE ARCHIVO ZIP EN MEMORIA
# =========================================================
def generar_zip_facturas(df_editado, datos_originales):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for idx, row in df_editado.iterrows():
            nombre_archivo = row["Nuevo Nombre"]
            binario_pdf = datos_originales[idx]["binario"]
            zip_file.writestr(nombre_archivo, binario_pdf)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# =========================================================
# UI PRINCIPAL
# =========================================================
st.sidebar.title("🏢 Administración")

mes_actual = datetime.datetime.now()
opciones_meses = [(mes_actual - datetime.timedelta(days=30 * i)).strftime("%Y-%m") for i in range(12)]
periodo_trabajo = st.sidebar.selectbox("📅 Mes de Trabajo", opciones_meses)

menu = st.sidebar.radio("Navegación:", [
    "📥 1. Ingresar Facturas",
    "📑 2. Planilla de Pagos",
    "⚙️ 3. Limpieza de Base de Datos"
])

# =========================================================
# MENU 1 - INGRESO Y RENOMBRADO DINÁMICO EN VIVO
# =========================================================
if menu == "📥 1. Ingresar Facturas":
    st.title("📂 Procesador de Facturas")
    st.info(f"📌 Las facturas se guardarán en el mes: **{periodo_trabajo}**")

    mapa_codigo_nombre, mapa_nombre_codigo, mapa_alias_codigo, mapa_servicios = construir_mapas_desde_sqlite()

    if "datos_procesados" not in st.session_state:
        st.session_state.datos_procesados = None
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    if client is None:
        st.warning("⚠️ Falta configurar OPENAI_API_KEY. No se podrá procesar con IA.")

    archivos = st.file_uploader(
        "Subí tus facturas PDF",
        type="pdf",
        accept_multiple_files=True,
        key=str(st.session_state.uploader_key)
    )

    df_editado = None

    if archivos and client is not None:
        modo = st.radio("⚙️ Motor:", ("📄 Modo Texto", "👀 Modo Visión"))
        colA, colB = st.columns(2)

        with colA:
            if st.button("🚀 Procesar", use_container_width=True):
                resultados = []
                conteo = {}
                bar = st.progress(0)
                contexto_ia = obtener_contexto_ia(df_cons, df_prov)

                for idx, arch in enumerate(archivos):
                    bytes_pdf = arch.getvalue()
                    info = {}

                    if "Modo Texto" in modo:
                        txt = ""
                        try:
                            with pdfplumber.open(arch) as pdf:
                                for p in pdf.pages[:2]:
                                    txt += p.extract_text() or ""
                        except Exception:
                            txt = ""
                        info = extraer_datos_ia_texto(txt, contexto_ia)
                    else:
                        info = extraer_datos_ia_vision(bytes_pdf, contexto_ia)

                    if "error" not in info:
                        nombre_leido = limpiar_str(info.get("nombre_consorcio_leido", ""))
                        direccion_leida = limpiar_str(info.get("direccion_consorcio_leida", ""))
                        cuit_leido = normalizar_cuit(info.get("cuit_consorcio", ""))
                        cuenta_leida = normalizar_cuit(info.get("nro_cuenta_cliente", ""))

                        candidatos = obtener_candidatos_consorcio(info, df_alias)

                        CONFIDENCE_THRESHOLD = 80
                        
                        if candidatos:
                            mejor = candidatos[0]
                            sugerido_consorcio = mejor["codigo"]
                            score_match = mejor["score"]
                            motivo_match = mejor["motivo"]
                            
                            if mejor["score"] >= CONFIDENCE_THRESHOLD:
                                codigo_cons = mejor["codigo"]
                            else:
                                codigo_cons = "SIN IDENTIFICAR"
                        else:
                            codigo_cons = "SIN IDENTIFICAR"
                            score_match = 0
                            motivo_match = ""
                            sugerido_consorcio = "SIN IDENTIFICAR"

                        cons_display = formato_consorcio(codigo_cons)
                        prov_ia = limpiar_str(info.get("proveedor", "")).upper()
                        alias_prov = resolver_proveedor_alias(prov_ia)

                        texto_detectado = " | ".join([
                            nombre_leido,
                            direccion_leida,
                            cuit_leido,
                            cuenta_leida
                        ]).strip(" |")

                        nb = f"{str(info.get('año_emision', '0000'))}-{str(info.get('mes_emision', '00'))}-{codigo_cons}-{alias_prov}".upper()
                        for char in r'<>:"/\|?*':
                            nb = nb.replace(char, "")

                        if nb in conteo:
                            conteo[nb] += 1
                            nuevo_nombre = f"{nb}-{conteo[nb]}.pdf"
                        else:
                            conteo[nb] = 1
                            nuevo_nombre = f"{nb}.pdf"

                        try:
                            monto_limpio = float(
                                str(info.get("monto_a_pagar", "0"))
                                .replace("$", "")
                                .replace(".", "")
                                .replace(",", ".")
                                .strip()
                            )
                        except:
                            monto_limpio = 0.0

                        resultados.append({
                            "Codigo Consorcio": codigo_cons,
                            "Consorcio": cons_display,
                            "Sugerencia Consorcio": sugerido_consorcio,
                            "Score": score_match,
                            "Motivo": motivo_match,
                            "Nombre Leido": nombre_leido,
                            "Direccion Leida": direccion_leida,
                            "CUIT Leido": cuit_leido,
                            "Cuenta Leida": cuenta_leida,
                            "Texto Detectado": texto_detectado,
                            "Proveedor": alias_prov,
                            "Monto": monto_limpio,
                            "Emisión": info.get("fecha_emision"),
                            "Vencimiento": info.get("vencimiento_completo"),
                            "Factura": info.get("numero_factura"),
                            "Nuevo Nombre": nuevo_nombre,
                            "binario": bytes_pdf
                        })

                    bar.progress((idx + 1) / len(archivos))

                st.session_state.datos_procesados = resultados
                st.session_state.df_trabajo = pd.DataFrame(resultados) if resultados else None
                st.rerun()

        with colB:
            if st.button("🗑️ Borrar Tanda", use_container_width=True):
                st.session_state.uploader_key += 1
                st.session_state.datos_procesados = None
                st.session_state.df_trabajo = None
                st.rerun()

    if st.session_state.datos_procesados and st.session_state.get("df_trabajo") is not None:
        st.warning("⚠️ Los nombres en la columna 'Nombre Planificado' se actualizan en vivo al cambiar los parámetros de la tabla.")
        
        conn = sqlite3.connect(DB_FILE)
        df_cons_dropdown = pd.read_sql_query("SELECT codigo FROM consorcios WHERE activo = 1", conn)
        conn.close()

        lista_consorcios_validos = ["SIN IDENTIFICAR"] + [
            formato_consorcio(c)
            for c in df_cons_dropdown["codigo"].dropna().astype(str).unique().tolist()
        ]

        df_anterior = st.session_state.df_trabajo.copy()

        df_editado = st.data_editor(
            df_anterior,
            column_config={
                "Consorcio": st.column_config.SelectboxColumn("Consorcio", options=lista_consorcios_validos, required=True),
                "Monto": st.column_config.NumberColumn("Monto", format="$%.2f"),
                "Score": st.column_config.NumberColumn("Score", disabled=True),
                "Motivo": st.column_config.TextColumn("Motivo", disabled=True),
                "Sugerencia Consorcio": st.column_config.TextColumn("Sugerencia Consorcio", disabled=True),
                "Nombre Leido": st.column_config.TextColumn("Nombre Leido", disabled=True),
                "Direccion Leida": st.column_config.TextColumn("Direccion Leida", disabled=True),
                "CUIT Leido": st.column_config.TextColumn("CUIT Leido", disabled=True),
                "Cuenta Leida": st.column_config.TextColumn("Cuenta Leida", disabled=True),
                "Texto Detectado": st.column_config.TextColumn("Texto Detectado", disabled=True),
                "Nuevo Nombre": st.column_config.TextColumn("Nombre Planificado", disabled=True),
                "Codigo Consorcio": st.column_config.TextColumn("Codigo Inicial", disabled=True),
            },
            use_container_width=True,
            hide_index=True,
            key="grilla_editor_facturas"
        )

        hubo_cambio_nombre = False
        for idx, row in df_editado.iterrows():
            codigo_sel = parse_codigo_desde_display(row["Consorcio"])
            
            fecha_emi_str = limpiar_str(row["Emisión"])
            anio_emi = "0000"
            mes_emi = "00"
            match_fecha = re.search(r'(\d{2})[-/](\d{2})[-/](\d{4})', fecha_emi_str)
            if match_fecha:
                mes_emi = match_fecha.group(2)
                anio_emi = match_fecha.group(3)
            else:
                match_fecha_alt = re.search(r'(\d{4})[-/](\d{2})[-/](\d{2})', fecha_emi_str)
                if match_fecha_alt:
                    anio_emi = match_fecha_alt.group(1)
                    mes_emi = match_fecha_alt.group(2)

            prov_clean = normalizar_texto(row["Proveedor"])
            nombre_archivo_recalculado = f"{anio_emi}-{mes_emi}-{codigo_sel}-{prov_clean}".upper()
            nombre_archivo_recalculado = re.sub(r'[^A-Z0-9.-]', '', nombre_archivo_recalculado)
            nombre_archivo_recalculado = re.sub(r'\.+', '.', nombre_archivo_recalculado)
            if not nombre_archivo_recalculado.endswith(".pdf"):
                nombre_archivo_recalculado += ".pdf"

            if row["Nuevo Nombre"] != nombre_archivo_recalculado:
                df_editado.at[idx, "Nuevo Nombre"] = nombre_archivo_recalculado
                hubo_cambio_nombre = True

        if hubo_cambio_nombre:
            st.session_state.df_trabajo = df_editado
            st.rerun()

        # DOS COLUMNAS AL PIE DE LA TABLA: GUARDAR EN SISTEMA | DESCARGAR ZIP
        col_guardado, col_descarga = st.columns(2)

        with col_guardado:
            btn_guardar_db = st.button("💾 Guardar en Sistema", type="primary", use_container_width=True)
            
        with col_descarga:
            try:
                zip_data_bytes = generar_zip_facturas(df_editado, st.session_state.datos_procesados)
                st.download_button(
                    label="📦 Descargar Tanda Renombrada (.ZIP)",
                    data=zip_data_bytes,
                    file_name=f"Tanda_Facturas_{periodo_trabajo}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error al generar archivo de descarga ZIP: {e}")

        if btn_guardar_db:
            os.makedirs("facturas_guardadas", exist_ok=True)

            conn = sqlite3.connect(DB_FILE, timeout=30)
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA busy_timeout = 30000;")
            c = conn.cursor()

            try:
                for i, row in df_editado.iterrows():
                    codigo_sel = parse_codigo_desde_display(row["Consorcio"])
                    cons_sel = formato_consorcio(codigo_sel) if codigo_sel != "SIN IDENTIFICAR" else "SIN IDENTIFICAR"
                    nombre_archivo_final = row["Nuevo Nombre"]

                    with open(os.path.join("facturas_guardadas", nombre_archivo_final), "wb") as f:
                        f.write(st.session_state.datos_procesados[i]["binario"])

                    c.execute("""
                        INSERT INTO facturas
                        (consorcio_codigo, consorcio, proveedor, monto, fecha_emision, fecha_vencimiento, numero_factura,
                         nombre_archivo, estado, periodo, texto_detectado, nombre_leido, direccion_leida, cuit_leido, cuenta_leida,
                         score_match, motivo_match, consorcio_sugerido)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        codigo_sel,
                        cons_sel,
                        row["Proveedor"],
                        row["Monto"],
                        row["Emisión"],
                        row["Vencimiento"],
                        row["Factura"],
                        nombre_archivo_final,
                        "Pendiente",
                        periodo_trabajo,
                        row.get("Texto Detectado", ""),
                        row.get("Nombre Leido", ""),
                        row.get("Direccion Leida", ""),
                        row.get("CUIT Leido", ""),
                        row.get("Cuenta Leida", ""),
                        row.get("Score", 0),
                        row.get("Motivo", ""),
                        row.get("Sugerencia Consorcio", "")
                    ))

                conn.commit()

                for i, row in df_editado.iterrows():
                    codigo_sel = parse_codigo_desde_display(row["Consorcio"])
                    if codigo_sel != "SIN IDENTIFICAR":
                        nombre_leido = row.get("Nombre Leido", "")
                        if nombre_leido:
                            guardar_alias_detectado(nombre_leido, codigo_sel, tipo="OCR", prioridad=8, conn=conn)

                conn.commit()
                st.success("✅ Guardado exitoso con nombres corregidos de archivos.")
                st.session_state.datos_procesados = None
                st.session_state.df_trabajo = None
                time.sleep(1.5)
                st.rerun()

            except Exception as e:
                conn.rollback()
                st.error(f"❌ Error al guardar: {e}")
            finally:
                conn.close()

# =========================================================
# MENU 2 - PLANILLA DE PAGOS Y CONTROL
# =========================================================
elif menu == "📑 2. Planilla de Pagos":
    st.title(f"📑 Planilla de Pagos - {periodo_trabajo}")

    mapa_codigo_nombre, mapa_nombre_codigo, mapa_alias_codigo, mapa_servicios = construir_mapas_desde_sqlite()

    conn = sqlite3.connect(DB_FILE)
    df_cons_db = pd.read_sql_query("SELECT * FROM consorcios WHERE activo = 1", conn)
    conn.close()

    if df_cons_db.empty:
        st.error("⚠️ La base de datos de Consorcios está vacía. Añada consorcios en la pestaña de Gestión de Maestros.")
    else:
        df_matriz = generar_matriz_control(periodo_trabajo)

        tab_consorcios, tab_proveedores, tab_maestros, tab_rollover = st.tabs([
            "🏢 Por Consorcios", 
            "🚚 Por Proveedores",
            "⚙️ Gestión de Maestros",
            "🔄 Rollover de Mes"
        ])

        with tab_consorcios:
            if df_matriz.empty:
                st.success("No hay datos en el período actual.")
            else:
                consorcios_pendientes = []
                consorcios_aldia = []
                consorcios_revisar = []
                consorcios_sin_id = []

                lista_consorcios = df_matriz["Codigo Consorcio"].dropna().astype(str).unique().tolist()

                for c in lista_consorcios:
                    if c == "SIN IDENTIFICAR":
                        consorcios_sin_id.append(c)
                        continue

                    df_c = df_matriz[df_matriz["Codigo Consorcio"] == c]
                    estados = df_c["Estado"].tolist()

                    if "A Revisar" in estados:
                        consorcios_revisar.append(c)
                    elif any(e in ["Pendiente", "Falta Factura", "Sin factura cargada"] for e in estados):
                        consorcios_pendientes.append(c)
                    else:
                        consorcios_aldia.append(c)

                tab_pend, tab_aldia, tab_rev, tab_sin = st.tabs([
                    f"⏳ Pendientes ({len(consorcios_pendientes)})",
                    f"✅ Al Día ({len(consorcios_aldia)})",
                    f"👀 A Revisar ({len(consorcios_revisar)})",
                    f"❓ Sin Identificar ({len(df_matriz[df_matriz['Codigo Consorcio'] == 'SIN IDENTIFICAR'])})"
                ])

                opciones_estado = ["Pendiente", "Falta Factura", "Sin factura cargada", "Pagada", "No hay plata", "No mandó", "A Revisar"]

                def render_consorcios(lista_c):
                    if not lista_c:
                        st.success("No hay consorcios en esta categoría.")
                        return

                    for codigo in lista_c:
                        with st.expander(f"🏢 {formato_consorcio(codigo)}"):
                            df_c = df_matriz[df_matriz["Codigo Consorcio"] == codigo].copy()

                            df_edit = st.data_editor(
                                df_c,
                                column_config={
                                    "Estado": st.column_config.SelectboxColumn("Estado Actual", options=opciones_estado, required=True),
                                    "Monto": st.column_config.NumberColumn("Monto ($)", format="$%.2f"),
                                    "id_db": None,
                                    "Codigo Consorcio": None,
                                    "Consorcio": None,
                                    "Texto Detectado": st.column_config.TextColumn("Texto Detectado", disabled=True),
                                    "Nombre Leido": st.column_config.TextColumn("Nombre Leido", disabled=True),
                                },
                                disabled=["Servicio", "Factura"],
                                hide_index=True,
                                use_container_width=True,
                                key=f"editor_{codigo}"
                            )

                            if st.button("💾 Guardar Cambios", key=f"btn_{codigo}", type="primary"):
                                guardar_cambios_consorcio(df_edit, codigo, periodo_trabajo)
                                st.success(f"✅ Se actualizaron los datos de {formato_consorcio(codigo)}.")
                                time.sleep(0.5)
                                st.rerun()

                with tab_pend:
                    render_consorcios(consorcios_pendientes)

                with tab_aldia:
                    render_consorcios(consorcios_aldia)

                with tab_rev:
                    render_consorcios(consorcios_revisar)

                with tab_sin:
                    df_sin = df_matriz[df_matriz["Codigo Consorcio"] == "SIN IDENTIFICAR"].copy()

                    if df_sin.empty:
                        st.success("No hay facturas sin identificar.")
                    else:
                        st.error("🚨 Las siguientes facturas no pudieron asignarse a ningún consorcio. Elegí el correcto y guardá.")
                        
                        lista_validos = ["SIN IDENTIFICAR"] + [
                            formato_consorcio(c)
                            for c in df_cons_db["codigo"].unique().tolist()
                        ]

                        df_sin_edit = st.data_editor(
                            df_sin,
                            column_config={
                                "Consorcio": st.column_config.SelectboxColumn("Mover a Consorcio:", options=lista_validos, required=True),
                                "Estado": st.column_config.TextColumn("Estado", disabled=True),
                                "Monto": st.column_config.NumberColumn("Monto", disabled=True),
                                "id_db": st.column_config.TextColumn("id_db", disabled=True),
                                "Codigo Consorcio": st.column_config.TextColumn("Codigo Consorcio", disabled=True),
                            },
                            disabled=["Servicio", "Factura", "Estado", "Monto", "Texto Detectado", "Nombre Leido", "Codigo Consorcio", "id_db"],
                            hide_index=True,
                            use_container_width=True,
                            key="editor_sin_id"
                        )

                        if st.button("🔄 Reasignar Facturas", type="primary"):
                            reasignar_facturas_sin_identificar(df_sin_edit)
                            st.success("✅ Facturas reasignadas.")
                            time.sleep(1)
                            st.rerun()

        with tab_proveedores:
            st.subheader("Estado de Cuenta por Proveedor")

            conn = sqlite3.connect(DB_FILE)
            df_prov_list = pd.read_sql_query("SELECT * FROM proveedores", conn)
            conn.close()

            if df_prov_list.empty:
                st.warning("No hay proveedores cargados en la base de datos.")
            else:
                lista_proveedores = []
                for _, r in df_prov_list.iterrows():
                    lista_proveedores.append(r["alias_prov"] if r["alias_prov"] else r["nombre_real"])

                for p in sorted(list(set(filter(None, lista_proveedores)))):
                    prov_norm = normalizar_texto(p)
                    alias_norm = normalizar_texto(resolver_proveedor_alias(p))

                    def match_proveedor(x):
                        x_norm = normalizar_texto(x)
                        return (
                            prov_norm in x_norm or x_norm in prov_norm or
                            alias_norm in x_norm or x_norm in alias_norm
                        )

                    df_p = df_matriz[df_matriz["Servicio"].astype(str).apply(match_proveedor)].copy()

                    with st.expander(f"🚚 {p}"):
                        if df_p.empty:
                            st.write("Sin facturas cargadas para este proveedor.")
                        else:
                            df_p_pend = df_p[(df_p["Estado"] == "Pendiente") & (df_p["id_db"].notna())].copy()
                            total = df_p_pend["Monto"].sum() if not df_p_pend.empty else 0.0

                            st.markdown(f"**Total a transferir: ${total:,.2f}**")

                            st.dataframe(
                                df_p[["Consorcio", "Servicio", "Estado", "Factura", "Monto"]],
                                hide_index=True,
                                use_container_width=True
                            )

                            if not df_p_pend.empty:
                                buf = BytesIO()
                                df_p_pend[["Consorcio", "Factura", "Monto"]].to_excel(buf, index=False)
                                st.download_button(
                                    "📥 Descargar Detalle Excel",
                                    buf.getvalue(),
                                    f"Pago_{normalizar_texto(p)}_{periodo_trabajo}.xlsx",
                                    key=f"dl_{normalizar_texto(p)}"
                                )

        # ---------------------------------------------------------
        # SUBTAB 3: Gestión de Maestros (Consorcios / Servicios / Proveedores)
        # ---------------------------------------------------------
        with tab_maestros:
            col_izq, col_med, col_der = st.columns(3)

            # --- COLUMNA 1: CONSORCIOS ---
            with col_izq:
                st.subheader("🏢 Gestión de Consorcios")
                
                with st.form("nuevo_consorcio"):
                    st.markdown("**Agregar Nuevo Consorcio**")
                    nuevo_cod = st.text_input("Código de Consorcio (Ej: C321):").strip().upper()
                    nuevo_nom = st.text_input("Nombre Canónico (Ej: Av Santa Fe 1234):").strip()
                    nuevo_cuit = st.text_input("CUIT (Ej: 30123456789):").strip()
                    nuevo_dom = st.text_input("Domicilio:").strip()
                    
                    btn_enviar = st.form_submit_button("Añadir Consorcio")
                    if btn_enviar:
                        if not nuevo_cod or not nuevo_nom:
                            st.error("Código y Nombre son campos obligatorios.")
                        else:
                            conn = sqlite3.connect(DB_FILE)
                            cursor = conn.cursor()
                            try:
                                cursor.execute("""
                                    INSERT INTO consorcios (codigo, nombre_canonico, cuit, domicilio, activo)
                                    VALUES (?, ?, ?, ?, 1)
                                """, (normalizar_codigo(nuevo_cod), nuevo_nom, normalizar_cuit(nuevo_cuit), nuevo_dom))
                                conn.commit()
                                st.success(f"Consorcio {nuevo_cod} agregado.")
                                time.sleep(0.5)
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("El Código de Consorcio ya existe.")
                            finally:
                                conn.close()

                st.markdown("---")
                st.markdown("**Eliminar / Dar de Baja Consorcio**")
                
                conn = sqlite3.connect(DB_FILE)
                todos_c = pd.read_sql_query("SELECT codigo, nombre_canonico FROM consorcios", conn)
                conn.close()

                if not todos_c.empty:
                    c_eliminar = st.selectbox(
                        "Seleccioná el consorcio a eliminar:",
                        todos_c["codigo"],
                        format_func=lambda x: f"{x} - {todos_c[todos_c['codigo'] == x]['nombre_canonico'].values[0]}",
                        key="select_consorcio_eliminar"
                    )
                    
                    if st.button("❌ Eliminar Consorcio Seleccionado", type="primary", use_container_width=True):
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM consorcios WHERE codigo = ?", (c_eliminar,))
                        cursor.execute("DELETE FROM servicios WHERE consorcio_codigo = ?", (c_eliminar,))
                        conn.commit()
                        conn.close()
                        st.success("Consorcio eliminado.")
                        time.sleep(0.5)
                        st.rerun()

            # --- COLUMNA 2: SERVICIOS ---
            with col_med:
                st.subheader("🔌 Gestión de Servicios")
                
                with st.form("nuevo_servicio"):
                    st.markdown("**Vincular Nuevo Servicio a Consorcio**")
                    
                    conn = sqlite3.connect(DB_FILE)
                    consorcios_activos = pd.read_sql_query("SELECT codigo, nombre_canonico FROM consorcios WHERE activo = 1", conn)
                    conn.close()

                    consorcio_sel = st.selectbox(
                        "Consorcio:",
                        consorcios_activos["codigo"],
                        format_func=lambda x: f"{x} - {consorcios_activos[consorcios_activos['codigo'] == x]['nombre_canonico'].values[0]}"
                    ) if not consorcios_activos.empty else None

                    nuevo_serv = st.text_input("Nombre de Servicio (Ej: METROGAS, AYSA):").strip().upper()
                    nuevo_nro_cli = st.text_input("Número de Cuenta/Cliente:").strip()

                    btn_serv_enviar = st.form_submit_button("Vincular Servicio")
                    if btn_serv_enviar:
                        if not consorcio_sel or not nuevo_serv:
                            st.error("Falta seleccionar el consorcio o completar el servicio.")
                        else:
                            conn = sqlite3.connect(DB_FILE)
                            cursor = conn.cursor()
                            cursor.execute("""
                                INSERT INTO servicios (consorcio_codigo, servicio, nro_cliente)
                                VALUES (?, ?, ?)
                            """, (consorcio_sel, nuevo_serv, nuevo_nro_cli))
                            conn.commit()
                            conn.close()
                            st.success(f"Servicio {nuevo_serv} añadido correctamente.")
                            time.sleep(0.5)
                            st.rerun()

                st.markdown("---")
                st.markdown("**Eliminar Servicios Existentes**")
                
                conn = sqlite3.connect(DB_FILE)
                servicios_actuales = pd.read_sql_query("""
                    SELECT s.id, s.consorcio_codigo, c.nombre_canonico, s.servicio, s.nro_cliente 
                    FROM servicios s
                    JOIN consorcios c ON s.consorcio_codigo = c.codigo
                """, conn)
                conn.close()

                if not servicios_actuales.empty:
                    servicio_eliminar_id = st.selectbox(
                        "Seleccioná el servicio a eliminar:",
                        servicios_actuales["id"],
                        format_func=lambda x: (
                            f"{servicios_actuales[servicios_actuales['id'] == x]['consorcio_codigo'].values[0]} | "
                            f"{servicios_actuales[servicios_actuales['id'] == x]['servicio'].values[0]} "
                            f"({servicios_actuales[servicios_actuales['id'] == x]['nro_cliente'].values[0]})"
                        )
                    )

                    if st.button("❌ Eliminar Servicio Seleccionado", type="primary", use_container_width=True):
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM servicios WHERE id = ?", (int(servicio_eliminar_id),))
                        conn.commit()
                        conn.close()
                        st.success("Servicio eliminado de la grilla de control.")
                        time.sleep(0.5)
                        st.rerun()
                else:
                    st.info("No hay servicios cargados.")

            # --- COLUMNA 3: PROVEEDORES ---
            with col_der:
                st.subheader("🚚 Gestión de Proveedores")
                
                with st.form("nuevo_proveedor"):
                    st.markdown("**Agregar Nuevo Proveedor**")
                    nuevo_prov_real = st.text_input("Razón Social (Ej: AYSA S.A.):").strip()
                    nuevo_prov_alias = st.text_input("Alias Proveedor (Ej: AYSA):").strip().upper()
                    nuevo_prov_key = st.text_input("Palabras Clave (Separar por ';' Ej: AYSA;AGUAS):").strip().upper()
                    
                    btn_prov_enviar = st.form_submit_button("Añadir Proveedor")
                    if btn_prov_enviar:
                        if not nuevo_prov_real:
                            st.error("La Razón Social es obligatoria.")
                        else:
                            conn = sqlite3.connect(DB_FILE)
                            cursor = conn.cursor()
                            try:
                                cursor.execute("""
                                    INSERT INTO proveedores (nombre_real, alias_prov, palabras_clave)
                                    VALUES (?, ?, ?)
                                """, (nuevo_prov_real, nuevo_prov_alias, nuevo_prov_key))
                                conn.commit()
                                st.success(f"Proveedor '{nuevo_prov_real}' añadido.")
                                time.sleep(0.5)
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("Este proveedor ya existe en el sistema.")
                            finally:
                                conn.close()

                st.markdown("---")
                st.markdown("**Eliminar Proveedor Existente**")
                
                conn = sqlite3.connect(DB_FILE)
                provs_sistema = pd.read_sql_query("SELECT id, nombre_real, alias_prov FROM proveedores", conn)
                conn.close()

                if not provs_sistema.empty:
                    prov_eliminar_id = st.selectbox(
                        "Seleccioná el proveedor a eliminar:",
                        provs_sistema["id"],
                        format_func=lambda x: (
                            f"{provs_sistema[provs_sistema['id'] == x]['nombre_real'].values[0]}"
                            f" ({provs_sistema[provs_sistema['id'] == x]['alias_prov'].values[0]})"
                        ),
                        key="select_proveedor_eliminar"
                    )

                    if st.button("❌ Eliminar Proveedor Seleccionado", type="primary", use_container_width=True):
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM proveedores WHERE id = ?", (int(prov_eliminar_id),))
                        conn.commit()
                        conn.close()
                        st.success("Proveedor eliminado.")
                        time.sleep(0.5)
                        st.rerun()
                else:
                    st.info("No hay proveedores registrados.")

        with tab_rollover:
            st.subheader("🔄 Rollover de Período")
            st.info("Esta herramienta permite trasladar los saldos y facturas no resueltos (Pendientes, En Revisión, etc.) de un mes al período posterior.")

            try:
                actual_dt = datetime.datetime.strptime(periodo_trabajo, "%Y-%m")
                siguiente_mes_dt = (actual_dt + datetime.timedelta(days=32)).replace(day=1)
                siguiente_periodo_propuesto = siguiente_mes_dt.strftime("%Y-%m")
            except Exception:
                siguiente_periodo_propuesto = ""

            periodo_destino = st.selectbox("Mes Destino del Rollover:", opciones_meses, index=max(0, opciones_meses.index(siguiente_periodo_propuesto) if siguiente_periodo_propuesto in opciones_meses else 0))

            if periodo_destino == periodo_trabajo:
                st.error("El período de destino debe ser diferente al período de trabajo actual.")
            else:
                conn = sqlite3.connect(DB_FILE)
                pendientes_recuento = pd.read_sql_query("""
                    SELECT COUNT(*) as q FROM facturas 
                    WHERE periodo = ? AND estado IN ('Pendiente', 'Falta Factura', 'A Revisar', 'No hay plata')
                """, conn, params=(periodo_trabajo,)).iloc[0]["q"]
                conn.close()

                st.markdown(f"Se detectaron **{pendientes_recuento}** ítems sin pagar o pendientes en **{periodo_trabajo}**.")

                if pendientes_recuento > 0:
                    if st.button("🚀 Ejecutar Rollover (Traspasar Pendientes)", type="primary", use_container_width=True):
                        conn = sqlite3.connect(DB_FILE)
                        cursor = conn.cursor()
                        try:
                            cursor.execute("""
                                SELECT consorcio_codigo, consorcio, proveedor, monto, fecha_emision, fecha_vencimiento,
                                       numero_factura, nombre_archivo, estado, texto_detectado, nombre_leido,
                                       direccion_leida, cuit_leido, cuenta_leida, score_match, motivo_match, consorcio_sugerido
                                FROM facturas
                                WHERE periodo = ? AND estado IN ('Pendiente', 'Falta Factura', 'A Revisar', 'No hay plata')
                            """, (periodo_trabajo,))
                            
                            items = cursor.fetchall()
                            
                            for item in items:
                                cursor.execute("""
                                    INSERT INTO facturas (
                                        consorcio_codigo, consorcio, proveedor, monto, fecha_emision, fecha_vencimiento,
                                        numero_factura, nombre_archivo, estado, periodo, texto_detectado, nombre_leido,
                                        direccion_leida, cuit_leido, cuenta_leida, score_match, motivo_match, consorcio_sugerido
                                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """, (
                                    item[0], item[1], item[2], item[3], item[4], item[5],
                                    item[6], item[7], item[8], periodo_destino, item[9], item[10],
                                    item[11], item[12], item[13], item[14], item[15], item[16]
                                ))
                            
                            conn.commit()
                            st.success(f"Rollover completado. Se copiaron {pendientes_recuento} comprobantes a {periodo_destino}.")
                            time.sleep(1)
                            st.rerun()
                        except Exception as ex:
                            conn.rollback()
                            st.error(f"Error durante el traspaso: {ex}")
                        finally:
                            conn.close()
                else:
                    st.success("Todo al día. No hay ítems pendientes de traspaso en este período.")

# =========================================================
# MENU 3 - CONFIGURACIÓN Y LIMPIEZA
# =========================================================
elif menu == "⚙️ 3. Limpieza de Base de Datos":
    st.title("⚙️ Configuración y Limpieza")
    st.info("Mantenimiento general de facturas cargadas en la base de datos.")

    conn = sqlite3.connect(DB_FILE)
    df_db = pd.read_sql_query(
        "SELECT * FROM facturas WHERE periodo = ?",
        conn,
        params=(periodo_trabajo,)
    )
    conn.close()

    if df_db.empty:
        st.success(f"No hay facturas guardadas en el mes {periodo_trabajo}.")
    else:
        st.subheader("Borrar una factura específica")

        df_db["etiqueta"] = (
            "ID: " + df_db["id"].astype(str) +
            " | " + df_db["consorcio"].astype(str) +
            " | " + df_db["proveedor"].astype(str) +
            " | $" + df_db["monto"].astype(str)
        )

        factura_a_borrar = st.selectbox(
            "Seleccioná la factura que querés eliminar:",
            df_db["id"],
            format_func=lambda x: df_db[df_db["id"] == x]["etiqueta"].values[0]
        )

        if st.button("❌ Eliminar Factura Seleccionada", type="primary"):
            eliminar_factura_db(factura_a_borrar)
            st.success("Factura eliminada correctamente de la base de datos.")
            time.sleep(1)
            st.rerun()

        st.divider()
        st.subheader("⚠️ Zona Peligrosa")
        if st.button(f"🗑️ Eliminar TODAS las facturas de {periodo_trabajo}", type="secondary"):
            conn = sqlite3.connect(DB_FILE)
            conn.cursor().execute("DELETE FROM facturas WHERE periodo = ?", (periodo_trabajo,))
            conn.commit()
            conn.close()
            st.success(f"Base de datos reseteada para {periodo_trabajo}.")
            time.sleep(1)
            st.rerun()