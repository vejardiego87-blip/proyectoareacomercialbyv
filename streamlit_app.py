import os
import sqlite3
import tempfile
import subprocess
import shutil
import base64
import time
import re
import gspread
from google.oauth2.service_account import Credentials
from io import StringIO
from datetime import datetime
from zoneinfo import ZoneInfo

import unicodedata
import plotly.express as px
import pandas as pd
import streamlit as st

try:
    import requests
except ImportError:
    requests = None
def agregar_logo_central_tenue(ruta_logo: str):
    if not os.path.exists(ruta_logo):
        return

    with open(ruta_logo, "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()

    st.markdown(
        f"""
        <style>

        .stApp {{
            background-color: white;
        }}

        .stApp::after {{
            content: "";
            position: fixed;
            top: 60%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 400px;
            height: 400px;

            background-image: url("data:image/png;base64,{logo_base64}");
            background-repeat: no-repeat;
            background-position: center;
            background-size: contain;

            opacity: 0.03;
            z-index: 0;
            pointer-events: none;
        }}

        </style>
        """,
        unsafe_allow_html=True
    )

import altair as alt
from docxtpl import DocxTemplate, RichText

try:
    import plotly.graph_objects as go
except ImportError:
    st.error("Falta instalar plotly. Ejecuta: pip install plotly")
    st.stop()

try:
    import pydeck as pdk
except ImportError:
    st.error("Falta instalar pydeck. Ejecuta: pip install pydeck")
    st.stop()

# =========================================================
# CONFIGURACION GENERAL
# =========================================================
st.set_page_config(
    page_title="APP Área Comercial Buses y Vans",
    page_icon="🚌",
    layout="wide"
)

DB_FILE = "cotizaciones.db"
LOGO_FILE = "logo_andes_motor.png"
agregar_logo_central_tenue(LOGO_FILE)

USUARIOS = {
    "dvejar": {"nombre": "Diego Vejar"},
    "ssilva": {"nombre": "Sergio Silva"},
    "acorrea": {"nombre": "Alvaro Correa"},
    "forellana": {"nombre": "Fabian Orellana"},
    "rsepulveda": {"nombre": "Rodrigo Sepulveda"},
}

COTIZANTES = {
    "Diego Vejar": {
        "prefijo": "DV",
        "firma_nombre": "Diego Vejar",
        "firma_cargo": "Subgerente Comercial Buses y Vans (IVECO)",
        "firma_correo": "dvejar@andesmotor.cl",
        "firma_telefono": "981774604",
    },
    "Sergio Silva": {
        "prefijo": "SS",
        "firma_nombre": "Sergio Silva",
        "firma_cargo": "Ejecutivo de ventas Zona Sur Buses y Vans (IVECO)",
        "firma_correo": "sergio.silva@andesmotor.cl",
        "firma_telefono": "",
    },
    "Alvaro Correa": {
        "prefijo": "AC",
        "firma_nombre": "Alvaro Correa",
        "firma_cargo": "Ejecutivo de ventas Zona Norte Buses y Vans (IVECO)",
        "firma_correo": "alvaro.correa@andesmotor.cl",
        "firma_telefono": "",
    },
    "Fabian Orellana": {
        "prefijo": "FO",
        "firma_nombre": "Fabian Orellana",
        "firma_cargo": "Product Manager Senior Buses y Vans (IVECO)",
        "firma_correo": "forellana@andesmotor.cl",
        "firma_telefono": "989356449",
    },
    "Rodrigo Sepulveda": {
        "prefijo": "RS",
        "firma_nombre": "Rodrigo Sepulveda",
        "firma_cargo": "Gerente de Buses y Vans (IVECO)",
        "firma_correo": "rsepulveda@andesmotor.cl",
        "firma_telefono": "979783254",
    },
}

MODELOS = {
    "Foton U9": {
        "tipo": "electrico",
        "template": "plantilla_cotizacion_foton_u9.docx",
        "capacidades": ["231,8 kWh", "255 kWh", "266 kWh"],
    },
    "Foton U10": {
        "tipo": "electrico",
        "template": "plantilla_cotizacion_foton_u10.docx",
        "capacidades": ["266 kWh", "310 kWh"],
    },
    "Foton U12": {
        "tipo": "electrico",
        "template": "plantilla_cotizacion_foton_u12.docx",
        "capacidades": ["347 kWh", "382 kWh"],
    },
    "Foton DU9": {
        "tipo": "diesel",
        "template": "plantilla_cotizacion_foton_du9.docx",
        "capacidades": [],
    },
    "Foton DU10": {
        "tipo": "diesel",
        "template": "plantilla_cotizacion_foton_du10.docx",
        "capacidades": [],
    },
}



# =========================================================
# UTILIDADES
# =========================================================
def ahora_santiago():
    return datetime.now(ZoneInfo("America/Santiago"))


def hoy_santiago():
    return ahora_santiago().date()


def fecha_larga_es(fecha):
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    return f"Santiago, {dias[fecha.weekday()]} {fecha.day} de {meses[fecha.month - 1]} de {fecha.year}"


def fecha_corta(fecha):
    return fecha.strftime("%Y%m%d")


def usd_fmt(valor):
    try:
        return f"USD {valor:,.0f}".replace(",", ".")
    except Exception:
        return "USD 0"



def clp_fmt(valor, decimales=0):
    try:
        if decimales == 0:
            return f"CLP {valor:,.0f}".replace(",", ".")
        txt = f"CLP {valor:,.{decimales}f}"
        txt = txt.replace(",", "X").replace(".", ",").replace("X", ".")
        return txt
    except Exception:
        return "CLP 0"


def texto_a_float_chileno(texto):
    try:
        return float(str(texto).replace(".", "").replace(",", ".").strip())
    except Exception:
        return None


@st.cache_data(ttl=3600, show_spinner=False)
def obtener_dolar_observado_bcch():
    url = "https://si3.bcentral.cl/Siete/ES/Siete/Cuadro/CAP_TIPO_CAMBIO/MN_TIPO_CAMBIO4/DOLAR_OBS_ADO"

    if requests is None:
        return None, "No está instalada la librería requests."

    try:
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Accept-Language": "es-CL,es;q=0.9,en;q=0.8",
        }
        resp = requests.get(url, headers=headers, timeout=20)
        resp.raise_for_status()
        html = resp.text

        # Intento 1: extraer el bloque del cuadro y tomar el último valor publicado
        bloque = re.search(
            r"D[óo]lar observado(.*?)(?:Eliminar canasta|Exportar a Excel|Mi BDE)",
            html,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if bloque:
            valores = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", bloque.group(1))
            if valores:
                valor = texto_a_float_chileno(valores[-1])
                if valor is not None:
                    return valor, "Fuente: Banco Central de Chile"

        # Intento 2: leer tablas HTML si el sitio entrega la tabla renderizada
        try:
            tablas = pd.read_html(StringIO(html))
            for tabla in tablas:
                tabla = tabla.copy()
                for _, fila in tabla.iterrows():
                    fila_txt = " | ".join(fila.astype(str).tolist()).lower()
                    if "dólar observado" in fila_txt or "dolar observado" in fila_txt:
                        for valor_txt in reversed(fila.astype(str).tolist()):
                            valor = texto_a_float_chileno(valor_txt)
                            if valor is not None and valor > 0:
                                return valor, "Fuente: Banco Central de Chile"
        except Exception:
            pass

        return None, "No fue posible identificar el valor publicado en el sitio del Banco Central."

    except Exception as e:
        return None, f"No fue posible consultar el Banco Central: {e}"


def limpiar_nombre_archivo(texto):
    texto = str(texto).strip().replace(" ", "_")
    reemplazos = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "ñ": "n", "Ñ": "N",
        ",": "", ".": "", "/": "-", "\\": "-",
        ":": "-", ";": "", "(": "", ")": "", "[": "", "]": "",
    }
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    return texto


def formatear_texto_mantto(texto):
    rt = RichText()
    bloques = [b.strip() for b in texto.split("\n") if b.strip()]
    for i, bloque in enumerate(bloques):
        rt.add(bloque)
        if i < len(bloques) - 1:
            rt.add("\n\n")
    return rt


def detectar_motor_excel(uploaded_file):
    nombre = uploaded_file.name.lower()
    if nombre.endswith(".xlsx"):
        return "openpyxl"
    if nombre.endswith(".xls"):
        return None
    return "openpyxl"


def leer_excel_hoja(uploaded_file, sheet_name):
    engine = detectar_motor_excel(uploaded_file)
    if engine:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def sugerir_columna(columnas, candidatos):
    columnas_lower = {str(c).lower(): c for c in columnas}
    for cand in candidatos:
        for col_lower, col_real in columnas_lower.items():
            if cand in col_lower:
                return col_real
    return None


def obtener_template_por_modelo(modelo):
    return MODELOS[modelo]["template"]

def normalizar_texto(texto):
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    return texto


def normalizar_columnas(df):
    nuevas = {}
    for c in df.columns:
        c_norm = normalizar_texto(c)
        nuevas[c] = c_norm
    return df.rename(columns=nuevas)
# =========================================================
# GOOGLE SHEETS - HISTORIAL
# =========================================================
GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def get_gsheet_client():
    creds_info = dict(st.secrets["gsheets"]["service_account"])
    if "private_key" in creds_info:
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")

    creds = Credentials.from_service_account_info(
        creds_info,
        scopes=GSHEET_SCOPES
    )
    return gspread.authorize(creds)

@st.cache_resource
def get_gsheet():
    client = get_gsheet_client()
    return client.open_by_key("12iAqv8Gj6a3LRgTg95oJFS5ip1FSVetM2Q62ozuNGs4")

def asegurar_hoja_historial():
    sh = get_gsheet()
    try:
        ws = sh.worksheet("HistorialCotizaciones")
    except Exception:
        ws = sh.add_worksheet(title="HistorialCotizaciones", rows=2000, cols=20)
        ws.append_row([
            "id",
            "fecha",
            "cliente",
            "cotizante",
            "prefijo",
            "correlativo",
            "numero_cotizacion",
            "modelo",
            "capacidad_bateria",
            "cantidad_unidades",
            "precio_unitario",
            "total_negocio",
            "lugar_entrega",
            "contrato_mantto",
            "texto_mantto",
            "creado_en",
        ])
    return ws
# =========================================================
# BASE DE DATOS
# =========================================================
def get_conn():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def table_exists(conn, table_name):
    cur = conn.cursor()
    cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,)
    )
    return cur.fetchone() is not None


def get_columns(conn, table_name):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table_name})")
    return [row[1] for row in cur.fetchall()]


def crear_tabla_cotizaciones(conn):
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS cotizaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            cliente TEXT NOT NULL,
            cotizante TEXT NOT NULL,
            prefijo TEXT NOT NULL,
            correlativo INTEGER NOT NULL,
            numero_cotizacion TEXT NOT NULL UNIQUE,
            modelo TEXT,
            capacidad_bateria TEXT,
            cantidad_unidades INTEGER NOT NULL,
            precio_unitario REAL NOT NULL,
            total_negocio REAL NOT NULL,
            lugar_entrega TEXT,
            contrato_mantto TEXT,
            texto_mantto TEXT,
            creado_en TEXT NOT NULL
        )
    """)
    conn.commit()


def migrar_base_si_corresponde():
    conn = get_conn()

    if not table_exists(conn, "cotizaciones"):
        crear_tabla_cotizaciones(conn)
        conn.close()
        return

    columnas = get_columns(conn, "cotizaciones")
    cur = conn.cursor()

    faltantes = {
        "modelo": "TEXT",
        "capacidad_bateria": "TEXT",
        "lugar_entrega": "TEXT",
        "contrato_mantto": "TEXT",
        "texto_mantto": "TEXT",
        "creado_en": "TEXT",
        "fecha": "TEXT",
        "prefijo": "TEXT",
        "correlativo": "INTEGER",
        "cantidad_unidades": "INTEGER DEFAULT 1",
        "precio_unitario": "REAL DEFAULT 0",
        "total_negocio": "REAL DEFAULT 0",
    }

    if "numero_cotizacion" in columnas:
        for col, tipo in faltantes.items():
            if col not in columnas:
                cur.execute(f"ALTER TABLE cotizaciones ADD COLUMN {col} {tipo}")
        conn.commit()
        conn.close()
        return

    if "numero" in columnas:
        cur.execute("ALTER TABLE cotizaciones RENAME TO cotizaciones_old")
        conn.commit()
        crear_tabla_cotizaciones(conn)

        columnas_old = get_columns(conn, "cotizaciones_old")
        fecha_expr = "fecha" if "fecha" in columnas_old else "''"
        cliente_expr = "cliente" if "cliente" in columnas_old else "''"
        cotizante_expr = "cotizante" if "cotizante" in columnas_old else "''"
        numero_expr = "numero" if "numero" in columnas_old else "''"

        cur.execute(f"""
            INSERT INTO cotizaciones (
                fecha, cliente, cotizante, prefijo, correlativo, numero_cotizacion,
                modelo, capacidad_bateria, cantidad_unidades, precio_unitario, total_negocio,
                lugar_entrega, contrato_mantto, texto_mantto, creado_en
            )
            SELECT
                {fecha_expr},
                {cliente_expr},
                {cotizante_expr},
                '',
                0,
                {numero_expr},
                '',
                '',
                1,
                0,
                0,
                '',
                '',
                '',
                ''
            FROM cotizaciones_old
        """)
        conn.commit()
        conn.close()
        return

    cur.execute("DROP TABLE IF EXISTS cotizaciones")
    conn.commit()
    crear_tabla_cotizaciones(conn)
    conn.close()


def init_db():
    migrar_base_si_corresponde()


def siguiente_correlativo(cotizante):
    prefijo = COTIZANTES[cotizante]["prefijo"]
    ws = asegurar_hoja_historial()
    registros = ws.get_all_records()

    usados = []
    for row in registros:
        try:
            if str(row.get("prefijo", "")).strip() == prefijo:
                usados.append(int(row.get("correlativo", 0)))
        except Exception:
            pass

    correlativo = 1
    while correlativo in usados:
        correlativo += 1
    return correlativo


def guardar_cotizacion(data):
    ws = asegurar_hoja_historial()
    registros = ws.get_all_records()
    creado_en = ahora_santiago().strftime("%Y-%m-%d %H:%M:%S")

    next_id = 1
    if registros:
        ids = []
        for r in registros:
            try:
                ids.append(int(r.get("id", 0)))
            except Exception:
                pass
        if ids:
            next_id = max(ids) + 1

    ws.append_row([
        next_id,
        data["fecha_iso"],
        data["cliente"],
        data["cotizante"],
        data["prefijo"],
        data["correlativo"],
        data["numero_cotizacion"],
        data["modelo"],
        data["capacidad_bateria"],
        data["cantidad_unidades"],
        data["precio_unitario_raw"],
        data["total_negocio_raw"],
        data["lugar_entrega"],
        data["contrato_mantto"],
        data["texto_mantto"],
        creado_en,
    ])


def cargar_historial():
    ws = asegurar_hoja_historial()
    registros = ws.get_all_records()

    if not registros:
        return pd.DataFrame(columns=[
            "id",
            "fecha",
            "cliente",
            "cotizante",
            "numero_cotizacion",
            "modelo",
            "cantidad_unidades",
            "precio_unitario",
            "total_negocio",
            "capacidad_bateria",
            "lugar_entrega",
            "creado_en",
        ])

    df = pd.DataFrame(registros)

    columnas_esperadas = [
        "id",
        "fecha",
        "cliente",
        "cotizante",
        "numero_cotizacion",
        "modelo",
        "cantidad_unidades",
        "precio_unitario",
        "total_negocio",
        "capacidad_bateria",
        "lugar_entrega",
        "creado_en",
    ]

    for c in columnas_esperadas:
        if c not in df.columns:
            df[c] = ""

    return df[columnas_esperadas].sort_values("id", ascending=False)


def eliminar_cotizacion_por_id(cotizacion_id):
    ws = asegurar_hoja_historial()
    values = ws.get_all_values()

    if not values or len(values) < 2:
        return

    header = values[0]
    filas = values[1:]

    nuevas_filas = [header]

    for fila in filas:
        try:
            if int(fila[0]) != int(cotizacion_id):
                nuevas_filas.append(fila)
        except Exception:
            nuevas_filas.append(fila)

    ws.clear()
    ws.update("A1", nuevas_filas)

# =========================================================
# DOCX / PDF
# =========================================================
def generar_docx(contexto, template_file, nombre_salida):
    if not os.path.exists(template_file):
        raise FileNotFoundError(f"No se encontró la plantilla '{template_file}'.")

    doc = DocxTemplate(template_file)
    doc.render(contexto)

    tmp_dir = tempfile.mkdtemp()
    salida = os.path.join(tmp_dir, f"{nombre_salida}.docx")
    doc.save(salida)
    return salida


def convertir_docx_a_pdf(docx_path):
    posibles = [
        shutil.which("soffice"),
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
    ]
    soffice_path = next((p for p in posibles if p and os.path.exists(p)), None)

    if soffice_path is None:
        raise RuntimeError(
            "No se encontró LibreOffice/soffice en el servidor. Instala LibreOffice para habilitar PDF."
        )

    output_dir = os.path.dirname(docx_path)
    comando = [
        soffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ]

    resultado = subprocess.run(comando, capture_output=True, text=True)

    if resultado.returncode != 0:
        raise RuntimeError(
            f"Error al convertir PDF. STDOUT: {resultado.stdout} | STDERR: {resultado.stderr}"
        )

    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    if not os.path.exists(pdf_path):
        raise RuntimeError("No se generó el archivo PDF.")

    return pdf_path

# =========================================================
# LOGIN
# =========================================================
if "usuario" not in st.session_state:
    st.session_state.usuario = None

if st.session_state.usuario is None:
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if os.path.exists(LOGO_FILE):
            st.image(LOGO_FILE, width=180)
        st.title("Ingreso Área Comercial")
        user = st.text_input("Usuario")
        login_btn = st.button("Ingresar", use_container_width=True)

        if login_btn:
            if user in USUARIOS:
                st.session_state.usuario = user
                st.rerun()
            else:
                st.error("Usuario no válido")
    st.stop()

usuario_actual = USUARIOS[st.session_state.usuario]["nombre"]

# =========================================================
# APP
# =========================================================
init_db()

header1, header2 = st.columns([1.2, 6])

with header1:
    if os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=180)

with header2:
    st.markdown(
        """
        <div style="padding-top: 8px;">
            <h1 style='margin-bottom:0;'>APP Área Comercial Buses y Vans</h1>
            <p style='margin-top:0;color:gray;'>Andes Motor - Plataforma Comercial</p>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

st.sidebar.success(f"Usuario: {usuario_actual}")
st.sidebar.caption(f"Hora Santiago: {ahora_santiago().strftime('%d-%m-%Y %H:%M:%S')}")
if st.sidebar.button("Cerrar sesión"):
    st.session_state.usuario = None
    st.rerun()

usuarios_costos = {"rsepulveda", "forellana", "dvejar"}
usuarios_costos_editan = {"forellana", "dvejar"}

if st.session_state.usuario in usuarios_costos:
    tab_cot, tab_hist, tab_efi, tab_dash, tab_cost = st.tabs([
        "🧾 Nueva cotización",
        "📚 Historial",
        "⚡ Eficiencia energética",
        "📊 Dashboard Comercial",
        "💰 Estructura de costo"
    ])
else:
    tab_cot, tab_hist, tab_efi, tab_dash = st.tabs([
        "🧾 Nueva cotización",
        "📚 Historial",
        "⚡ Eficiencia energética",
        "📊 Dashboard Comercial"
    ])

# =========================================================
# TAB 1 - COTIZACION
# =========================================================
with tab_cot:
    st.subheader("Cotizaciones Foton")

    cotizante = usuario_actual
    modelo = st.selectbox("Modelo", list(MODELOS.keys()))
    modelo_info = MODELOS[modelo]

    c1, c2 = st.columns(2)

    with c1:
        fecha = st.date_input("Fecha", value=hoy_santiago())
        cliente = st.text_input("Cliente", value="")
        st.text_input("Cotizante", value=cotizante, disabled=True)
        lugar_entrega = st.text_input("Lugar de entrega", value="")

    with c2:
        cantidad_unidades = st.number_input("Cantidad de unidades", min_value=1, value=1, step=1)
        precio_unitario = st.number_input("Precio unitario USD", min_value=0.0, value=130491.0, step=1000.0)
        contrato_mantto = st.text_input("Contrato mantto", value="48 meses")

        if modelo_info["tipo"] == "electrico":
            capacidad_bateria = st.selectbox("Capacidad nominal batería", modelo_info["capacidades"])
        else:
            capacidad_bateria = ""

    texto = st.text_area(
        "Texto mantenimiento",
        value="""I. La oferta incluye 24 meses de mantenimiento preventivo y correctivo (Correctivo solo desgaste Disco y Pastilla de frenos) sin costo para el cliente, con el fin de entregar conocimientos técnicos y aprendizaje continuo del mantenimiento para este tipo de vehículos, durante este periodo.
II. Adicionalmente se entregarán USD 1.000.- por bus, para la compra de repuestos de desgaste a elección del cliente (este ítem deberá ser utilizado dentro de los primeros 12 meses realizada la entrega de flota).
III. La oferta incluye 4 años de telemetría sin costo para el cliente.""",
        height=180
    )

    prox = siguiente_correlativo(cotizante)
    prefijo = COTIZANTES[cotizante]["prefijo"]
    st.info(f"Próximo correlativo automático: {prefijo}-{prox:02d}")

    if st.button("Generar cotización", use_container_width=True):
        if not cliente.strip():
            st.error("Debes ingresar el nombre del cliente.")
        elif not lugar_entrega.strip():
            st.error("Debes ingresar el lugar de entrega.")
        else:
            correlativo = siguiente_correlativo(cotizante)
            numero = f"{prefijo}-{correlativo:02d}"
            total_negocio = precio_unitario * cantidad_unidades

            nombre_archivo = (
                f"Propuesta_{limpiar_nombre_archivo(modelo)}_"
                f"{limpiar_nombre_archivo(cliente)}_"
                f"{limpiar_nombre_archivo(capacidad_bateria) if capacidad_bateria else 'diesel'}_"
                f"{fecha_corta(fecha)}"
            )

            contexto = {
                "cliente": cliente.strip(),
                "numero_cotizacion": numero,
                "fecha_larga": fecha_larga_es(fecha),
                "fecha_corta": fecha_corta(fecha),
                "cantidad_unidades": cantidad_unidades,
                "precio_unitario": usd_fmt(precio_unitario),
                "capacidad_bateria": capacidad_bateria,
                "texto_mantto": formatear_texto_mantto(texto) if texto.strip() else "",
                "firma_nombre": COTIZANTES[cotizante]["firma_nombre"],
                "firma_cargo": COTIZANTES[cotizante]["firma_cargo"],
                "firma_correo": COTIZANTES[cotizante]["firma_correo"],
                "firma_telefono": COTIZANTES[cotizante]["firma_telefono"],
                "lugar_entrega": lugar_entrega.strip(),
            }

            registro = {
                "fecha_iso": fecha.isoformat(),
                "cliente": cliente.strip(),
                "cotizante": cotizante,
                "prefijo": prefijo,
                "correlativo": correlativo,
                "numero_cotizacion": numero,
                "modelo": modelo,
                "capacidad_bateria": capacidad_bateria,
                "cantidad_unidades": int(cantidad_unidades),
                "precio_unitario_raw": float(precio_unitario),
                "total_negocio_raw": float(total_negocio),
                "lugar_entrega": lugar_entrega.strip(),
                "contrato_mantto": contrato_mantto,
                "texto_mantto": texto,
            }

            try:
                template_file = obtener_template_por_modelo(modelo)
                archivo_docx = generar_docx(contexto, template_file, nombre_archivo)
                guardar_cotizacion(registro)

                with open(archivo_docx, "rb") as f:
                    contenido_docx = f.read()

                st.success(f"Cotización {numero} generada correctamente.")
                st.write(f"**Fecha:** {fecha_larga_es(fecha)}")
                st.write(f"**Cliente:** {cliente}")
                st.write(f"**Cotizante:** {cotizante}")
                st.write(f"**Modelo:** {modelo}")
                st.write(f"**Precio unitario:** {usd_fmt(precio_unitario)}")
                st.write(f"**Total negocio:** {usd_fmt(total_negocio)}")

                d1, d2 = st.columns(2)

                with d1:
                    st.download_button(
                        "Descargar Word",
                        data=contenido_docx,
                        file_name=f"{nombre_archivo}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                try:
                    inicio_pdf = time.time()
                    with st.spinner("Generando PDF..."):
                        archivo_pdf = convertir_docx_a_pdf(archivo_docx)

                    fin_pdf = time.time()
                    segundos_pdf = round(fin_pdf - inicio_pdf, 2)

                    with open(archivo_pdf, "rb") as f:
                        contenido_pdf = f.read()

                    with d2:
                        st.download_button(
                            "Descargar PDF",
                            data=contenido_pdf,
                            file_name=f"{nombre_archivo}.pdf",
                            mime="application/pdf"
                        )

                    st.caption(f"PDF generado en {segundos_pdf} segundos.")

                except Exception as e_pdf:
                    st.warning(f"No fue posible habilitar el PDF: {e_pdf}")

            except Exception as e:
                st.error(f"Error al generar la cotización: {e}")

# =========================================================
# TAB 2 - HISTORIAL / ELIMINAR
# =========================================================
with tab_hist:
    st.subheader("Historial")

    try:
        df_hist = cargar_historial()
        if df_hist.empty:
            st.info("Aún no hay cotizaciones registradas.")
        else:
            vista = df_hist.copy()
            vista["precio_unitario"] = vista["precio_unitario"].apply(usd_fmt)
            vista["total_negocio"] = vista["total_negocio"].apply(usd_fmt)
            st.dataframe(vista, use_container_width=True)

            st.markdown("### Eliminar cotización")
            opciones = [
                f"{row['id']} | {row['numero_cotizacion']} | {row['cliente']} | {row['modelo']}"
                for _, row in df_hist.iterrows()
            ]
            seleccion = st.selectbox("Selecciona una cotización para eliminar", opciones)

            if st.button("Eliminar cotización seleccionada", type="secondary"):
                cotizacion_id = int(seleccion.split("|")[0].strip())
                eliminar_cotizacion_por_id(cotizacion_id)
                st.success("Cotización eliminada. El correlativo queda disponible nuevamente.")
                st.rerun()

    except Exception as e:
        st.error(f"No fue posible cargar el historial: {e}")



# =========================================================
# TAB 3 - EFICIENCIA ENERGÉTICA
# =========================================================
with tab_efi:
    import pandas as pd
    import plotly.graph_objects as go
    import plotly.express as px
    import unicodedata
    from io import BytesIO
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Image
    )
    from reportlab.lib.styles import getSampleStyleSheet
    import matplotlib.pyplot as plt
    from matplotlib.gridspec import GridSpec
    from matplotlib.patches import Wedge, Circle
    from matplotlib import image as mpimg
    import numpy as np

    st.subheader("⚡ Eficiencia energética")

    archivo = st.file_uploader("Subir Excel (.xlsx)", type=["xlsx"], key="efi_tab3")

    def norm(txt):
        if pd.isna(txt):
            return ""
        txt = str(txt).lower().strip()
        return unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("utf-8")

    def buscar_columna(cols, candidatos):
        cols_norm = {norm(c): c for c in cols}
        for cand in candidatos:
            cand_n = norm(cand)
            for c_n, c_real in cols_norm.items():
                if cand_n in c_n:
                    return c_real
        return None

    def dibujar_gauge(ax, valor, vmin, vmax, rangos, titulo, texto_valor, color_barra="#15803d"):
        ax.set_aspect("equal")
        ax.axis("off")

        for inicio, fin, color in rangos:
            th1 = 180 - 180 * ((inicio - vmin) / (vmax - vmin))
            th2 = 180 - 180 * ((fin - vmin) / (vmax - vmin))
            wedge = Wedge((0, 0), 1.0, th1, th2, width=0.22, facecolor=color, edgecolor="none")
            ax.add_patch(wedge)

        if valor is not None:
            valor_clip = max(vmin, min(vmax, valor))
            angle = np.deg2rad(180 - 180 * ((valor_clip - vmin) / (vmax - vmin)))
            x = 0.78 * np.cos(angle)
            y = 0.78 * np.sin(angle)
            ax.plot([0, x], [0, y], color=color_barra, linewidth=9, solid_capstyle="round")

        inner = Circle((0, 0), 0.58, color="white", ec="none")
        ax.add_patch(inner)

        ax.text(0, 0.20, titulo, ha="center", va="center", fontsize=13, fontweight="bold")
        ax.text(0, -0.02, texto_valor, ha="center", va="center", fontsize=17, color="#1f2937")
        ax.set_xlim(-1.15, 1.15)
        ax.set_ylim(-0.15, 1.15)

    def generar_dashboard_png(base, trazado, bateria, distancia, consumo, rendimiento, autonomia_total, autonomia_15, vel_prom, logo_path=None):
        fig = plt.figure(figsize=(16, 9), facecolor="#eef3f8")
        gs = GridSpec(3, 2, figure=fig, height_ratios=[2.0, 1.1, 1.5], width_ratios=[1.1, 1.0], hspace=0.38, wspace=0.18)

        if logo_path and os.path.exists(logo_path):
            try:
                img = mpimg.imread(logo_path)
                ax_w = fig.add_axes([0.23, 0.27, 0.54, 0.46], zorder=0)
                ax_w.imshow(img, alpha=0.07)
                ax_w.axis("off")
            except Exception:
                pass

        fig.text(0.035, 0.95, f"Resultados Pruebas {trazado}", fontsize=20, fontweight="bold", color="#1f2937")

        ax1 = fig.add_subplot(gs[0, 0])
        ax1.set_facecolor("#eef3f8")
        x = base["odometro"].to_numpy()
        yv = base["velocidad"].to_numpy()
        ax1.plot(x, yv, color="#2185f5", linewidth=2.5, marker="o", markersize=3)
        ax1.set_title("Velocidad y Estado de Carga", loc="left", fontsize=13, fontweight="bold")
        ax1.set_xlabel("Odómetro km")
        ax1.set_ylabel("Velocidad km/h")
        ax1.grid(True, linestyle=":", alpha=0.35)

        if base["soc"].notna().any():
            ax1b = ax1.twinx()
            ys = base["soc"].to_numpy()
            ax1b.plot(x, ys, color="#1e3a8a", linewidth=2.5, marker="o", markersize=3)
            ax1b.set_ylabel("Estado de Carga SoC")

        ax2 = fig.add_subplot(gs[0, 1])
        ax2.set_facecolor("#eef3f8")
        mapa = base.dropna(subset=["lat", "lon"]).copy()
        if not mapa.empty:
            ax2.plot(mapa["lon"], mapa["lat"], color="#2563eb", linewidth=2.2)
            ax2.scatter(mapa["lon"], mapa["lat"], s=18, color="#1d4ed8", alpha=0.8)
            ax2.scatter(mapa["lon"].iloc[0], mapa["lat"].iloc[0], s=60, color="#16a34a", label="Inicio", zorder=5)
            ax2.scatter(mapa["lon"].iloc[-1], mapa["lat"].iloc[-1], s=60, color="#dc2626", label="Fin", zorder=5)
            ax2.set_title("Recorrido georreferenciado", loc="left", fontsize=13, fontweight="bold")
            ax2.set_xlabel("Longitud")
            ax2.set_ylabel("Latitud")
            ax2.grid(True, linestyle=":", alpha=0.35)
            ax2.legend(loc="upper right", frameon=False, fontsize=9)
        else:
            ax2.text(0.5, 0.5, "Sin coordenadas válidas", ha="center", va="center", fontsize=12)
            ax2.axis("off")

        gsg = gs[1, :].subgridspec(1, 3, wspace=0.28)
        axg1 = fig.add_subplot(gsg[0, 0])
        dibujar_gauge(axg1, rendimiento, 0, 1.2,
                      [(0, 0.90, "#22c55e"), (0.90, 1.00, "#facc15"), (1.00, 1.20, "#ef4444")],
                      "Rendimiento kWh/km", f"{rendimiento:.3f}")
        axg2 = fig.add_subplot(gsg[0, 1])
        dibujar_gauge(axg2, autonomia_total, 0, 500,
                      [(0, 280, "#ef4444"), (280, 350, "#facc15"), (350, 500, "#22c55e")],
                      f"Autonomía total ({bateria:.0f} kWh)", f"{autonomia_total:.0f} km")
        axg3 = fig.add_subplot(gsg[0, 2])
        dibujar_gauge(axg3, autonomia_15, 0, 500,
                      [(0, 250, "#ef4444"), (250, 320, "#facc15"), (320, 500, "#22c55e")],
                      "Autonomía útil al 15% SoC", f"{autonomia_15:.0f} km")

        ax3 = fig.add_subplot(gs[2, :])
        ax3.set_facecolor("#eef3f8")
        if base["altitud"].notna().any():
            ax3.fill_between(base["odometro"], base["altitud"], color="#60a5fa", alpha=0.55)
            ax3.plot(base["odometro"], base["altitud"], color="#2185f5", linewidth=2.5)
            ax3.set_title("Perfil de altura", loc="left", fontsize=13, fontweight="bold")
            ax3.set_xlabel("Odómetro")
            ax3.set_ylabel("Altura (m)")
            ax3.grid(True, linestyle=":", alpha=0.35)

        fig.text(0.09, 0.315, "Velocidad promedio", fontsize=12, fontweight="bold")
        fig.text(0.12, 0.235, f"{vel_prom:.1f}", fontsize=28, color="#111827")
        fig.text(0.165, 0.24, "km/h", fontsize=11)

        fig.text(0.28, 0.315, "Kms recorridos", fontsize=12, fontweight="bold")
        fig.text(0.295, 0.235, f"{distancia:.1f}", fontsize=28, color="#111827")
        fig.text(0.345, 0.24, "km", fontsize=11)

        fig.text(0.46, 0.315, "Consumo energético", fontsize=12, fontweight="bold")
        fig.text(0.47, 0.235, f"{consumo:.2f}", fontsize=28, color="#111827")
        fig.text(0.535, 0.24, "kWh", fontsize=11)

        fig.text(0.65, 0.315, "Batería HV", fontsize=12, fontweight="bold")
        fig.text(0.675, 0.235, f"{bateria:.0f}", fontsize=28, color="#111827")
        fig.text(0.72, 0.24, "kWh", fontsize=11)

        out = BytesIO()
        fig.savefig(out, format="png", dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
        plt.close(fig)
        out.seek(0)
        return out.getvalue()

    def generar_pdf_ejecutivo(trazado, bateria, distancia, consumo, rendimiento, autonomia_total, autonomia_15, vel_prom, base, logo_path=None):
        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=landscape(A4),
            rightMargin=0.8 * cm,
            leftMargin=0.8 * cm,
            topMargin=0.8 * cm,
            bottomMargin=0.8 * cm
        )

        styles = getSampleStyleSheet()
        story = []

        titulo = Paragraph(f"<b>Informe Eficiencia Energética - {trazado}</b>", styles["Title"])
        story.append(titulo)
        story.append(Spacer(1, 0.15 * cm))

        resumen = Paragraph(
            "Dashboard ejecutivo del recorrido seleccionado, construido con cálculos desde la hoja resumen "
            "y visualizaciones desde la hoja base.",
            styles["BodyText"]
        )
        story.append(resumen)
        story.append(Spacer(1, 0.2 * cm))

        img_dashboard = generar_dashboard_png(
            base=base,
            trazado=trazado,
            bateria=bateria,
            distancia=distancia,
            consumo=consumo,
            rendimiento=rendimiento,
            autonomia_total=autonomia_total,
            autonomia_15=autonomia_15,
            vel_prom=vel_prom,
            logo_path=logo_path
        )

        img = Image(BytesIO(img_dashboard))
        img.drawWidth = 27.8 * cm
        img.drawHeight = 17.3 * cm
        story.append(img)

        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()

    if archivo:
        try:
            xls = pd.ExcelFile(archivo)
            hojas = xls.sheet_names

            if len(hojas) < 2:
                st.error("El archivo debe traer al menos 2 hojas: base y resumen.")
                st.stop()

            df_base = pd.read_excel(archivo, sheet_name=0)
            df_resumen = pd.read_excel(archivo, sheet_name=1)

            df_base_original = df_base.copy()
            df_resumen_original = df_resumen.copy()

            df_base.columns = [norm(c) for c in df_base.columns]
            df_resumen.columns = [norm(c) for c in df_resumen.columns]

            col_trazado_base = buscar_columna(df_base.columns, ["trazado", "ruta"])
            col_odo = buscar_columna(df_base.columns, ["odometro", "odómetro"])
            col_vel = buscar_columna(df_base.columns, ["velocidad"])
            col_soc = buscar_columna(df_base.columns, ["soc", "estado de carga"])
            col_alt = buscar_columna(df_base.columns, ["altitud", "altura"])
            col_lat = buscar_columna(df_base.columns, ["latitud", "latitude", "lat"])
            col_lon = buscar_columna(df_base.columns, ["longitud", "longitude", "long", "lon"])

            col_trazado_res = buscar_columna(df_resumen.columns, ["trazado", "ruta"])
            col_distancia = buscar_columna(df_resumen.columns, ["distancia"])
            col_consumo = buscar_columna(df_resumen.columns, ["consumo energetico", "consumo energético", "consumo"])

            faltantes_base = []
            if col_trazado_base is None: faltantes_base.append("trazado")
            if col_odo is None: faltantes_base.append("odometro")
            if col_vel is None: faltantes_base.append("velocidad")
            if col_soc is None: faltantes_base.append("soc")
            if col_alt is None: faltantes_base.append("altitud")
            if col_lat is None: faltantes_base.append("latitud")
            if col_lon is None: faltantes_base.append("longitud")

            faltantes_res = []
            if col_trazado_res is None: faltantes_res.append("trazado")
            if col_distancia is None: faltantes_res.append("distancia")
            if col_consumo is None: faltantes_res.append("consumo energetico")

            if faltantes_base:
                st.error(f"Faltan columnas en hoja base: {', '.join(faltantes_base)}")
                st.stop()

            if faltantes_res:
                st.error(f"Faltan columnas en hoja resumen: {', '.join(faltantes_res)}")
                st.stop()

            df_base["trazado_key"] = df_base[col_trazado_base].astype(str).apply(norm)
            df_resumen["trazado_key"] = df_resumen[col_trazado_res].astype(str).apply(norm)

            df_base["odometro"] = pd.to_numeric(df_base[col_odo], errors="coerce")
            df_base["velocidad"] = pd.to_numeric(df_base[col_vel], errors="coerce")
            df_base["soc"] = pd.to_numeric(df_base[col_soc], errors="coerce")
            df_base["altitud"] = pd.to_numeric(df_base[col_alt], errors="coerce")
            df_base["lat"] = pd.to_numeric(df_base[col_lat], errors="coerce")
            df_base["lon"] = pd.to_numeric(df_base[col_lon], errors="coerce")

            df_resumen["distancia_calc"] = pd.to_numeric(df_resumen[col_distancia], errors="coerce")
            df_resumen["consumo_calc"] = pd.to_numeric(df_resumen[col_consumo], errors="coerce")

            trazados = (
                df_resumen[[col_trazado_res, "trazado_key"]]
                .dropna()
                .drop_duplicates()
                .rename(columns={col_trazado_res: "trazado"})
                .sort_values("trazado")
            )

            if trazados.empty:
                st.error("No se encontraron trazados válidos.")
                st.stop()

            csel1, csel2 = st.columns([2, 1])

            with csel1:
                trazado_sel = st.selectbox("Seleccionar trazado", trazados["trazado"].tolist())

            with csel2:
                bateria = st.selectbox("Batería HV (kWh)", [231.8, 255.0, 266.0, 310.0, 382.0], index=1)

            key = norm(trazado_sel)

            base = df_base[df_base["trazado_key"] == key].copy()
            base = base.dropna(subset=["odometro"]).sort_values("odometro")

            resumen_sel = df_resumen[df_resumen["trazado_key"] == key]

            if resumen_sel.empty:
                st.error("No se encontró el trazado en la hoja resumen.")
                st.stop()

            res = resumen_sel.iloc[0]

            distancia = float(res["distancia_calc"]) if pd.notna(res["distancia_calc"]) else None
            consumo = float(res["consumo_calc"]) if pd.notna(res["consumo_calc"]) else None

            if distancia is None or distancia <= 0:
                st.error("La distancia de la hoja resumen no es válida.")
                st.stop()

            if consumo is None or consumo <= 0:
                st.error("El consumo energético de la hoja resumen no es válido.")
                st.stop()

            rendimiento = consumo / distancia
            autonomia = bateria / rendimiento if rendimiento > 0 else None
            bateria_util_15 = bateria * 0.85
            autonomia_15 = bateria_util_15 / rendimiento if rendimiento > 0 else None
            vel_prom = base["velocidad"].mean() if base["velocidad"].notna().any() else None

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Rendimiento", f"{rendimiento:.3f} kWh/km")
            c2.metric("Autonomía proyectada", f"{autonomia:.0f} km" if autonomia is not None else "Sin dato")
            c3.metric("Velocidad promedio", f"{vel_prom:.1f} km/h" if vel_prom is not None else "Sin dato")
            c4.metric("Distancia", f"{distancia:.1f} km")

            st.markdown("<div style='height:25px;'></div>", unsafe_allow_html=True)

            g1, g2, g3 = st.columns(3)

            with g1:
                fig_g1 = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=rendimiento,
                    number={"suffix": " kWh/km", "font": {"size": 26}},
                    title={"text": "Rendimiento", "font": {"size": 18}},
                    gauge={
                        "axis": {"range": [0, 1.2], "tickfont": {"size": 11}},
                        "steps": [
                            {"range": [0, 0.90], "color": "#22c55e"},
                            {"range": [0.90, 1.00], "color": "#facc15"},
                            {"range": [1.00, 1.20], "color": "#ef4444"},
                        ],
                        "bar": {"color": "#15803d"},
                    }
                ))
                fig_g1.update_layout(height=240, margin=dict(l=5, r=5, t=45, b=10))
                st.plotly_chart(fig_g1, use_container_width=True)

            with g2:
                fig_g2 = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=autonomia if autonomia is not None else 0,
                    number={"suffix": " km", "font": {"size": 26}},
                    title={"text": f"Autonomía total ({bateria:.0f} kWh)", "font": {"size": 16}},
                    gauge={
                        "axis": {"range": [0, 500], "tickfont": {"size": 11}},
                        "steps": [
                            {"range": [0, 280], "color": "#ef4444"},
                            {"range": [280, 350], "color": "#facc15"},
                            {"range": [350, 500], "color": "#22c55e"},
                        ],
                        "bar": {"color": "#15803d"},
                    }
                ))
                fig_g2.update_layout(height=240, margin=dict(l=5, r=5, t=45, b=10))
                st.plotly_chart(fig_g2, use_container_width=True)

            with g3:
                fig_g3 = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=autonomia_15 if autonomia_15 is not None else 0,
                    number={"suffix": " km", "font": {"size": 26}},
                    title={"text": "Autonomía útil al 15% SoC", "font": {"size": 16}},
                    gauge={
                        "axis": {"range": [0, 500], "tickfont": {"size": 11}},
                        "steps": [
                            {"range": [0, 250], "color": "#ef4444"},
                            {"range": [250, 320], "color": "#facc15"},
                            {"range": [320, 500], "color": "#22c55e"},
                        ],
                        "bar": {"color": "#166534"},
                    }
                ))
                fig_g3.update_layout(height=240, margin=dict(l=5, r=5, t=45, b=10))
                st.plotly_chart(fig_g3, use_container_width=True)

            st.markdown("### Velocidad y Estado de Carga")

            fig_vs = go.Figure()
            fig_vs.add_trace(go.Scatter(
                x=base["odometro"],
                y=base["velocidad"],
                mode="lines+markers",
                name="Velocidad",
                line=dict(color="#2563eb", width=3, shape="spline"),
                marker=dict(size=5)
            ))

            if base["soc"].notna().any():
                fig_vs.add_trace(go.Scatter(
                    x=base["odometro"],
                    y=base["soc"],
                    mode="lines+markers",
                    name="SoC",
                    yaxis="y2",
                    line=dict(color="#60a5fa", width=3, shape="spline"),
                    marker=dict(size=4)
                ))

            fig_vs.update_layout(
                height=380,
                margin=dict(l=10, r=10, t=20, b=10),
                xaxis=dict(title="Odómetro km"),
                yaxis=dict(title="Velocidad km/h"),
                yaxis2=dict(title="Estado de carga SoC", overlaying="y", side="right"),
                legend=dict(orientation="h")
            )
            st.plotly_chart(fig_vs, use_container_width=True)

            st.markdown("### 🛰️ Mapa del recorrido")
            mapa = base.dropna(subset=["lat", "lon"]).copy()

            if mapa.empty:
                st.warning("No hay coordenadas válidas para mostrar el mapa.")
            else:
                mapa["tipo_punto"] = "Punto recorrido"
                mapa_inicio = mapa.iloc[[0]].copy()
                mapa_inicio["tipo_punto"] = "Inicio"
                mapa_fin = mapa.iloc[[-1]].copy()
                mapa_fin["tipo_punto"] = "Fin"

                mapa_plot = pd.concat([mapa, mapa_inicio, mapa_fin], ignore_index=True)

                fig_map = px.scatter_mapbox(
                    mapa_plot,
                    lat="lat",
                    lon="lon",
                    color="tipo_punto",
                    color_discrete_map={
                        "Punto recorrido": "#1d4ed8",
                        "Inicio": "#16a34a",
                        "Fin": "#dc2626"
                    },
                    hover_data={
                        "odometro": True,
                        "velocidad": True,
                        "soc": True,
                        "lat": False,
                        "lon": False
                    },
                    zoom=12,
                    height=520
                )

                fig_map.add_trace(go.Scattermapbox(
                    lat=mapa["lat"],
                    lon=mapa["lon"],
                    mode="lines",
                    line=dict(width=4, color="#2563eb"),
                    name="Ruta"
                ))

                fig_map.update_layout(
                    mapbox_style="open-street-map",
                    margin=dict(r=0, t=0, l=0, b=0),
                    legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1)
                )
                st.plotly_chart(fig_map, use_container_width=True)

            if base["altitud"].notna().any():
                st.markdown("### Perfil de altura")
                fig_alt = go.Figure()
                fig_alt.add_trace(go.Scatter(
                    x=base["odometro"],
                    y=base["altitud"],
                    fill="tozeroy",
                    mode="lines",
                    line=dict(color="#3b82f6", width=3, shape="spline")
                ))
                fig_alt.update_layout(
                    height=320,
                    margin=dict(l=10, r=10, t=20, b=10),
                    xaxis=dict(title="Odómetro"),
                    yaxis=dict(title="Altura (m)")
                )
                st.plotly_chart(fig_alt, use_container_width=True)

            try:
                pdf_bytes = generar_pdf_ejecutivo(
                    trazado=trazado_sel,
                    bateria=bateria,
                    distancia=distancia,
                    consumo=consumo,
                    rendimiento=rendimiento,
                    autonomia_total=autonomia,
                    autonomia_15=autonomia_15,
                    vel_prom=vel_prom,
                    base=base,
                    logo_path=LOGO_FILE
                )

                st.download_button(
                    "📄 Descargar informe PDF",
                    data=pdf_bytes,
                    file_name=f"Informe_Eficiencia_{trazado_sel.replace(' ', '_')}.pdf",
                    mime="application/pdf"
                )
            except Exception as e_pdf:
                st.warning(f"No fue posible generar el PDF: {e_pdf}")

            st.markdown("### Tabla resumen")
            st.dataframe(df_resumen_original, use_container_width=True)

            with st.expander("Ver detalle de la base"):
                st.dataframe(df_base_original.head(200), use_container_width=True)

        except Exception as e:
            st.error(f"Error: {e}")
# =========================================================
# TAB 4 - DASHBOARD
# =========================================================
with tab_dash:
    st.subheader("Dashboard Comercial")

    try:
        df_dash = cargar_historial()

        if df_dash.empty:
            st.info("No hay datos aún.")
        else:
            df_dash["total_negocio"] = pd.to_numeric(df_dash["total_negocio"], errors="coerce").fillna(0)
            df_dash["precio_unitario"] = pd.to_numeric(df_dash["precio_unitario"], errors="coerce").fillna(0)
            df_dash["cantidad_unidades"] = pd.to_numeric(df_dash["cantidad_unidades"], errors="coerce").fillna(0)
            df_dash["fecha_dt"] = pd.to_datetime(df_dash["fecha"], errors="coerce")

            c1, c2, c3 = st.columns(3)
            c1.metric("Total cotizaciones", len(df_dash))
            c2.metric("Monto total negocio", usd_fmt(df_dash["total_negocio"].sum()))
            c3.metric("Promedio por cotización", usd_fmt(df_dash["total_negocio"].mean()))

            df_precio = df_dash.dropna(subset=["fecha_dt"]).sort_values("fecha_dt")

            st.markdown("### Evolución del precio unitario")
            graf_precio = alt.Chart(df_precio).mark_line(point=True).encode(
                x=alt.X("fecha_dt:T", title="Fecha"),
                y=alt.Y("precio_unitario:Q", title="Precio unitario (USD)"),
                tooltip=["fecha", "cliente", "cotizante", "precio_unitario"]
            ).properties(height=320)
            st.altair_chart(graf_precio, use_container_width=True)

            st.markdown("### Evolución del total negocio")
            graf_total = alt.Chart(df_precio).mark_bar().encode(
                x=alt.X("fecha_dt:T", title="Fecha"),
                y=alt.Y("total_negocio:Q", title="Total negocio (USD)"),
                color=alt.Color("cotizante:N", title="Cotizante"),
                tooltip=["fecha", "cliente", "cotizante", "total_negocio"]
            ).properties(height=320)
            st.altair_chart(graf_total, use_container_width=True)

            st.markdown("### Relación entre precio unitario y total negocio")
            graf_scatter = alt.Chart(df_dash).mark_circle(size=120).encode(
                x=alt.X("precio_unitario:Q", title="Precio unitario (USD)"),
                y=alt.Y("total_negocio:Q", title="Total negocio (USD)"),
                color=alt.Color("cotizante:N", title="Cotizante"),
                size=alt.Size("cantidad_unidades:Q", title="Cantidad unidades"),
                tooltip=["cliente", "cotizante", "cantidad_unidades", "precio_unitario", "total_negocio"]
            ).properties(height=350)
            st.altair_chart(graf_scatter, use_container_width=True)

            st.markdown("### Detalle")
            df_v = df_dash.copy()
            df_v["precio_unitario"] = df_v["precio_unitario"].apply(usd_fmt)
            df_v["total_negocio"] = df_v["total_negocio"].apply(usd_fmt)
            st.dataframe(df_v, use_container_width=True)

    except Exception as e:
        st.error(f"Error dashboard: {e}")
        
# =========================================================
# TAB 5 - ESTRUCTURA DE COSTO
# =========================================================
if st.session_state.usuario in {"rsepulveda", "forellana", "dvejar"}:
    with tab_cost:
        st.subheader("Estructura de costo por modelo")

        puede_editar_costos = st.session_state.usuario in {"forellana", "dvejar"}

        st.caption("Visualización habilitada para Rodrigo Sepúlveda. Edición habilitada solo para Fabián Orellana y Diego Vejar.")

        modelos_iveco = [
            "IVECO BUS 10.5",
            "IVECO BUS 12",
            "IVECO DAILY",
            "IVECO OTRO"
        ]

        csel1, csel2 = st.columns([2, 1])

        with csel1:
            modelo_costo = st.selectbox(
                "Modelo",
                modelos_iveco,
                key="modelo_costo_tab5"
            )

        with csel2:
            dolar_bcch, mensaje_dolar = obtener_dolar_observado_bcch()

            if dolar_bcch is not None:
                st.caption(f"{mensaje_dolar}: {clp_fmt(dolar_bcch, decimales=2)}")
            else:
                st.warning(mensaje_dolar)

            dolar_observado = st.number_input(
                "Dólar observado CLP",
                min_value=1.0,
                value=float(dolar_bcch) if dolar_bcch is not None else 950.0,
                step=1.0,
                disabled=not puede_editar_costos,
                key="dolar_observado_tab5"
            )

        st.markdown("### Parámetros comerciales")

        p1, p2, p3 = st.columns(3)

        with p1:
            valor_final_usd_sin_iva = st.number_input(
                "Valor final USD sin IVA",
                min_value=0.0,
                value=100000.0,
                step=1000.0,
                disabled=not puede_editar_costos,
                key="valor_final_usd_sin_iva_tab5"
            )

            margen_importer_pct = st.number_input(
                "Margen Importer %",
                min_value=0.0,
                max_value=100.0,
                value=8.0,
                step=0.1,
                disabled=not puede_editar_costos,
                key="margen_importer_pct_tab5"
            )

        with p2:
            margen_dealer_pct = st.number_input(
                "% Mg Dealer CCS",
                min_value=0.0,
                max_value=100.0,
                value=5.0,
                step=0.1,
                disabled=not puede_editar_costos,
                key="margen_dealer_pct_tab5"
            )

            bono_vendedor_interno = st.number_input(
                "Bono vendedor interno USD",
                min_value=0.0,
                value=0.0,
                step=100.0,
                disabled=not puede_editar_costos,
                key="bono_vendedor_interno_tab5"
            )

        with p3:
            precio_venta_dealer_usd = st.number_input(
                "Precio Venta Dealer USD",
                min_value=0.0,
                value=0.0,
                step=1000.0,
                disabled=not puede_editar_costos,
                key="precio_venta_dealer_usd_tab5"
            )

            iva_pct = st.number_input(
                "IVA %",
                min_value=0.0,
                max_value=100.0,
                value=19.0,
                step=0.1,
                disabled=not puede_editar_costos,
                key="iva_pct_tab5"
            )

        iva_factor = 1 + (iva_pct / 100.0)

        margen_importer_usd = valor_final_usd_sin_iva * (margen_importer_pct / 100.0)

        st.markdown("### Indicadores base")

        k1, k2, k3, k4 = st.columns(4)

        with k1:
            st.metric("Valor final sin IVA", usd_fmt(valor_final_usd_sin_iva))

        with k2:
            st.metric("Margen Importer", usd_fmt(margen_importer_usd))

        with k3:
            st.metric("Precio Venta Dealer", usd_fmt(precio_venta_dealer_usd))

        with k4:
            st.metric("Dólar observado", clp_fmt(dolar_observado))

        st.markdown("### Objetivos de venta sobre valor final")

        porcentajes_objetivo = [5, 10, 15, 20]
        filas_objetivo = []

        for pct in porcentajes_objetivo:
            usd_sin_iva = valor_final_usd_sin_iva * (1 + pct / 100.0)
            usd_con_iva = usd_sin_iva * iva_factor
            clp_sin_iva = usd_sin_iva * dolar_observado
            clp_con_iva = usd_con_iva * dolar_observado

            filas_objetivo.append({
                "Objetivo": f"Objetivo +{pct}%",
                "USD sin IVA": usd_sin_iva,
                "USD con IVA": usd_con_iva,
                "CLP sin IVA": clp_sin_iva,
                "CLP con IVA": clp_con_iva,
            })

        df_objetivos = pd.DataFrame(filas_objetivo)

        st.markdown("#### Tabla de objetivos")

        df_objetivos_vista = df_objetivos.copy()
        df_objetivos_vista["USD sin IVA"] = df_objetivos_vista["USD sin IVA"].apply(usd_fmt)
        df_objetivos_vista["USD con IVA"] = df_objetivos_vista["USD con IVA"].apply(usd_fmt)
        df_objetivos_vista["CLP sin IVA"] = df_objetivos_vista["CLP sin IVA"].apply(clp_fmt)
        df_objetivos_vista["CLP con IVA"] = df_objetivos_vista["CLP con IVA"].apply(clp_fmt)

        st.dataframe(df_objetivos_vista, use_container_width=True)

        st.markdown("#### Gráfico de objetivos en USD sin IVA")
        fig_obj = px.bar(
            df_objetivos,
            x="Objetivo",
            y="USD sin IVA",
            text="USD sin IVA",
            title=f"Objetivos de venta - {modelo_costo}"
        )
        fig_obj.update_traces(texttemplate="%{y:,.0f}", textposition="outside")
        fig_obj.update_layout(height=420)
        st.plotly_chart(fig_obj, use_container_width=True)

        st.markdown("### Resumen comercial")

        resumen_df = pd.DataFrame({
            "Concepto": [
                "Modelo",
                "Valor final USD sin IVA",
                "Margen Importer %",
                "Margen Importer USD",
                "% Mg Dealer CCS",
                "Bono vendedor interno USD",
                "Precio Venta Dealer USD",
                "Dólar observado CLP",
                "IVA %"
            ],
            "Valor": [
                modelo_costo,
                usd_fmt(valor_final_usd_sin_iva),
                f"{margen_importer_pct:.1f}%",
                usd_fmt(margen_importer_usd),
                f"{margen_dealer_pct:.1f}%",
                usd_fmt(bono_vendedor_interno),
                usd_fmt(precio_venta_dealer_usd),
                clp_fmt(dolar_observado),
                f"{iva_pct:.1f}%"
            ]
        })

        st.dataframe(resumen_df, use_container_width=True, hide_index=True)

        if not puede_editar_costos:
            st.info("Este usuario tiene acceso solo de visualización.")