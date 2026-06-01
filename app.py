import hashlib
import hmac
from datetime import date
from io import BytesIO

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


st.set_page_config(
    page_title="SIGE-CTAR",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# CONFIGURACIÓN
# =========================

REQUIRED_SHEETS = [
    "FACT_CTAR_SEGUIMIENTO",
    "DIM_EQUIPO",
    "DIM_SERVICIO",
    "DIM_TIPO_PROCESO",
    "DIM_ESTADO",
    "DIM_RESPONSABLE",
    "DIM_PRIORIDAD",
    "FACT_BAJAS",
    "FACT_REPOSICIONES",
    "FACT_ADQUISICIONES",
    "FACT_ALERTAS",
]

DRIVE_FOLDER_ID = "1wS6MoYjcfZGbZRtbEjPQPQTC0g_A4IGQ"

# ID DEL ARCHIVO XLSX DE RESERVA PRESUPUESTARIA
RESERVA_FILE_ID = "1YU-dxAJlkUW7BLl2DWxssHb1Fo9dmH3m"
RESERVA_WORKSHEET_NAME = "Anexo I f)"


# =========================
# ESTILO
# =========================

st.markdown(
    """
    <style>
    .main { background: #f4f6fb; }
    .block-container { padding-top: 1rem; }

    .header {
        background: linear-gradient(90deg, #1a2b4a, #2563eb);
        color: white;
        padding: 22px;
        border-radius: 16px;
        margin-bottom: 18px;
    }

    .header h1 { color: white; margin: 0; }
    .header p { color: #dbeafe; margin: 4px 0 0 0; }

    .metric-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 16px;
        box-shadow: 0 4px 12px rgba(15,23,42,.06);
        margin-bottom: 10px;
    }

    .metric-title {
        font-size: 12px;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 700;
    }

    .metric-value {
        font-size: 30px;
        font-weight: 800;
        color: #1f2937;
    }

    .doc-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 15px;
        margin-bottom: 12px;
        box-shadow: 0 4px 12px rgba(15,23,42,.06);
    }

    .doc-title {
        font-size: 16px;
        font-weight: 700;
        color: #1f2937;
    }

    .doc-meta {
        font-size: 12px;
        color: #64748b;
        margin-top: 4px;
    }

    .doc-link {
        display: inline-block;
        margin-top: 8px;
        color: #2563eb;
        font-weight: 700;
        text-decoration: none;
    }

    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(5, minmax(160px, 1fr));
        gap: 14px;
        margin-bottom: 22px;
    }

    .kpi-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 4px 12px rgba(15,23,42,.08);
        min-height: 135px;
        overflow: hidden;
    }

    .kpi-title {
        font-size: 13px;
        font-weight: 800;
        color: #64748b;
        text-transform: uppercase;
        margin-bottom: 12px;
        line-height: 1.25;
    }

    .kpi-value {
        font-size: clamp(22px, 2.1vw, 30px);
        font-weight: 900;
        color: #1f2937;
        line-height: 1.1;
        white-space: nowrap;
    }

    .kpi-negative {
        color: #b91c1c;
    }

    .kpi-positive {
        color: #047857;
    }

    @media (max-width: 1200px) {
        .kpi-grid {
            grid-template-columns: repeat(2, 1fr);
        }
    }

    @media (max-width: 700px) {
        .kpi-grid {
            grid-template-columns: 1fr;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# LOGIN
# =========================

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def get_users():
    try:
        return dict(st.secrets["auth"]["users"])
    except Exception:
        return {
            "admin": {
                "name": "Administrador CTAR",
                "password_hash": hash_password("admin123"),
                "role": "admin",
            },
            "ctar": {
                "name": "Usuario CTAR",
                "password_hash": hash_password("ctar123"),
                "role": "ctar",
            },
            "hospital": {
                "name": "Usuario Hospital",
                "password_hash": hash_password("hospital123"),
                "role": "hospital",
            },
            "ifiscal": {
                "name": "Inspector Fiscal",
                "password_hash": hash_password("if123"),
                "role": "if",
            },
        }


def authenticate(username, password):
    users = get_users()

    if username not in users:
        return None

    user = users[username]

    if hmac.compare_digest(
        user["password_hash"],
        hash_password(password),
    ):
        return {
            "username": username,
            "name": user["name"],
            "role": user["role"],
        }

    return None


def login():
    st.markdown(
        """
        <style>
        .login-title {
            text-align: center;
            font-size: 42px;
            font-weight: 800;
            color: #1f2937;
            margin-bottom: 0px;
        }

        .login-subtitle {
            text-align: center;
            color: #6b7280;
            font-size: 18px;
            margin-bottom: 25px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1.2, 1, 1.2])

    with col2:
        st.markdown(
            """
            <div class="login-title">🏥 SIGE-CTAR</div>
            <div class="login-subtitle">Sistema de Gestión y Trazabilidad CTAR</div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("login"):
            username = st.text_input("Usuario")
            password = st.text_input("Clave", type="password")
            submit = st.form_submit_button(
                "Ingresar",
                use_container_width=True,
            )

    if submit:
        user = authenticate(username.strip(), password)

        if user:
            st.session_state["user"] = user
            st.rerun()
        else:
            st.error("Usuario o clave incorrecta.")


if "user" not in st.session_state:
    login()
    st.stop()


def role():
    return st.session_state["user"]["role"]


def can_edit():
    return role() in ["admin", "ctar", "if"]


# =========================
# GOOGLE SHEETS / DRIVE
# =========================

def get_credentials():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes,
    )

    return creds


@st.cache_data(ttl=60)
def load_google_sheets():
    creds = get_credentials()
    client = gspread.authorize(creds)

    spreadsheet_id = st.secrets["google_sheet"]["spreadsheet_id"]
    spreadsheet = client.open_by_key(spreadsheet_id)

    tables = {}

    for ws in spreadsheet.worksheets():
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        df.columns = [str(c).strip().replace(" ", "_") for c in df.columns]
        tables[ws.title] = df

    return tables


def append_to_sheet(sheet_name, row_values):
    creds = get_credentials()
    client = gspread.authorize(creds)

    spreadsheet_id = st.secrets["google_sheet"]["spreadsheet_id"]
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet(sheet_name)
    worksheet.append_row(row_values, value_input_option="USER_ENTERED")

    st.cache_data.clear()


@st.cache_data(ttl=60)
def list_drive_files(folder_id):
    creds = get_credentials()

    service = build(
        "drive",
        "v3",
        credentials=creds,
    )

    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(id, name, mimeType, webViewLink, modifiedTime)",
        orderBy="modifiedTime desc",
        pageSize=100,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    return results.get("files", [])


@st.cache_data(ttl=60)
def load_reserva_presupuestaria():
    creds = get_credentials()

    service = build(
        "drive",
        "v3",
        credentials=creds,
    )

    request = service.files().get_media(
        fileId=RESERVA_FILE_ID,
        supportsAllDrives=True,
    )

    file_buffer = BytesIO()
    downloader = MediaIoBaseDownload(file_buffer, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    file_buffer.seek(0)

    df = pd.read_excel(
        file_buffer,
        sheet_name=RESERVA_WORKSHEET_NAME,
        header=None,
        engine="openpyxl",
    )

    return df


def check_sheets(tables):
    missing = [s for s in REQUIRED_SHEETS if s not in tables]
    return missing


# =========================
# MODELO DE DATOS
# =========================

def build_model(tables):
    fact = tables.get("FACT_CTAR_SEGUIMIENTO", pd.DataFrame()).copy()

    if fact.empty:
        return fact

    for col in ["Fecha_Ingreso", "Fecha_Compromiso"]:
        if col in fact.columns:
            fact[col] = pd.to_datetime(fact[col], errors="coerce")

    joins = [
        ("DIM_EQUIPO", "ID_Equipo"),
        ("DIM_SERVICIO", "ID_Servicio"),
        ("DIM_TIPO_PROCESO", "ID_Tipo_Proceso"),
        ("DIM_ESTADO", "ID_Estado"),
        ("DIM_RESPONSABLE", "ID_Responsable"),
        ("DIM_PRIORIDAD", "ID_Prioridad"),
    ]

    df = fact.copy()

    for sheet, key in joins:
        dim = tables.get(sheet, pd.DataFrame()).copy()

        if not dim.empty and key in df.columns and key in dim.columns:
            df = df.merge(dim, on=key, how="left")

    today = pd.Timestamp(date.today())

    if "Fecha_Ingreso" in df.columns:
        df["Dias_Desde_Ingreso"] = (today - df["Fecha_Ingreso"]).dt.days
    else:
        df["Dias_Desde_Ingreso"] = 0

    if "Fecha_Compromiso" in df.columns:
        df["Dias_Atraso"] = (today - df["Fecha_Compromiso"]).dt.days
        df["Vencido"] = df["Dias_Atraso"] > 0
    else:
        df["Dias_Atraso"] = 0
        df["Vencido"] = False

    for col in [
        "Equipo",
        "Servicio",
        "Tipo_Proceso",
        "Estado",
        "Responsable",
        "Prioridad",
        "Riesgo_Clinico",
        "Proxima_Accion",
    ]:
        if col not in df.columns:
            df[col] = ""

    return df


# =========================
# COMPONENTES
# =========================

def header(title, subtitle):
    st.markdown(
        f"""
        <div class="header">
            <h1>{title}</h1>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(title, value):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def filters(df):
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        servicio = st.selectbox(
            "Servicio",
            ["Todos"] + sorted(df["Servicio"].dropna().astype(str).unique().tolist()),
        )

    with c2:
        tipo = st.selectbox(
            "Tipo proceso",
            ["Todos"] + sorted(df["Tipo_Proceso"].dropna().astype(str).unique().tolist()),
        )

    with c3:
        estado = st.selectbox(
            "Estado",
            ["Todos"] + sorted(df["Estado"].dropna().astype(str).unique().tolist()),
        )

    with c4:
        responsable = st.selectbox(
            "Responsable",
            ["Todos"] + sorted(df["Responsable"].dropna().astype(str).unique().tolist()),
        )

    out = df.copy()

    if servicio != "Todos":
        out = out[out["Servicio"] == servicio]

    if tipo != "Todos":
        out = out[out["Tipo_Proceso"] == tipo]

    if estado != "Todos":
        out = out[out["Estado"] == estado]

    if responsable != "Todos":
        out = out[out["Responsable"] == responsable]

    return out


def icon_by_mimetype(mime_type):
    mime_type = str(mime_type).lower()

    if "pdf" in mime_type:
        return "📕"

    if "spreadsheet" in mime_type or "excel" in mime_type:
        return "📊"

    if "document" in mime_type or "word" in mime_type:
        return "📄"

    if "presentation" in mime_type or "powerpoint" in mime_type:
        return "📽️"

    if "folder" in mime_type:
        return "📁"

    return "📎"


def clean_number(value):
    value = str(value).strip()
    value = value.replace(",", "")
    value = value.replace(" ", "")

    return pd.to_numeric(value, errors="coerce")


# =========================
# CARGA
# =========================

try:
    tables = load_google_sheets()
except Exception as e:
    st.error("No se pudo conectar con Google Sheets.")
    st.code(str(e))
    st.stop()

missing = check_sheets(tables)

if missing:
    st.warning("Faltan hojas en Google Sheets:")
    st.write(missing)

df = build_model(tables)


# =========================
# MENÚ
# =========================

with st.sidebar:
    st.markdown("## 🏥 SIGE-CTAR")
    st.caption("Sistema CTAR")

    user = st.session_state["user"]

    st.markdown("---")
    st.caption(f"Usuario: {user['name']}")
    st.caption(f"Rol: {user['role']}")

    pages = [
        "Resumen Ejecutivo",
        "Seguimiento",
        "Bajas",
        "Reposiciones",
        "Adquisiciones",
        "Repositorio Documental",
        "Reserva Presupuestaria",
    ]

    if role() != "hospital":
        pages.append("Alertas")

    if can_edit():
        pages.append("Registro")

    if role() == "admin":
        pages.append("Configuración")

    page = st.radio("Menú", pages)

    if st.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()


# =========================
# PÁGINAS
# =========================

if page == "Resumen Ejecutivo":
    header(
        "SIGE-CTAR · Resumen Ejecutivo",
        "Seguimiento de bajas, reposiciones, adquisiciones, SIC y alertas.",
    )

    if df.empty:
        st.info("No hay datos cargados en Google Sheets.")
        st.stop()

    view = filters(df)

    total = len(view)
    revision = view["Estado"].astype(str).str.lower().str.contains("revisión|revision").sum()
    aprobadas = view["Estado"].astype(str).str.lower().str.contains("aprob").sum()
    compra = view["Estado"].astype(str).str.lower().str.contains("compra").sum()
    vencidas = int(view["Vencido"].sum())
    altas = view["Prioridad"].astype(str).str.lower().eq("alta").sum()

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    with c1:
        metric_card("Total", total)

    with c2:
        metric_card("En revisión", revision)

    with c3:
        metric_card("Aprobadas", aprobadas)

    with c4:
        metric_card("En compra", compra)

    with c5:
        metric_card("Vencidas", vencidas)

    with c6:
        metric_card("Prioridad alta", altas)

    st.markdown("---")

    g1, g2, g3 = st.columns(3)

    with g1:
        st.markdown("### Estados")
        fig = px.bar(
            view.groupby("Estado").size().reset_index(name="Cantidad"),
            x="Estado",
            y="Cantidad",
            text="Cantidad",
        )
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        st.markdown("### Tipo proceso")
        fig = px.pie(view, names="Tipo_Proceso", hole=0.45)
        st.plotly_chart(fig, use_container_width=True)

    with g3:
        st.markdown("### Prioridad")
        fig = px.pie(view, names="Prioridad", hole=0.45)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Alertas principales")

    alertas = view[
        (view["Prioridad"].astype(str).str.lower() == "alta")
        | (view["Vencido"] == True)
        | (view["Estado"].astype(str).str.lower().str.contains("observ", na=False))
    ]

    st.dataframe(alertas, use_container_width=True, hide_index=True)


elif page == "Seguimiento":
    header(
        "Seguimiento CTAR",
        "Consulta por SIC, equipo, servicio, estado, responsable y prioridad.",
    )

    if df.empty:
        st.info("No hay datos.")
        st.stop()

    view = filters(df)

    search = st.text_input("Buscar SIC, equipo, servicio o inventario")

    if search:
        s = search.lower()
        mask = pd.Series(False, index=view.index)

        for col in ["SIC", "Equipo", "Servicio", "Nro_Inventario", "Motivo"]:
            if col in view.columns:
                mask = mask | view[col].astype(str).str.lower().str.contains(s, na=False)

        view = view[mask]

    cols = [
        "ID_CTAR",
        "SIC",
        "Equipo",
        "Nro_Inventario",
        "Servicio",
        "Tipo_Proceso",
        "Estado",
        "Responsable",
        "Prioridad",
        "Fecha_Ingreso",
        "Fecha_Compromiso",
        "Dias_Desde_Ingreso",
        "Dias_Atraso",
        "Motivo",
        "Riesgo_Clinico",
        "Proxima_Accion",
        "Link_Documento",
    ]

    st.dataframe(
        view[[c for c in cols if c in view.columns]],
        use_container_width=True,
        hide_index=True,
    )


elif page == "Bajas":
    header(
        "Bajas y Extravíos",
        "Control ejecutivo de equipos dados de baja, extravíos, hurtos y seguimiento CTAR.",
    )

    def prioridad_baja(row):
        texto = " ".join([
            str(row.get("Motivo_Baja", "")),
            str(row.get("Estado_Baja", "")),
            str(row.get("CAUSAL RECLAMADA", "")),
            str(row.get("CAUSAL APROBADA", "")),
            str(row.get("Observaciones", "")),
            str(row.get("DETALLE", "")),
            str(row.get("BAJA DE:", "")),
            str(row.get("TIPO/EQUIPO", "")),
        ]).lower()

        if "hurto" in texto or "robo" in texto or "extrav" in texto:
            return "Alta"

        if "en revisión" in texto or "revision" in texto or "pendiente" in texto:
            return "Media"

        if "aprob" in texto or "cerrado" in texto or "entregado" in texto:
            return "Baja"

        return "Media"

    def color_prioridad(row):
        prioridad = str(row.get("Prioridad_Baja", "")).lower()

        if prioridad == "alta":
            return [
                "background-color: #fee2e2; color: #7f1d1d; font-weight: bold"
            ] * len(row)

        if prioridad == "media":
            return [
                "background-color: #fef3c7; color: #78350f"
            ] * len(row)

        if prioridad == "baja":
            return [
                "background-color: #dcfce7; color: #14532d"
            ] * len(row)

        return [""] * len(row)

    st.markdown("## 📌 Resumen de bajas")

    bajas = tables.get("FACT_BAJAS", pd.DataFrame()).copy()

    if bajas.empty:
        st.info("No hay registros cargados en FACT_BAJAS.")
    else:
        if "ID_CTAR" in bajas.columns and "ID_CTAR" in df.columns:
            base = bajas.merge(
                df,
                on="ID_CTAR",
                how="left",
                suffixes=("_Baja", ""),
            )
        else:
            base = bajas.copy()

        base["Prioridad_Baja"] = base.apply(prioridad_baja, axis=1)

        total_bajas = len(base)
        altas = int((base["Prioridad_Baja"] == "Alta").sum())
        medias = int((base["Prioridad_Baja"] == "Media").sum())
        bajas_prioridad = int((base["Prioridad_Baja"] == "Baja").sum())

        revision = int(
            base.astype(str)
            .apply(
                lambda x: x.str.lower().str.contains(
                    "revisión|revision|pendiente",
                    na=False,
                    regex=True,
                )
            )
            .any(axis=1)
            .sum()
        )

        c1, c2, c3, c4, c5 = st.columns(5)

        with c1:
            st.metric("Total bajas", total_bajas)

        with c2:
            st.metric("Prioridad alta", altas)

        with c3:
            st.metric("Prioridad media", medias)

        with c4:
            st.metric("En revisión / pendiente", revision)

        with c5:
            st.metric("Prioridad baja", bajas_prioridad)

        st.markdown("---")

        st.markdown("## 🚦 Priorización")

        resumen_prioridad = base.groupby("Prioridad_Baja").size().reset_index(
            name="Cantidad"
        )

        fig = px.bar(
            resumen_prioridad,
            x="Prioridad_Baja",
            y="Cantidad",
            text="Cantidad",
            title="Distribución de bajas según prioridad",
        )

        fig.update_layout(
            height=420,
            xaxis_title="Prioridad",
            yaxis_title="Cantidad de registros",
            title_x=0.02,
        )

        st.plotly_chart(fig, use_container_width=True)

        st.markdown("## 📋 Registro consolidado de bajas")

        columnas_preferidas = [
            "ID_Baja",
            "ID_CTAR",
            "SIC",
            "Equipo",
            "Nro_Inventario",
            "Servicio",
            "Motivo_Baja",
            "Estado_Baja",
            "Fecha_Baja",
            "Prioridad_Baja",
            "Responsable",
            "Estado",
            "Proxima_Accion",
            "Link_Documento",
        ]

        columnas_disponibles = [
            c for c in columnas_preferidas if c in base.columns
        ]

        if columnas_disponibles:
            tabla_bajas = base[columnas_disponibles].copy()
        else:
            tabla_bajas = base.copy()

        st.dataframe(
            tabla_bajas.style.apply(color_prioridad, axis=1),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("---")

    st.markdown("## 📤 Cargar archivo hospital")

    archivo_hospital = st.file_uploader(
        "Subir archivo Excel del Hospital, por ejemplo: Copia de 7.- Equipos de Baja LRP.xlsx",
        type=["xlsx"],
    )

    if archivo_hospital is not None:
        try:
            try:
                hospital_df = pd.read_excel(
                    archivo_hospital,
                    sheet_name="V1",
                    header=1,
                    engine="openpyxl",
                )
            except Exception:
                archivo_hospital.seek(0)
                hospital_df = pd.read_excel(
                    archivo_hospital,
                    sheet_name=0,
                    header=1,
                    engine="openpyxl",
                )

            hospital_df = hospital_df.dropna(how="all")
            hospital_df.columns = [str(c).strip() for c in hospital_df.columns]

            hospital_df["Prioridad_Baja"] = hospital_df.apply(
                prioridad_baja,
                axis=1,
            )

            st.success(
                f"Archivo hospital cargado correctamente: {len(hospital_df)} registros detectados."
            )

            st.markdown("## 🏥 Información entregada por Hospital")

            columnas_hospital = [
                "CTAR",
                "OFICIO",
                "FECHA",
                "CARTA/ORD",
                "FECHA CARTA",
                "FECHA INCIDENCIA",
                "BAJA DE:",
                "TIPO/EQUIPO",
                "C.RECINTO",
                "SIC",
                "Estado SIC",
                "UNIDAD",
                "NOMBRE EQUIPO",
                "MARCA",
                "MODELO",
                "N° SERIE",
                "N° INVENTARIO",
                "PROVEEDOR",
                "CAUSAL RECLAMADA",
                "CAUSAL APROBADA",
                "DETALLE",
                "VALOR (NETO)",
                "CARGADO A FONDO",
                "ORDEN DE COMPRA",
                "OT ENTREGA AL SERVICIO",
                "CARGADO AL SIC",
                "Observaciones",
                "Prioridad_Baja",
            ]

            columnas_hospital_disponibles = [
                c for c in columnas_hospital if c in hospital_df.columns
            ]

            if columnas_hospital_disponibles:
                vista_hospital = hospital_df[columnas_hospital_disponibles].copy()
            else:
                vista_hospital = hospital_df.copy()

            st.dataframe(
                vista_hospital.style.apply(color_prioridad, axis=1),
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("## 📊 Resumen archivo hospital")

            resumen_hospital = hospital_df.groupby("Prioridad_Baja").size().reset_index(
                name="Cantidad"
            )

            fig_hospital = px.bar(
                resumen_hospital,
                x="Prioridad_Baja",
                y="Cantidad",
                text="Cantidad",
                title="Cantidad de bajas por prioridad según archivo hospital",
            )

            fig_hospital.update_layout(
                height=420,
                xaxis_title="Prioridad",
                yaxis_title="Cantidad",
                title_x=0.02,
            )

            st.plotly_chart(fig_hospital, use_container_width=True)

        except Exception as e:
            st.error("No se pudo procesar el archivo del Hospital.")
            st.code(str(e))


elif page == "Reposiciones":
    header(
        "Reposiciones",
        "Seguimiento desde solicitud hasta compra, recepción o cierre.",
    )

    view = df[
        df["Tipo_Proceso"].astype(str).str.lower().str.contains("repos", na=False)
    ]

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Adquisiciones":
    header(
        "Adquisiciones",
        "Control de procesos de compra, BACO, OC y proveedor.",
    )

    view = df[
        df["Tipo_Proceso"].astype(str).str.lower().str.contains("adquis|compra", na=False)
    ]

    adq = tables.get("FACT_ADQUISICIONES", pd.DataFrame())

    if not adq.empty and "ID_CTAR" in adq.columns:
        view = view.merge(adq, on="ID_CTAR", how="left")

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Repositorio Documental":
    header(
        "Repositorio Documental CTAR",
        "Consulta de documentos almacenados en Google Drive.",
    )

    st.markdown("### 📂 Documentos disponibles")

    try:
        files = list_drive_files(DRIVE_FOLDER_ID)

        if not files:
            st.info("No se encontraron documentos en la carpeta de Google Drive.")
        else:
            st.success(f"Se encontraron {len(files)} documentos en el repositorio.")

            search_doc = st.text_input("Buscar documento por nombre")

            if search_doc:
                files = [
                    f for f in files
                    if search_doc.lower() in f.get("name", "").lower()
                ]

            if not files:
                st.warning("No se encontraron documentos con ese criterio de búsqueda.")

            for f in files:
                icon = icon_by_mimetype(f.get("mimeType", ""))
                name = f.get("name", "Sin nombre")
                link = f.get("webViewLink", "#")
                modified = f.get("modifiedTime", "")

                st.markdown(
                    f"""
                    <div class="doc-card">
                        <div class="doc-title">{icon} {name}</div>
                        <div class="doc-meta">Última modificación: {modified}</div>
                        <a class="doc-link" href="{link}" target="_blank">Abrir documento</a>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    except Exception as e:
        st.error("No se pudo acceder a la carpeta de Google Drive.")
        st.warning(
            "Verifica que la Google Drive API esté habilitada y que la cuenta de servicio tenga permiso de lector en la carpeta."
        )
        st.code(str(e))


elif page == "Reserva Presupuestaria":
    header(
        "Reserva Presupuestaria",
        "Dashboard ejecutivo del Anexo I f) · Fondo de reserva y gastos durante la explotación.",
    )

    try:
        df_raw = load_reserva_presupuestaria()

        if df_raw.empty:
            st.info("No se encontraron datos en la hoja Anexo I f).")
            st.stop()

        tabla = df_raw.iloc[22:30, 1:9].copy()

        tabla.columns = [
            "Año",
            "Suplemento",
            "VMA",
            "Total general",
            "VMA año explotación",
            "Desembolsos explotación",
            "En revisión desembolsos",
            "Diferencia proyectada",
        ]

        tabla = tabla.dropna(how="all")
        tabla = tabla[tabla["Año"].astype(str).str.strip() != ""]

        columnas_uf = [
            "Suplemento",
            "VMA",
            "Total general",
            "VMA año explotación",
            "Desembolsos explotación",
            "En revisión desembolsos",
            "Diferencia proyectada",
        ]

        for col in columnas_uf:
            tabla[col] = tabla[col].apply(clean_number).fillna(0)

        fila_total = tabla[
            tabla["Año"].astype(str).str.contains("Total", case=False, na=False)
        ]

        tabla_anios = tabla[
            ~tabla["Año"].astype(str).str.contains("Total", case=False, na=False)
        ].copy()

        total_general = fila_total["Total general"].sum()
        total_vma = fila_total["VMA año explotación"].sum()
        total_desembolsos = fila_total["Desembolsos explotación"].sum()
        total_revision = fila_total["En revisión desembolsos"].sum()
        diferencia = fila_total["Diferencia proyectada"].sum()

        porcentaje_usado = 0
        if total_vma > 0:
            porcentaje_usado = ((total_desembolsos + total_revision) / total_vma) * 100

        saldo_disponible = total_vma - total_desembolsos - total_revision

        st.markdown("## 📌 Resumen ejecutivo financiero")

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric(
                "Fondo VMA explotación",
                f"{total_vma:,.2f} UF"
            )

        with c2:
            st.metric(
                "Total desembolsado",
                f"{total_desembolsos:,.2f} UF"
            )

        with c3:
            st.metric(
                "Saldo disponible estimado",
                f"{saldo_disponible:,.2f} UF"
            )

        c4, c5, c6 = st.columns(3)

        with c4:
            st.metric(
                "En revisión",
                f"{total_revision:,.2f} UF"
            )

        with c5:
            st.metric(
                "Diferencia proyectada",
                f"{diferencia:,.2f} UF"
            )

        with c6:
            st.metric(
                "Uso estimado del fondo",
                f"{porcentaje_usado:,.1f}%"
            )

        st.progress(min(max(porcentaje_usado / 100, 0), 1))

        if diferencia < 0:
            st.error(
                "El análisis muestra una diferencia proyectada negativa, por lo que corresponde mantener seguimiento y control financiero del fondo de reserva."
            )
        else:
            st.success(
                "El análisis muestra una diferencia proyectada positiva o controlada respecto del fondo de reserva."
            )

        st.markdown("---")

        st.markdown("## 📋 Tabla ejecutiva consolidada")

        st.dataframe(
            tabla.style.format({
                col: "{:,.2f}" for col in columnas_uf
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("---")

        st.markdown("## 📊 Comparativo financiero por año")

        grafico_base = tabla_anios.melt(
            id_vars="Año",
            value_vars=[
                "VMA año explotación",
                "Desembolsos explotación",
                "En revisión desembolsos",
            ],
            var_name="Concepto",
            value_name="UF",
        )

        fig = px.bar(
            grafico_base,
            x="Año",
            y="UF",
            color="Concepto",
            barmode="group",
            text_auto=",.0f",
            title="VMA, desembolsos y montos en revisión por año de explotación",
        )

        fig.update_layout(
            height=520,
            xaxis_title="Año de explotación",
            yaxis_title="Monto UF",
            legend_title="Concepto",
            title_x=0.02,
            bargap=0.25,
        )

        st.plotly_chart(fig, use_container_width=True)

        st.markdown("## 🔎 Diferencia proyectada por año")

        fig2 = px.bar(
            tabla_anios,
            x="Año",
            y="Diferencia proyectada",
            text="Diferencia proyectada",
            title="Diferencia proyectada respecto del fondo de reserva",
        )

        fig2.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
        )

        fig2.update_layout(
            height=500,
            xaxis_title="Año de explotación",
            yaxis_title="Diferencia proyectada UF",
            title_x=0.02,
        )

        st.plotly_chart(fig2, use_container_width=True)

        st.markdown("---")

        st.markdown("## 🧾 Lectura ejecutiva")

        st.info(
            f"""
            De acuerdo con el Anexo I f), el VMA total del año de explotación asciende a **{total_vma:,.2f} UF**. 
            A la fecha, se registran desembolsos por **{total_desembolsos:,.2f} UF** y montos en revisión por 
            **{total_revision:,.2f} UF**, lo que representa un uso estimado del fondo de **{porcentaje_usado:,.1f}%**.

            La diferencia proyectada consolidada corresponde a **{diferencia:,.2f} UF**.
            """
        )

        with st.expander("📄 Ver planilla original completa"):
            st.dataframe(
                df_raw,
                use_container_width=True,
                hide_index=True,
            )

    except Exception as e:
        st.error("No se pudo cargar la Reserva Presupuestaria.")
        st.warning(
            "Verifica que la Google Drive API esté habilitada, que el archivo sea .xlsx y que la cuenta de servicio tenga permiso de lector."
        )
        st.code(str(e))


elif page == "Alertas":
    header(
        "Alertas",
        "Procesos críticos, atrasados u observados.",
    )

    view = df[
        (df["Prioridad"].astype(str).str.lower() == "alta")
        | (df["Vencido"] == True)
        | (df["Estado"].astype(str).str.lower().str.contains("observ", na=False))
    ]

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Registro":
    header(
        "Registro de Solicitud",
        "Ingreso directo a Google Sheets.",
    )

    if not can_edit():
        st.error("No tiene permisos para registrar.")
        st.stop()

    with st.form("registro"):
        id_ctar = st.text_input("ID CTAR", value=f"CTAR-{date.today().year}-XXX")
        sic = st.text_input("SIC")
        id_equipo = st.text_input("ID Equipo")
        id_servicio = st.text_input("ID Servicio")
        id_tipo = st.text_input("ID Tipo Proceso")
        id_estado = st.text_input("ID Estado")
        id_resp = st.text_input("ID Responsable")
        id_prioridad = st.text_input("ID Prioridad")
        fecha_ingreso = st.date_input("Fecha ingreso", value=date.today())
        fecha_compromiso = st.date_input("Fecha compromiso", value=date.today())
        motivo = st.text_area("Motivo")
        riesgo = st.text_area("Riesgo clínico")
        ultima = st.text_area("Última gestión")
        proxima = st.text_area("Próxima acción")
        link = st.text_input("Link documento")

        submit = st.form_submit_button("Guardar en Google Sheets")

    if submit:
        row = [
            id_ctar,
            sic,
            id_equipo,
            id_servicio,
            id_tipo,
            id_estado,
            id_resp,
            id_prioridad,
            str(fecha_ingreso),
            str(fecha_compromiso),
            "",
            motivo,
            riesgo,
            ultima,
            proxima,
            link,
        ]

        try:
            append_to_sheet("FACT_CTAR_SEGUIMIENTO", row)
            st.success("Solicitud guardada correctamente.")
            st.rerun()

        except Exception as e:
            st.error("No se pudo guardar en Google Sheets.")
            st.code(str(e))


elif page == "Configuración":
    header(
        "Configuración",
        "Validación de conexión y estructura Google Sheets.",
    )

    st.markdown("### Hojas detectadas")
    st.write(list(tables.keys()))

    st.markdown("### Hojas faltantes")
    st.write(missing if missing else "No faltan hojas.")

    st.markdown("### Cantidad de registros")
    st.write({k: len(v) for k, v in tables.items()})
