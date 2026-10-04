import streamlit as st
import pandas as pd
import os

# Configuración de página
st.set_page_config(
    page_title="INLAND | Diccionario Eólico de Alarmas",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -------------------------------------------------------------
# ESTILOS CSS - ECO-TECH / SCADA AEROGENERADORES
# -------------------------------------------------------------
st.markdown("""
<style>
    /* Estilos generales */
    .main {
        background-color: #0b111e;
    }
    
    /* Banner Cabecera */
    .hero-banner {
        background: linear-gradient(135deg, #064e3b 0%, #065f46 40%, #0f172a 100%);
        border: 1px solid #10b981;
        padding: 24px;
        border-radius: 14px;
        color: #ffffff;
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(16, 185, 129, 0.15);
    }
    .hero-title {
        font-size: 26px;
        font-weight: 700;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .hero-subtitle {
        font-size: 14px;
        color: #a7f3d0;
        margin: 0;
    }

    /* Tarjetas KPI */
    .kpi-container {
        display: flex;
        gap: 15px;
        margin-bottom: 25px;
    }
    .kpi-card {
        background: #131d2e;
        border: 1px solid #1e293b;
        border-radius: 10px;
        padding: 14px 18px;
        flex: 1;
        border-left: 4px solid #10b981;
    }
    .kpi-card-crit {
        border-left: 4px solid #ef4444;
    }
    .kpi-num {
        font-size: 24px;
        font-weight: 700;
        color: #ffffff;
    }
    .kpi-label {
        font-size: 12px;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* Ficha Técnica de Alarma */
    .alarm-card {
        background-color: #111a28;
        border: 1px solid #1e293b;
        border-radius: 12px;
        padding: 22px;
        margin-top: 15px;
        margin-bottom: 25px;
    }
    .alarm-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #1e293b;
        padding-bottom: 14px;
        margin-bottom: 16px;
    }
    .alarm-code {
        font-size: 22px;
        font-weight: 800;
        color: #38bdf8;
        background: #0c4a6e;
        padding: 4px 12px;
        border-radius: 6px;
        display: inline-block;
    }
    .badge-stopped {
        background: rgba(239, 68, 68, 0.2);
        color: #f87171;
        border: 1px solid #ef4444;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 13px;
    }
    .badge-running {
        background: rgba(16, 185, 129, 0.2);
        color: #34d399;
        border: 1px solid #10b981;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 13px;
    }

    /* Cajas internas de la ficha */
    .prop-box {
        background: #172234;
        border: 1px solid #24344d;
        border-radius: 8px;
        padding: 12px 14px;
        margin-bottom: 10px;
    }
    .prop-title {
        font-size: 11px;
        text-transform: uppercase;
        color: #64748b;
        font-weight: 700;
        margin-bottom: 4px;
    }
    .prop-value {
        font-size: 14px;
        color: #e2e8f0;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# DICCIONARIOS DE DECODIFICACIÓN TÉCNICA
# -------------------------------------------------------------
RESET_MODE_MAP = {
    "-": "Sin modo de restablecimiento",
    "A": "Restablecimiento automático de la alarma",
    "M": "Se requiere restablecimiento manual",
    "M(A)": "Restablecimiento manual (Automático por ARS)",
    "ML": "Restablecimiento manual por el operador localmente en el aerogenerador",
    "SL": "Restablecimiento local de la alarma de seguridad requerido",
    "SR": "Restablecimiento remoto de la alarma de seguridad (Solo K08Delta)"
}

BRAKE_PROG_MAP = {
    0: "Sin frenado",
    1: "Frenado suave",
    2: "Frenado rápido",
    3: "Frenado rápido con parada del rotor",
    5: "Frenado suave con giro libre rápido (fast idling)",
    6: "Frenado rápido con giro libre rápido (fast idling)"
}

YAW_PROG_MAP = {
    0: "Modo automático",
    10: "Evitación de sector (Sector Avoidance)",
    20: "Modo económico",
    24: "Modo automático a 90°",
    25: "Modo automático bloqueado",
    28: "Destorsión / desenrollado de cables",
    29: "Posición de parada / estacionamiento (Parking)",
    30: "Modo manual",
    33: "Modo manual: Posición objetivo",
    35: "Modo manual: Panel de operador",
    46: "Prueba de frenos",
    47: "Prueba del encoder del motor",
    48: "Prueba de ejes",
    50: "Parada de orientación (Yaw Stop)",
    55: "Parada de orientación y desconexión / apagado"
}

# -------------------------------------------------------------
# DETECCIÓN DINÁMICA DE COLUMNAS (FLEXIBLE AL EXCEL)
# -------------------------------------------------------------
def detectar_columna(df, variantes, default_idx=0):
    columnas_limpias = {str(c).strip().lower(): c for c in df.columns}
    for variante in variantes:
        for c_lower, c_original in columnas_limpias.items():
            if variante in c_lower:
                return c_original
    if default_idx < len(df.columns):
        return df.columns[default_idx]
    return df.columns[0]

EXCEL_FILENAME = "alarmas_inland.xlsx"

@st.cache_data
def cargar_datos():
    if os.path.exists(EXCEL_FILENAME):
        df = pd.read_excel(EXCEL_FILENAME)
    else:
        df = pd.DataFrame({
            "FM": ["FM0", "FM1001", "FM2015", "FM3050"],
            "Description": ["WTG System ok", "Sobrevelocidad de rotor", "Torsión excesiva de cables", "Temperatura generador alta"],
            "Reset Mode": ["-", "SL", "A", "M"],
            "Brake-Program": [0, 3, 0, 1],
            "Yaw-Program": [0, 29, 28, 0],
            "StatCom": ["Yes", "No", "Yes", "Yes"],
            "Reduce Availability": ["No", "Yes", "No", "Yes"],
            "Reset Level": ["-", "Nivel 3", "Operador", "Nivel 2"],
            "Deactivation Level": ["-", "Nivel 4", "Nivel 2", "Nivel 3"]
        })
    df.columns = [str(c).strip() for c in df.columns]
    return df

df = cargar_datos()

# Asignar nombres reales de columnas encontrados en el Excel
COL_ID = detectar_columna(df, ["codigo", "código", "fm", "code", "alarm", "numero", "id"], 0)
COL_DESC = detectar_columna(df, ["descripcion", "descripción", "description", "message", "texto", "falla"], 1)
COL_RESET = detectar_columna(df, ["reset mode", "reset_mode", "resetmode", "reset"], 2)
COL_BRAKE = detectar_columna(df, ["brake-program", "brake_program", "brakeprogram", "brake", "freno"], 3)
COL_YAW = detectar_columna(df, ["yaw-program", "yaw_program", "yawprogram", "yaw", "orientacion"], 4)
COL_STATCOM = detectar_columna(df, ["statcom", "stat_com", "reactiva"], 5)
COL_DISP = detectar_columna(df, ["reduce availability", "reduce_availability", "disponibilidad", "availability"], 6)
COL_LVL_RESET = detectar_columna(df, ["reset level", "reset_level", "nivel reset"], 7)
COL_LVL_DEACT = detectar_columna(df, ["deactivation level", "deactivation_level", "nivel desact"], 8)

# -------------------------------------------------------------
# ENCABEZADO ECO-TECNOLÓGICO
# -------------------------------------------------------------
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">
        <span>🌱 DICCIONARIO DE ALARMAS | CENTRALES EÓLICAS - INLAND</span>
    </div>
    <div class="hero-subtitle">
        Decodificador operativo inteligente de códigos de falla, programas de frenado, orientación y rearme en aerogeneradores.
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# BARRA DE BÚSQUEDA Y FILTRO RÁPIDO (CENTRAL Y PROTAGONISTA)
# -------------------------------------------------------------
col_busq, col_filtro = st.columns([3, 1])

with col_busq:
    busqueda = st.text_input(
        "🔎 Buscador de Alarmas / Fallas:",
        placeholder="Escribe el código (ej. FM120, 1001) o palabras clave (ej. Rotor, Presión, Yaw, Sensor)...",
        label_visibility="collapsed"
    )

with col_filtro:
    opcion_filtro = st.selectbox(
        "Impacto:",
        ["Todas las Alarmas", "⚠️ Solo con Parada (Disponibilidad)", "✅ Solo Informativas / Operativas"],
        label_visibility="collapsed"
    )

# Filtrado de DataFrame
df_filtrado = df.copy()

if busqueda:
    filtro_texto = df_filtrado.astype(str).apply(
        lambda row: row.str.contains(busqueda, case=False, na=False)
    ).any(axis=1)
    df_filtrado = df_filtrado[filtro_texto]

if opcion_filtro == "⚠️ Solo con Parada (Disponibilidad)":
    df_filtrado = df_filtrado[df_filtrado[COL_DISP].astype(str).str.strip().str.upper().isin(["YES", "SÍ", "SI"])]
elif opcion_filtro == "✅ Solo Informativas / Operativas":
    df_filtrado = df_filtrado[~df_filtrado[COL_DISP].astype(str).str.strip().str.upper().isin(["YES", "SÍ", "SI"])]

# Cálculos para KPIs
total_base = len(df)
coincidentes = len(df_filtrado)
criticas_base = len(df[df[COL_DISP].astype(str).str.strip().str.upper().isin(["YES", "SÍ", "SI"])])

# Render de KPIs
st.markdown(f"""
<div class="kpi-container">
    <div class="kpi-card">
        <div class="kpi-num">{coincidentes}</div>
        <div class="kpi-label">Resultados Visibles</div>
    </div>
    <div class="kpi-card kpi-card-crit">
        <div class="kpi-num">{criticas_base}</div>
        <div class="kpi-label">Alarmas que paran Turbina (Base)</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-num">{total_base}</div>
        <div class="kpi-label">Total en Catálogo INLAND</div>
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# TABS PRINCIPALES
# -------------------------------------------------------------
tab_ficha, tab_catalogo, tab_leyenda = st.tabs([
    "🎯 Ficha Operativa Interactiva",
    "📋 Base de Datos Tabular",
    "📖 Leyenda de Especificaciones"
])

# ------------------ PESTAÑA 1: FICHA AMIGABLE ------------------
with tab_ficha:
    if len(df_filtrado) == 0:
        st.warning("No se encontraron fallas con los términos buscados. Intenta con otra palabra clave.")
    else:
        # Selector de alarma estilizado
        etiquetas = df_filtrado[COL_ID].astype(str) + "  —  " + df_filtrado[COL_DESC].astype(str)
        seleccion = st.selectbox(
            "Selecciona la alarma a decodificar:",
            etiquetas,
            label_visibility="visible"
        )
        
        fila_idx = etiquetas[etiquetas == seleccion].index[0]
        alarma = df_filtrado.loc[fila_idx]

        # Extraer variables con limpieza
        val_id = str(alarma[COL_ID]).strip()
        val_desc = str(alarma[COL_DESC]).strip()
        val_reset = str(alarma.get(COL_RESET, "-")).strip()
        
        try:
            val_brake = int(alarma.get(COL_BRAKE, 0))
        except:
            val_brake = -1
            
        try:
            val_yaw = int(alarma.get(COL_YAW, 0))
        except:
            val_yaw = -1
            
        val_disp = str(alarma.get(COL_DISP, "No")).strip().upper() in ["YES", "SÍ", "SI"]
        val_statcom = str(alarma.get(COL_STATCOM, "No")).strip().upper() in ["YES", "SÍ", "SI"]
        val_lvl_reset = str(alarma.get(COL_LVL_RESET, "No especificado")).strip()
        val_lvl_deact = str(alarma.get(COL_LVL_DEACT, "No especificado")).strip()

        # Decodificaciones
        desc_reset = RESET_MODE_MAP.get(val_reset, "Modo particular no catalogado")
        desc_brake = BRAKE_PROG_MAP.get(val_brake, f"Programa estándar / código {val_brake}")
        desc_yaw = YAW_PROG_MAP.get(val_yaw, f"Programa estándar / código {val_yaw}")

        # Badge de disponibilidad
        badge_disp = """<span class="badge-stopped">⚠️ PARADA FORZADA (Afecta Disponibilidad)</span>""" if val_disp else """<span class="badge-running">✅ TURBINA OPERATIVA (No reduce disponibilidad)</span>"""

        # RENDER DE LA FICHA
        st.markdown(f"""
        <div class="alarm-card">
            <div class="alarm-header">
                <div>
                    <span class="alarm-code">{val_id}</span>
                    <span style="font-size: 20px; font-weight: 600; color: #ffffff; margin-left: 12px;">{val_desc}</span>
                </div>
                <div>{badge_disp}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.markdown(f"""
            <div class="prop-box" style="border-top: 3px solid #f59e0b;">
                <div class="prop-title">🛑 Programa de Frenado ({val_brake})</div>
                <div class="prop-value">{desc_brake}</div>
            </div>
            """, unsafe_allow_html=True)

        with c2:
            st.markdown(f"""
            <div class="prop-box" style="border-top: 3px solid #3b82f6;">
                <div class="prop-title">🔄 Orientación / Yaw ({val_yaw})</div>
                <div class="prop-value">{desc_yaw}</div>
            </div>
            """, unsafe_allow_html=True)

        with c3:
            st.markdown(f"""
            <div class="prop-box" style="border-top: 3px solid #10b981;">
                <div class="prop-title">🔑 Modo Restablecimiento ({val_reset})</div>
                <div class="prop-value">{desc_reset}</div>
            </div>
            """, unsafe_allow_html=True)

        with c4:
            txt_statcom = "Habilitado durante parada" if val_statcom else "Inhibido por alarma"
            color_statcom = "#10b981" if val_statcom else "#94a3b8"
            st.markdown(f"""
            <div class="prop-box" style="border-top: 3px solid {color_statcom};">
                <div class="prop-title">⚡ Suministro StatCom</div>
                <div class="prop-value">{txt_statcom}</div>
            </div>
            """, unsafe_allow_html=True)

        # Permisos y niveles
        st.markdown(f"""
        <div style="background: #0f172a; padding: 12px 16px; border-radius: 8px; font-size: 13px; color: #94a3b8; border: 1px solid #1e293b;">
            🛡️ <b>Nivel de usuario para Reset:</b> <span style="color:#e2e8f0;">{val_lvl_reset}</span> &nbsp;&nbsp;|&nbsp;&nbsp; 
            🔒 <b>Nivel para Desactivación:</b> <span style="color:#e2e8f0;">{val_lvl_deact}</span>
        </div>
        """, unsafe_allow_html=True)

# ------------------ PESTAÑA 2: TABLA COMPLETA ------------------
with tab_catalogo:
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
    csv = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        "📥 Descargar registros filtrados (CSV)",
        data=csv,
        file_name="alarmas_inland_export.csv",
        mime="text/csv"
    )

# ------------------ PESTAÑA 3: LEYENDA TÉCNICA ------------------
with tab_leyenda:
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        st.markdown("#### 🔑 Modos de Restablecimiento (*Reset Mode*)")
        st.dataframe(pd.DataFrame(list(RESET_MODE_MAP.items()), columns=["Modo", "Significado"]), hide_index=True, use_container_width=True)
        
        st.markdown("#### 🔄 Programas de Orientación (*Yaw-Program*)")
        st.dataframe(pd.DataFrame(list(YAW_PROG_MAP.items()), columns=["Código", "Significado"]), hide_index=True, use_container_width=True)

    with col_l2:
        st.markdown("#### 🛑 Programas de Frenado (*Brake-Program*)")
        st.dataframe(pd.DataFrame(list(BRAKE_PROG_MAP.items()), columns=["Código", "Significado"]), hide_index=True, use_container_width=True)
        
        st.markdown("""
        #### ⚡ Propiedades Adicionales
        * **StatCom:**
          * `Yes`: Suministro de potencia reactiva posible durante la parada del aerogenerador.
          * `No`: Suministro de potencia reactiva inhibido por la alarma.
        * **Reduce Availability:**
          * `Yes`: Afecta la disponibilidad técnica (turbina parada).
          * `No`: La turbina continúa operando.
        """)
