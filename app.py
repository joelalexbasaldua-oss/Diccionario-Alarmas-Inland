import streamlit as st
import pandas as pd
import os

# Configuración de página
st.set_page_config(
    page_title="Terminal SCADA | Diccionario de Alarmas INLAND",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -------------------------------------------------------------
# ESTILOS VISUALES TIPO TERMINAL SCADA (MODERNO & DARK)
# -------------------------------------------------------------
st.markdown("""
<style>
    /* Fondo principal y fuentes */
    .stApp {
        background-color: #0b0f19;
        color: #e2e8f0;
    }
    
    /* Terminal Header */
    .scada-header {
        background: linear-gradient(90deg, #091e3a 0%, #0d3b66 50%, #051c2c 100%);
        border: 1px solid #1e40af;
        border-radius: 12px;
        padding: 20px 25px;
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(14, 165, 233, 0.15);
    }
    .scada-title {
        font-size: 24px;
        font-weight: 800;
        letter-spacing: 0.5px;
        color: #38bdf8;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .scada-subtitle {
        font-size: 13px;
        color: #94a3b8;
        margin-top: 4px;
    }

    /* Tarjeta Principal de la Alarma */
    .main-alarm-box {
        background: #111827;
        border-radius: 14px;
        border: 1px solid #1f2937;
        padding: 24px;
        margin-top: 15px;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5);
    }
    .code-badge {
        font-size: 26px;
        font-weight: 900;
        font-family: 'Courier New', monospace;
        color: #38bdf8;
        background: #082f49;
        border: 1px solid #0284c7;
        padding: 6px 16px;
        border-radius: 8px;
        display: inline-block;
    }
    .alarm-desc-text {
        font-size: 18px;
        font-weight: 600;
        color: #f8fafc;
        line-height: 1.5;
        margin-top: 15px;
        padding: 15px;
        background: #162032;
        border-left: 4px solid #38bdf8;
        border-radius: 6px;
    }

    /* Badges de Impacto Operativo */
    .status-trip {
        background: rgba(220, 38, 38, 0.2);
        border: 1px solid #ef4444;
        color: #fca5a5;
        padding: 8px 18px;
        border-radius: 30px;
        font-weight: 700;
        font-size: 13px;
        letter-spacing: 0.5px;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }
    .status-ok {
        background: rgba(16, 185, 129, 0.2);
        border: 1px solid #10b981;
        color: #86efac;
        padding: 8px 18px;
        border-radius: 30px;
        font-weight: 700;
        font-size: 13px;
        letter-spacing: 0.5px;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }

    /* Cajas Modulares de Parámetros */
    .module-card {
        background: #131c2e;
        border: 1px solid #1e293b;
        border-radius: 10px;
        padding: 16px;
        min-height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .module-label {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.8px;
    }
    .module-code {
        font-size: 18px;
        font-weight: 800;
        color: #f1f5f9;
        margin: 6px 0;
    }
    .module-meaning {
        font-size: 13px;
        font-weight: 500;
        color: #cbd5e1;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# DICCIONARIOS OFICIALES DE DECODIFICACIÓN
# -------------------------------------------------------------
RESET_MODE_MAP = {
    "-": "Sin modo de restablecimiento",
    "A": "Restablecimiento automático de alarma",
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
# CARGA Y MAPEO INTELIGENTE DE COLUMNAS (POR CONTENIDO REAL)
# -------------------------------------------------------------
EXCEL_FILENAME = "alarmas_inland.xlsx"

@st.cache_data
def cargar_dataset():
    if os.path.exists(EXCEL_FILENAME):
        df = pd.read_excel(EXCEL_FILENAME)
    else:
        df = pd.DataFrame({
            "FM": ["FM0", "FM50", "FM51", "FM1001"],
            "Description": [
                "WTG System ok",
                "Cadena de seguridad parada de emergencia general",
                "Cadena de seguridad activada a través del interruptor Bottombox",
                "Sobrevelocidad en rotor"
            ],
            "Reset Mode": ["-", "SL", "ML", "SL"],
            "Brake-Program": [0, 3, 3, 3],
            "Yaw-Program": [0, 50, 50, 29],
            "StatCom": ["Yes", "No", "No", "No"],
            "Reduce Availability": ["No", "Yes", "Yes", "Yes"],
            "Reset Level": ["-", "Operador", "Nivel 2", "Nivel 3"],
            "Deactivation Level": ["-", "Nivel 2", "Nivel 3", "Nivel 4"]
        })
    df.columns = [str(c).strip() for c in df.columns]
    return df

df = cargar_dataset()

# Identificar columnas por contenido para evitar confusiones
col_id = None
col_desc = None
col_reset = None
col_brake = None
col_yaw = None
col_statcom = None
col_disp = None
col_lvl_reset = None
col_lvl_deact = None

for c in df.columns:
    c_lower = c.lower()
    
    # Columna Código / FM
    if any(k in c_lower for k in ["fm", "codigo", "código", "code", "alarm"]):
        if col_id is None:
            col_id = c
            
    # Columna Descripción (textos largos)
    elif any(k in c_lower for k in ["descrip", "message", "mensaje", "texto", "falla"]):
        if col_desc is None:
            col_desc = c
            
    # Columna Reset Mode
    elif "reset mode" in c_lower or "modo reset" in c_lower:
        col_reset = c
    elif "reset level" in c_lower or "nivel reset" in c_lower:
        col_lvl_reset = c
    elif "deactivation" in c_lower or "desactiv" in c_lower:
        col_lvl_deact = c
    elif "brake" in c_lower or "freno" in c_lower:
        col_brake = c
    elif "yaw" in c_lower or "orienta" in c_lower:
        col_yaw = c
    elif "statcom" in c_lower or "reactiv" in c_lower:
        col_statcom = c
    elif "availab" in c_lower or "disponib" in c_lower:
        col_disp = c

# Asignaciones por descarte seguro
if col_id is None: col_id = df.columns[0]
if col_desc is None:
    # Busca la columna con mayor longitud promedio de texto
    col_desc = max(df.columns, key=lambda col: df[col].astype(str).str.len().mean())

# Crear clave de búsqueda normalizada (ej: '51' o 'FM51' -> '51')
def limpiar_codigo(val):
    s = str(val).strip().upper().replace(" ", "")
    if s.startswith("FM"):
        return s[2:]
    return s

df["_CODIGO_KEY"] = df[col_id].apply(limpiar_codigo)
df["_CODIGO_LABEL"] = df[col_id].astype(str).str.strip()

# -------------------------------------------------------------
# CABECERA TERMINAL SCADA
# -------------------------------------------------------------
st.markdown("""
<div class="scada-header">
    <div class="scada-title">
        <span>⚡ TERMINAL DE CONSULTA DE ALARMAS | CENTRALES EÓLICAS - INLAND</span>
    </div>
    <div class="scada-subtitle">
        Buscador directo de ingeniería y operaciones para decodificación inmediata de fallas en aerogeneradores.
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# BUSCADOR DIRECTO (INPUT PRINCIPAL)
# -------------------------------------------------------------
c_input, c_select = st.columns([2, 1])

with c_input:
    codigo_ingresado = st.text_input(
        "🔎 Ingrese el Código de Falla exacto (ej. 51 o FM51):",
        value="",
        placeholder="Escriba aquí el número o código (ej: FM51, 51, 0, 1001)...",
        help="Escriba el código exacto y presione Enter."
    )

with c_select:
    # Selector directo como alternativa rápida ordenada
    todos_los_codigos = ["-- Seleccionar de la lista --"] + sorted(list(df["_CODIGO_LABEL"].unique()))
    seleccion_lista = st.selectbox("O elija directamente el código:", todos_los_codigos)

# Determinar qué código consultar
codigo_a_buscar = ""
if codigo_ingresado.strip():
    codigo_a_buscar = limpiar_codigo(codigo_ingresado)
elif seleccion_lista != "-- Seleccionar de la lista --":
    codigo_a_buscar = limpiar_codigo(seleccion_lista)

# Botones rápidos de acceso frecuente
st.markdown("**Accesos Rápidos:**")
col_chips = st.columns(6)
chips_ejemplo = ["FM0", "FM50", "FM51", "FM52", "FM100", "FM1001"]
for i, chip in enumerate(chips_ejemplo):
    if col_chips[i].button(chip, use_container_width=True):
        codigo_a_buscar = limpiar_codigo(chip)

st.markdown("---")

# -------------------------------------------------------------
# RESULTADO DE LA CONSULTA (EXACTA)
# -------------------------------------------------------------
if not codigo_a_buscar:
    st.info("💡 **Listo para consultar:** Ingrese un código en la barra superior o haga clic en un acceso rápido.")
else:
    # Búsqueda EXACTA
    resultado = df[df["_CODIGO_KEY"] == codigo_a_buscar]
    
    if len(resultado) == 0:
        st.error(f"❌ **Código no encontrado:** No existe ninguna alarma con el identificador `{codigo_ingresado}` en el catálogo de INLAND.")
        st.caption("Verifique si ingresó correctamente el número o elija un código del menú desplegable a la derecha.")
    else:
        # Tomar exactamente el registro encontrado
        alarma = resultado.iloc[0]
        
        # Extracción segura de datos
        val_id = str(alarma.get(col_id, "N/A")).strip()
        val_desc = str(alarma.get(col_desc, "Sin descripción disponible")).strip()
        
        # Reset Mode
        val_reset = str(alarma.get(col_reset, "-")).strip() if col_reset else "-"
        desc_reset = RESET_MODE_MAP.get(val_reset, val_reset if val_reset != "nan" else "No especificado")
        
        # Brake
        val_brake = str(alarma.get(col_brake, "-")).strip() if col_brake else "-"
        try:
            b_int = int(float(val_brake))
            desc_brake = BRAKE_PROG_MAP.get(b_int, f"Programa {b_int}")
        except:
            desc_brake = val_brake if val_brake != "nan" else "Estándar"
            
        # Yaw
        val_yaw = str(alarma.get(col_yaw, "-")).strip() if col_yaw else "-"
        try:
            y_int = int(float(val_yaw))
            desc_yaw = YAW_PROG_MAP.get(y_int, f"Programa {y_int}")
        except:
            desc_yaw = val_yaw if val_yaw != "nan" else "Estándar"
            
        # Disponibilidad
        val_disp_raw = str(alarma.get(col_disp, "No")).strip().upper() if col_disp else "NO"
        es_parada = val_disp_raw in ["YES", "SÍ", "SI", "TRUE", "1"]
        
        # StatCom
        val_statcom_raw = str(alarma.get(col_statcom, "No")).strip().upper() if col_statcom else "NO"
        es_statcom = val_statcom_raw in ["YES", "SÍ", "SI", "TRUE", "1"]
        
        # Niveles
        lvl_reset = str(alarma.get(col_lvl_reset, "No especificado")).strip() if col_lvl_reset else "No especificado"
        lvl_deact = str(alarma.get(col_lvl_deact, "No especificado")).strip() if col_lvl_deact else "No especificado"
        
        # Render del Badge de Estado
        if es_parada:
            badge_html = '<div class="status-trip">🛑 PARADA FORZADA — REDUCE DISPONIBILIDAD</div>'
        else:
            badge_html = '<div class="status-ok">🟢 TURBINA OPERATIVA — NO AFECTA DISPONIBILIDAD</div>'

        # FICHA PRINCIPAL
        st.markdown(f"""
        <div class="main-alarm-box">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <div>
                    <span class="code-badge">{val_id}</span>
                </div>
                <div>
                    {badge_html}
                </div>
            </div>
            <div class="alarm-desc-text">
                {val_desc}
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # CUADRÍCULA DE DIAGNÓSTICO TÉCNICO (4 MÓDULOS)
        m1, m2, m3, m4 = st.columns(4)
        
        with m1:
            st.markdown(f"""
            <div class="module-card" style="border-top: 3px solid #ef4444;">
                <div class="module-label">🛑 Frenado (Brake)</div>
                <div class="module-code">Código: {val_brake}</div>
                <div class="module-meaning">{desc_brake}</div>
            </div>
            """, unsafe_allow_html=True)

        with m2:
            st.markdown(f"""
            <div class="module-card" style="border-top: 3px solid #38bdf8;">
                <div class="module-label">🔄 Orientación (Yaw)</div>
                <div class="module-code">Código: {val_yaw}</div>
                <div class="module-meaning">{desc_yaw}</div>
            </div>
            """, unsafe_allow_html=True)

        with m3:
            st.markdown(f"""
            <div class="module-card" style="border-top: 3px solid #10b981;">
                <div class="module-label">🔑 Rearme (Reset Mode)</div>
                <div class="module-code">Modo: {val_reset}</div>
                <div class="module-meaning">{desc_reset}</div>
            </div>
            """, unsafe_allow_html=True)

        with m4:
            txt_stc = "Habilitado en parada" if es_statcom else "Inhibido por alarma"
            st.markdown(f"""
            <div class="module-card" style="border-top: 3px solid {'#10b981' if es_statcom else '#64748b'};">
                <div class="module-label">⚡ Potencia Reactiva</div>
                <div class="module-code">StatCom</div>
                <div class="module-meaning">{txt_stc}</div>
            </div>
            """, unsafe_allow_html=True)

        # Barra de permisos de seguridad
        st.markdown(f"""
        <div style="background: #111827; border: 1px solid #1f2937; border-radius: 8px; padding: 12px 18px; margin-top: 15px; font-size: 13px; color: #94a3b8;">
            🛡️ <b>Nivel de Usuario para Reset:</b> <span style="color: #f1f5f9;">{lvl_reset}</span> &nbsp;&nbsp;|&nbsp;&nbsp; 
            🔒 <b>Nivel para Desactivación:</b> <span style="color: #f1f5f9;">{lvl_deact}</span>
        </div>
        """, unsafe_allow_html=True)

        # Desplegable con todos los campos en bruto (para auditoría técnica)
        with st.expander("🔍 Ver todos los campos originales del archivo Excel"):
            st.dataframe(resultado.drop(columns=["_CODIGO_KEY", "_CODIGO_LABEL"], errors="ignore"), use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# PESTAÑA INFERIOR: LEYENDA OFICIAL
# -------------------------------------------------------------
with st.expander("📖 Consultar Leyenda Oficial de Códigos (Frenos, Yaw y Reset)"):
    c_tab1, c_tab2 = st.columns(2)
    with c_tab1:
        st.markdown("##### 🔑 Modos de Restablecimiento (*Reset Mode*)")
        st.dataframe(pd.DataFrame(list(RESET_MODE_MAP.items()), columns=["Modo", "Significado"]), hide_index=True, use_container_width=True)
        st.markdown("##### 🔄 Programas de Orientación (*Yaw-Program*)")
        st.dataframe(pd.DataFrame(list(YAW_PROG_MAP.items()), columns=["Código", "Significado"]), hide_index=True, use_container_width=True)
    with c_tab2:
        st.markdown("##### 🛑 Programas de Frenado (*Brake-Program*)")
        st.dataframe(pd.DataFrame(list(BRAKE_PROG_MAP.items()), columns=["Código", "Significado"]), hide_index=True, use_container_width=True)
