import streamlit as st
import pandas as pd
import os

# Configuración de página
st.set_page_config(
    page_title="Diccionario de Alarmas | INLAND",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos visuales personalizados
st.markdown("""
<style>
    .metric-card {
        background-color: #1e2129;
        border-radius: 8px;
        padding: 15px;
        border-left: 5px solid #ff4b4b;
        margin-bottom: 10px;
    }
    .badge-critical {
        background-color: #ff4b4b;
        color: white;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: bold;
    }
    .badge-ok {
        background-color: #00c04b;
        color: white;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# DICCIONARIOS DE DECODIFICACIÓN (Basados en la leyenda técnica)
# -------------------------------------------------------------
RESET_MODE_MAP = {
    "-": "Sin modo de restablecimiento",
    "A": "Restablecimiento automático de alarma",
    "M": "Restablecimiento manual requerido",
    "M(A)": "Restablecimiento manual (Automático por ARS)",
    "ML": "Restablecimiento manual local en turbina por operador",
    "SL": "Restablecimiento local de alarma de seguridad requerido",
    "SR": "Restablecimiento remoto de alarma de seguridad (Solo K08Delta)"
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
    24: "Modo automático 90°",
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
# CARGA DE DATOS
# -------------------------------------------------------------
EXCEL_FILENAME = "alarmas_inland.xlsx"

@st.cache_data
def load_data():
    if os.path.exists(EXCEL_FILENAME):
        df = pd.read_excel(EXCEL_FILENAME)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    else:
        # Datos de prueba en caso de que aún no cargues el Excel definitivo
        data_demo = {
            "Codigo": [1001, 1002, 2015, 3050, 4120],
            "Descripcion": [
                "Sobrevelocidad en rotor", 
                "Fallo en presión hidráulica de frenos", 
                "Torsión excesiva de cables en góndola", 
                "Temperatura alta en generador", 
                "Fallo de comunicación StatCom"
            ],
            "Reset_Mode": ["SL", "ML", "A", "M", "A"],
            "Brake_Program": [3, 2, 0, 1, 0],
            "Yaw_Program": [29, 0, 28, 0, 0],
            "StatCom": ["No", "No", "Yes", "Yes", "No"],
            "Reduce_Availability": ["Yes", "Yes", "No", "Yes", "No"],
            "Reset_Level": ["Nivel 3", "Nivel 2", "Operador", "Nivel 2", "Operador"],
            "Deactivation_Level": ["Nivel 4", "Nivel 3", "Nivel 2", "Nivel 3", "Nivel 2"]
        }
        return pd.DataFrame(data_demo)

df = load_data()

# Título y encabezado
st.title("⚡ Diccionario de Alarmas | Centrales Eólicas - INLAND")
st.markdown("Plataforma técnica de consulta y decodificación operativa de alarmas y paradas de aerogeneradores.")

# -------------------------------------------------------------
# BARRA LATERAL (FILTROS)
# -------------------------------------------------------------
st.sidebar.header("🔍 Filtros de Búsqueda")

# Búsqueda por texto / código
busqueda = st.sidebar.text_input("Buscar por Código o Descripción:", placeholder="Ej. 1001 o Rotor")

# Filtro por Impacto en Disponibilidad
if "Reduce_Availability" in df.columns:
    disp_options = ["Todos"] + sorted(list(df["Reduce_Availability"].dropna().unique()))
    filtro_disp = st.sidebar.selectbox("Afecta Disponibilidad:", disp_options)
else:
    filtro_disp = "Todos"

# Filtro por Reset Mode
if "Reset_Mode" in df.columns:
    reset_options = ["Todos"] + sorted(list(df["Reset_Mode"].dropna().unique()))
    filtro_reset = st.sidebar.selectbox("Modo de Restablecimiento (Reset):", reset_options)
else:
    filtro_reset = "Todos"

# Aplicar filtros
df_filtrado = df.copy()

if busqueda:
    filtro_query = (
        df_filtrado.astype(str).apply(lambda row: row.str.contains(busqueda, case=False)).any(axis=1)
    )
    df_filtrado = df_filtrado[filtro_query]

if filtro_disp != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Reduce_Availability"] == filtro_disp]

if filtro_reset != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Reset_Mode"] == filtro_reset]

# -------------------------------------------------------------
# MÉTRICAS RÁPIDAS
# -------------------------------------------------------------
col1, col2, col3 = st.columns(3)
col1.metric("Alarmas Coincidentes", len(df_filtrado))
if "Reduce_Availability" in df_filtrado.columns:
    criticas = len(df_filtrado[df_filtrado["Reduce_Availability"].astype(str).str.upper() == "YES"])
    col2.metric("Alarmas con Pérdida de Disponibilidad", criticas)
col3.metric("Total en Base de Datos", len(df))

st.markdown("---")

# -------------------------------------------------------------
# PESTAÑAS PRINCIPALES
# -------------------------------------------------------------
tab_consulta, tab_tabla, tab_leyenda = st.tabs([
    "🎯 Ficha y Decodificador Rápido", 
    "📋 Matriz General de Alarmas", 
    "📖 Leyenda de Propiedades Técnica"
])

# PESTAÑA 1: DECODIFICADOR RÁPIDO
with tab_consulta:
    st.subheader("Consulta Detallada de Falla")
    
    if len(df_filtrado) == 0:
        st.warning("No se encontraron alarmas con los filtros especificados.")
    else:
        # Selector de alarma para ver en detalle
        col_id = "Codigo" if "Codigo" in df_filtrado.columns else df_filtrado.columns[0]
        col_desc = "Descripcion" if "Descripcion" in df_filtrado.columns else df_filtrado.columns[1]
        
        opciones_alarmas = df_filtrado[col_id].astype(str) + " - " + df_filtrado[col_desc].astype(str)
        seleccion = st.selectbox("Seleccione el código a inspeccionar:", opciones_alarmas)
        
        idx_seleccionado = opciones_alarmas[opciones_alarmas == seleccion].index[0]
        alarma = df_filtrado.loc[idx_seleccionado]
        
        # Panel de visualización de la alarma
        st.markdown(f"### 🚨 Código: `{alarma.get('Codigo', 'N/A')}` — {alarma.get('Descripcion', 'Sin descripción')}")
        
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown("#### ⚙️ Parámetros de Seguridad y Freno")
            
            # Reset
            r_val = str(alarma.get("Reset_Mode", "-")).strip()
            r_desc = RESET_MODE_MAP.get(r_val, "No especificado")
            st.write(f"**Modo de Reset (`{r_val}`):** {r_desc}")
            
            # Brake
            try:
                b_val = int(alarma.get("Brake-Program", alarma.get("Brake_Program", -1)))
            except:
                b_val = -1
            b_desc = BRAKE_PROG_MAP.get(b_val, "No especificado o estándar")
            st.write(f"**Programa de Frenado (`{b_val}`):** {b_desc}")
            
            # Yaw
            try:
                y_val = int(alarma.get("Yaw-Program", alarma.get("Yaw_Program", -1)))
            except:
                y_val = -1
            y_desc = YAW_PROG_MAP.get(y_val, "No especificado o estándar")
            st.write(f"**Programa de Orientación / Yaw (`{y_val}`):** {y_desc}")

        with c2:
            st.markdown("#### ⚡ Disponibilidad e Impacto Eléctrico")
            
            # Disponibilidad
            red_disp = str(alarma.get("Reduce Availability", alarma.get("Reduce_Availability", "No"))).strip()
            if red_disp.upper() in ["YES", "SÍ", "SI"]:
                st.error("⚠️ **Reduce Disponibilidad:** SÍ (Turbina detenida / parada forzada)")
            else:
                st.success("✅ **Reduce Disponibilidad:** NO (Generación / servicio no interrumpido)")
                
            # StatCom
            statcom = str(alarma.get("StatCom", "No")).strip()
            st.write(f"**Suministro Potencia Reactiva (StatCom):** {'Permitido' if statcom.upper() == 'YES' else 'Inhibido por alarma'}")
            
            # Niveles
            st.write(f"**Nivel de usuario para Reset:** {alarma.get('Reset Level', alarma.get('Reset_Level', 'No requerido'))}")
            st.write(f"**Nivel de usuario para Desactivación:** {alarma.get('Deactivation Level', alarma.get('Deactivation_Level', 'No requerido'))}")

# PESTAÑA 2: MATRIZ COMPLETA
with tab_tabla:
    st.subheader("Registros en Base de Datos")
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
    
    # Botón de descarga
    csv = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Descargar datos filtrados (CSV)",
        data=csv,
        file_name="alarmas_filtradas_inland.csv",
        mime="text/csv"
    )

# PESTAÑA 3: LEYENDA TÉCNICA (DE TU IMAGEN)
with tab_leyenda:
    st.subheader("📖 Leyenda de Propiedades de Alarma")
    
    c_ley1, c_ley2 = st.columns(2)
    
    with c_ley1:
        st.markdown("**Modos de Restablecimiento (Reset Mode)**")
        df_reset = pd.DataFrame(list(RESET_MODE_MAP.items()), columns=["Código", "Significado"])
        st.table(df_reset)
        
        st.markdown("**Programas de Orientación (Yaw-Program)**")
        df_yaw = pd.DataFrame(list(YAW_PROG_MAP.items()), columns=["Código", "Significado"])
        st.table(df_yaw)
        
    with c_ley2:
        st.markdown("**Programas de Frenado (Brake-Program)**")
        df_brake = pd.DataFrame(list(BRAKE_PROG_MAP.items()), columns=["Código", "Significado"])
        st.table(df_brake)
        
        st.markdown("""
        **Otras Propiedades:**
        * **StatCom:**
          * `Yes`: Suministro de potencia reactiva posible durante la parada del aerogenerador.
          * `No`: Suministro de potencia reactiva inhibido por la alarma.
        * **Reduce Availability:**
          * `Yes`: Reduce la disponibilidad del aerogenerador si la alarma está activa.
          * `No`: La disponibilidad no se reduce.
        * **Reset Level:** Nivel de usuario requerido para restablecer la alarma.
        * **Deactivation Level:** Nivel de usuario requerido para desactivar la alarma.
        """)
