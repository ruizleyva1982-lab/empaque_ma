import streamlit as st
import pandas as pd
from datetime import date
import os
import json

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Registro de Cajas Empacadas",
    page_icon="📦",
    layout="wide",
)

DATA_FILE = "data.xlsx"
REGISTRO_FILE = "registro_cajas.json"
OPERARIOS_FILE = "operarios.json"

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* Header */
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    }
    .main-header h1 {
        color: #e94560;
        font-size: 2rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: 1px;
    }
    .main-header p {
        color: #a8b2d8;
        margin: 0.3rem 0 0 0;
        font-size: 0.95rem;
    }

    /* Cards de métricas */
    .metric-card {
        background: linear-gradient(135deg, #0f3460, #16213e);
        border: 1px solid #e94560;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        text-align: center;
        box-shadow: 0 2px 12px rgba(233,69,96,0.2);
    }
    .metric-card .value {
        font-size: 2rem;
        font-weight: 800;
        color: #e94560;
    }
    .metric-card .label {
        font-size: 0.8rem;
        color: #a8b2d8;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    /* Producto seleccionado */
    .product-card {
        background: linear-gradient(135deg, #0d1b2a, #1b263b);
        border: 2px solid #4cc9f0;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        margin: 0.8rem 0;
    }
    .product-card .prod-code {
        font-size: 0.8rem;
        color: #4cc9f0;
        font-weight: 700;
        letter-spacing: 2px;
        text-transform: uppercase;
    }
    .product-card .prod-name {
        font-size: 1.2rem;
        font-weight: 700;
        color: #f0f4f8;
        margin: 0.2rem 0;
    }
    .product-card .prod-factor {
        font-size: 0.9rem;
        color: #a8b2d8;
    }
    .prod-factor span {
        color: #f6c90e;
        font-weight: 700;
        font-size: 1rem;
    }

    /* Tabla de registro */
    .registro-header {
        background: linear-gradient(90deg, #e94560, #c1121f);
        color: white;
        padding: 0.7rem 1rem;
        border-radius: 8px 8px 0 0;
        font-weight: 700;
        font-size: 1.1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    /* Botones */
    div.stButton > button {
        border-radius: 8px;
        font-weight: 700;
        transition: all 0.2s;
    }

    /* Tabla */
    .dataframe-container {
        border-radius: 0 0 8px 8px;
        overflow: hidden;
    }

    /* Separador */
    .divider {
        height: 2px;
        background: linear-gradient(90deg, #e94560, transparent);
        margin: 1.5rem 0;
        border-radius: 2px;
    }

    /* Alerta */
    .stAlert {
        border-radius: 8px;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a1a2e, #0f3460);
    }
    section[data-testid="stSidebar"] .stMarkdown h2 {
        color: #4cc9f0;
    }
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] label {
        color: #a8b2d8 !important;
    }
    section[data-testid="stSidebar"] .stMarkdown strong {
        color: #e94560;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────
@st.cache_data
def cargar_catalogo():
    df = pd.read_excel(DATA_FILE)
    df.columns = df.columns.str.strip()
    df["CÓDIGO"] = df["CÓDIGO"].astype(str).str.strip()
    df["DESCRIPCIÓN PRODUCTO"] = df["DESCRIPCIÓN PRODUCTO"].astype(str).str.strip()
    df["FACTOR"] = pd.to_numeric(df["FACTOR"], errors="coerce").fillna(0)
    return df


def cargar_registro():
    if os.path.exists(REGISTRO_FILE):
        with open(REGISTRO_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def guardar_registro(registro):
    with open(REGISTRO_FILE, "w", encoding="utf-8") as f:
        json.dump(registro, f, ensure_ascii=False, indent=2)


def cargar_operarios():
    if os.path.exists(OPERARIOS_FILE):
        with open(OPERARIOS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def guardar_operarios(operarios):
    with open(OPERARIOS_FILE, "w", encoding="utf-8") as f:
        json.dump(operarios, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────
# ESTADO DE SESIÓN
# ─────────────────────────────────────────────
if "registro" not in st.session_state:
    st.session_state.registro = cargar_registro()

if "operarios" not in st.session_state:
    st.session_state.operarios = cargar_operarios()

if "producto_sel" not in st.session_state:
    st.session_state.producto_sel = None

if "busqueda" not in st.session_state:
    st.session_state.busqueda = ""

# ─────────────────────────────────────────────
# CARGA CATÁLOGO
# ─────────────────────────────────────────────
catalogo = cargar_catalogo()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📦 Registro de Cajas Empacadas</h1>
    <p>María Almenara · Control de Producción · Ingreso por Factor</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR — FILTROS / FECHA
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📅 Fecha de Registro")
    fecha_sel = st.date_input(
        "Selecciona la fecha",
        value=date.today(),
        format="DD/MM/YYYY",
    )
    st.markdown("---")
    st.markdown("## 🔍 Buscar Producto")
    modo_busqueda = st.radio(
        "Filtrar por:",
        ["CÓDIGO", "DESCRIPCIÓN"],
        horizontal=True,
    )
    busqueda = st.text_input(
        "Ingresa texto",
        placeholder="Ej: M6020039 o ALFAJOR",
        key="input_busqueda",
    )

    st.markdown("---")
    st.markdown("## 📊 Resumen de Sesión")

    # Métricas rápidas del registro actual
    reg_df = pd.DataFrame(st.session_state.registro)
    if not reg_df.empty and "fecha" in reg_df.columns:
        fecha_str = fecha_sel.strftime("%Y-%m-%d")
        hoy = reg_df[reg_df["fecha"] == fecha_str]
        st.markdown(f"**Fecha seleccionada:** {fecha_sel.strftime('%d/%m/%Y')}")
        st.markdown(f"**Líneas registradas hoy:** {len(hoy)}")
        total_cajas = hoy["cajas"].sum() if not hoy.empty else 0
        total_unid = hoy["total_unidades"].sum() if not hoy.empty else 0
        st.markdown(f"**Total cajas:** {int(total_cajas):,}")
        st.markdown(f"**Total unidades:** {total_unid:,.2f}")
    else:
        st.markdown("Sin registros aún.")

    st.markdown("---")
    if st.button("🗑️ Limpiar TODO el registro", type="secondary"):
        st.session_state.registro = []
        guardar_registro([])
        st.success("Registro limpiado.")
        st.rerun()

# ─────────────────────────────────────────────
# PANEL PRINCIPAL — BÚSQUEDA Y SELECCIÓN
# ─────────────────────────────────────────────
col_izq, col_der = st.columns([1.2, 1], gap="large")

with col_izq:
    st.markdown("### 🔎 Resultados de búsqueda")

    if busqueda.strip():
        term = busqueda.strip().upper()
        if modo_busqueda == "CÓDIGO":
            resultado = catalogo[catalogo["CÓDIGO"].str.upper().str.contains(term, na=False)]
        else:
            resultado = catalogo[catalogo["DESCRIPCIÓN PRODUCTO"].str.upper().str.contains(term, na=False)]
    else:
        resultado = catalogo.copy()

    if resultado.empty:
        st.warning("⚠️ No se encontraron productos con ese criterio.")
    else:
        st.markdown(f"<small style='color:#a8b2d8'>{len(resultado)} producto(s) encontrado(s)</small>", unsafe_allow_html=True)

        # Tabla seleccionable
        resultado_display = resultado[["CÓDIGO", "DESCRIPCIÓN PRODUCTO", "FACTOR"]].reset_index(drop=True)

        # Agregar columna de selección
        opciones = resultado_display["CÓDIGO"] + " — " + resultado_display["DESCRIPCIÓN PRODUCTO"]
        opciones_list = opciones.tolist()

        seleccion = st.selectbox(
            "Selecciona un producto:",
            options=opciones_list,
            index=None,
            placeholder="Haz clic para elegir...",
        )

        if seleccion:
            codigo_sel = seleccion.split(" — ")[0]
            fila = catalogo[catalogo["CÓDIGO"] == codigo_sel].iloc[0]
            st.session_state.producto_sel = {
                "codigo": fila["CÓDIGO"],
                "descripcion": fila["DESCRIPCIÓN PRODUCTO"],
                "factor": float(fila["FACTOR"]),
            }

        # Tabla de referencia
        st.dataframe(
            resultado_display.style.format({"FACTOR": "{:,.2f}"}),
            use_container_width=True,
            height=300,
            hide_index=True,
        )

with col_der:
    st.markdown("### 📝 Ingreso de Cajas")

    if st.session_state.producto_sel:
        p = st.session_state.producto_sel

        st.markdown(f"""
        <div class="product-card">
            <div class="prod-code">🏷️ {p['codigo']}</div>
            <div class="prod-name">{p['descripcion']}</div>
            <div class="prod-factor">Factor: <span>{p['factor']:,.2f}</span> unidades/caja</div>
        </div>
        """, unsafe_allow_html=True)

        # ── Presentación + Factor real ──────────────
        st.markdown("<small style='color:#a8b2d8;font-weight:600;letter-spacing:1px'>TIPO DE PRESENTACIÓN</small>", unsafe_allow_html=True)
        col_pres, col_fac = st.columns([1, 1])
        with col_pres:
            presentacion = st.selectbox(
                "📦 Presentación",
                options=["Caja", "Jaba", "Bolsa"],
                index=0,
                help="Selecciona el tipo de envase usado hoy",
            )
        with col_fac:
            factor_real = st.number_input(
                "🔢 Factor real",
                min_value=0.01,
                value=float(p["factor"]),
                step=1.0,
                format="%.2f",
                help=f"Factor referencial del producto: {p['factor']:,.2f}. Modifícalo si cambió.",
            )

        # Indicador visual si el factor fue modificado
        if abs(factor_real - p["factor"]) > 0.001:
            st.markdown(
                f"<small style='color:#f6c90e'>⚠️ Factor modificado "
                f"(ref: <b>{p['factor']:,.2f}</b> → actual: <b>{factor_real:,.2f}</b>)</small>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='margin-top:0.4rem'></div>", unsafe_allow_html=True)

        # ── Cantidad ─────────────────────────────────
        cajas = st.number_input(
            f"📦 Cantidad de {presentacion.lower()}s",
            min_value=0.0,
            value=0.0,
            step=1.0,
            format="%.2f",
            help="Ingresa la cantidad empacada",
        )

        if cajas > 0:
            total_preview = cajas * factor_real
            st.markdown(f"""
            <div style='background:#0d2137;border:1px solid #4cc9f0;border-radius:8px;padding:0.8rem 1rem;margin:0.5rem 0;'>
                <span style='color:#a8b2d8;font-size:0.85rem'>TOTAL UNIDADES (preview)</span><br>
                <span style='color:#4cc9f0;font-size:1.8rem;font-weight:800'>{total_preview:,.2f}</span>
                <span style='color:#a8b2d8;font-size:0.85rem'> = {cajas:,.2f} {presentacion.lower()}s × {factor_real:,.2f}</span>
            </div>
            """, unsafe_allow_html=True)

        col_turno, col_op = st.columns(2)
        with col_turno:
            turno = st.selectbox(
                "🕐 Turno",
                options=["Turno Mañana", "Turno Tarde", "Turno Noche"],
                index=0,
            )
        with col_op:
            lista_ops = st.session_state.operarios
            if lista_ops:
                operario = st.selectbox(
                    "👷 Operario",
                    options=["— Seleccionar —"] + lista_ops,
                    index=0,
                )
                operario = "" if operario == "— Seleccionar —" else operario
            else:
                st.selectbox("👷 Operario", options=["Sin operarios registrados"], disabled=True)
                operario = ""
                st.caption("⚠️ Registra operarios en la sección de abajo.")

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("✅ GRABAR", type="primary", use_container_width=True):
                if cajas <= 0:
                    st.error("⚠️ La cantidad debe ser mayor a 0.")
                else:
                    nuevo = {
                        "fecha": fecha_sel.strftime("%Y-%m-%d"),
                        "codigo": p["codigo"],
                        "descripcion": p["descripcion"],
                        "presentacion": presentacion,
                        "factor_ref": p["factor"],
                        "factor": factor_real,
                        "cajas": cajas,
                        "total_unidades": round(cajas * factor_real, 4),
                        "turno": turno,
                        "operario": operario,
                    }
                    st.session_state.registro.append(nuevo)
                    guardar_registro(st.session_state.registro)
                    st.success(f"✅ Registrado: {p['descripcion']} — {cajas:,.2f} cajas → {nuevo['total_unidades']:,.2f} unidades")
                    st.session_state.producto_sel = None
                    st.rerun()

        with col_btn2:
            if st.button("🔄 Limpiar", use_container_width=True):
                st.session_state.producto_sel = None
                st.rerun()

    else:
        st.info("👈 Busca y selecciona un producto para ingresar cajas.")

# ─────────────────────────────────────────────
# TABLA DE REGISTRO
# ─────────────────────────────────────────────
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

st.markdown("""
<div class="registro-header">
    📋 Registro del Día — Detalle de Cajas Empacadas
</div>
""", unsafe_allow_html=True)

if st.session_state.registro:
    reg_df = pd.DataFrame(st.session_state.registro)

    # Filtro por fecha seleccionada
    fecha_str = fecha_sel.strftime("%Y-%m-%d")
    filtrado = reg_df[reg_df["fecha"] == fecha_str].copy()

    # Filtro adicional — mostrar todo o solo el día
    col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
    with col_f1:
        mostrar_todo = st.checkbox("Mostrar todos los días en el registro", value=False)
    with col_f3:
        if filtrado.empty:
            pass
        else:
            # Botón exportar CSV
            csv_data = filtrado.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Exportar CSV",
                data=csv_data,
                file_name=f"cajas_{fecha_str}.csv",
                mime="text/csv",
                use_container_width=True,
            )

    tabla = reg_df if mostrar_todo else filtrado

    if tabla.empty:
        st.info(f"Sin registros para {fecha_sel.strftime('%d/%m/%Y')}. Usa el panel de arriba para ingresar.")
    else:
        # Preparar para mostrar — compatibilidad con registros antiguos
        for campo in ["turno", "operario", "presentacion", "factor_ref"]:
            if campo not in tabla.columns:
                tabla[campo] = ""
        tabla_show = tabla[["fecha", "codigo", "descripcion", "presentacion", "cajas", "factor_ref", "factor", "total_unidades", "turno", "operario"]].copy()
        tabla_show.columns = ["Fecha", "Código", "Descripción", "Presentación", "Cantidad", "Factor Ref.", "Factor Real", "Total Unidades", "Turno", "Operario"]
        tabla_show = tabla_show.reset_index(drop=True)
        tabla_show.index = tabla_show.index + 1  # Numerar desde 1

        # Totales
        total_cajas_dia = tabla_show["Cantidad"].sum()
        total_unid_dia = tabla_show["Total Unidades"].sum()

        # Métricas resumen
        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="value">{len(tabla_show)}</div>
                <div class="label">Líneas de registro</div>
            </div>""", unsafe_allow_html=True)
        with mc2:
            st.markdown(f"""
            <div class="metric-card">
                <div class="value">{total_cajas_dia:,.0f}</div>
                <div class="label">Total Cajas</div>
            </div>""", unsafe_allow_html=True)
        with mc3:
            st.markdown(f"""
            <div class="metric-card">
                <div class="value">{total_unid_dia:,.0f}</div>
                <div class="label">Total Unidades</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Tabla con formato
        st.dataframe(
            tabla_show.style
                .format({
                    "Cantidad": "{:,.2f}",
                    "Factor Ref.": "{:,.2f}",
                    "Factor Real": "{:,.2f}",
                    "Total Unidades": "{:,.2f}",
                })
                .background_gradient(subset=["Total Unidades"], cmap="Blues")
                .set_properties(**{"font-size": "13px"}),
            use_container_width=True,
            height=min(40 * len(tabla_show) + 80, 600),
            hide_index=False,
        )

        # Botón eliminar último registro
        st.markdown("<br>", unsafe_allow_html=True)
        col_del1, col_del2 = st.columns([2, 2])
        with col_del1:
            if st.button("↩️ Eliminar último registro", type="secondary"):
                # Encuentra y elimina el último de la fecha actual (o último global si mostrar todo)
                if mostrar_todo:
                    st.session_state.registro.pop()
                else:
                    # Busca el último index del día en el registro completo
                    indices_dia = [i for i, r in enumerate(st.session_state.registro) if r["fecha"] == fecha_str]
                    if indices_dia:
                        st.session_state.registro.pop(indices_dia[-1])
                guardar_registro(st.session_state.registro)
                st.rerun()
else:
    st.info("📭 No hay registros todavía. Ingresa el primer producto arriba.")

# ─────────────────────────────────────────────
# GESTIÓN DE OPERARIOS
# ─────────────────────────────────────────────
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

st.markdown("""
<div class="registro-header" style="background: linear-gradient(90deg, #0f3460, #1a1a6e);">
    👷 Gestión de Operarios
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

col_op1, col_op2 = st.columns([1, 1.4], gap="large")

with col_op1:
    st.markdown("#### ➕ Registrar nuevo operario")
    with st.container():
        nombre = st.text_input("Nombre", placeholder="Ej: Juan", key="op_nombre")
        apellido = st.text_input("Apellido", placeholder="Ej: García", key="op_apellido")

        if st.button("💾 Agregar Operario", type="primary", use_container_width=True):
            nombre = nombre.strip().title()
            apellido = apellido.strip().title()
            if not nombre or not apellido:
                st.error("⚠️ Ingresa nombre y apellido.")
            else:
                nombre_completo = f"{nombre} {apellido}"
                if nombre_completo in st.session_state.operarios:
                    st.warning(f"⚠️ '{nombre_completo}' ya está registrado.")
                else:
                    st.session_state.operarios.append(nombre_completo)
                    st.session_state.operarios.sort()
                    guardar_operarios(st.session_state.operarios)
                    st.success(f"✅ Operario '{nombre_completo}' registrado.")
                    st.rerun()

with col_op2:
    st.markdown("#### 📋 Operarios registrados")
    if st.session_state.operarios:
        ops_df = pd.DataFrame(
            {"#": range(1, len(st.session_state.operarios) + 1),
             "Nombre completo": st.session_state.operarios}
        )
        st.dataframe(ops_df, use_container_width=True, hide_index=True,
                     height=min(40 * len(ops_df) + 60, 350))

        op_eliminar = st.selectbox(
            "Eliminar operario:",
            options=["— Seleccionar —"] + st.session_state.operarios,
            key="op_del_select",
        )
        if st.button("🗑️ Eliminar operario seleccionado", type="secondary", use_container_width=True):
            if op_eliminar == "— Seleccionar —":
                st.warning("Selecciona un operario para eliminar.")
            else:
                st.session_state.operarios.remove(op_eliminar)
                guardar_operarios(st.session_state.operarios)
                st.success(f"Operario '{op_eliminar}' eliminado.")
                st.rerun()
    else:
        st.info("Sin operarios registrados aún. Agrega el primero en el panel de la izquierda.")


st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
<div style='text-align:center;color:#444;font-size:0.8rem;padding:1rem 0;border-top:1px solid #222;'>
    María Almenara · White Rolling Pin S.R.L. · Registro de Cajas Empacadas
</div>
""", unsafe_allow_html=True)
