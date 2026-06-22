import streamlit as st
import pandas as pd
from datetime import date
import json
import gspread
from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Registro de Cajas Empacadas",
    page_icon="📦",
    layout="wide",
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    }
    .main-header h1 { color: #e94560; font-size: 2rem; font-weight: 800; margin: 0; letter-spacing: 1px; }
    .main-header p  { color: #a8b2d8; margin: 0.3rem 0 0 0; font-size: 0.95rem; }

    .metric-card {
        background: linear-gradient(135deg, #0f3460, #16213e);
        border: 1px solid #e94560; border-radius: 10px;
        padding: 1rem 1.2rem; text-align: center;
        box-shadow: 0 2px 12px rgba(233,69,96,0.2);
    }
    .metric-card .value { font-size: 2rem; font-weight: 800; color: #e94560; }
    .metric-card .label { font-size: 0.8rem; color: #a8b2d8; text-transform: uppercase; letter-spacing: 1px; }

    .product-card {
        background: linear-gradient(135deg, #0d1b2a, #1b263b);
        border: 2px solid #4cc9f0; border-radius: 12px;
        padding: 1.2rem 1.5rem; margin: 0.8rem 0;
    }
    .product-card .prod-code  { font-size: 0.8rem; color: #4cc9f0; font-weight: 700; letter-spacing: 2px; text-transform: uppercase; }
    .product-card .prod-name  { font-size: 1.2rem; font-weight: 700; color: #f0f4f8; margin: 0.2rem 0; }
    .product-card .prod-factor{ font-size: 0.9rem; color: #a8b2d8; }
    .prod-factor span          { color: #f6c90e; font-weight: 700; font-size: 1rem; }

    .registro-header {
        background: linear-gradient(90deg, #e94560, #c1121f);
        color: white; padding: 0.7rem 1rem; border-radius: 8px 8px 0 0;
        font-weight: 700; font-size: 1.1rem;
    }
    .divider { height: 2px; background: linear-gradient(90deg, #e94560, transparent); margin: 1.5rem 0; border-radius: 2px; }

    div.stButton > button { border-radius: 8px; font-weight: 700; transition: all 0.2s; }

    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #1a1a2e, #0f3460); }
    section[data-testid="stSidebar"] .stMarkdown h2 { color: #4cc9f0; }
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] label { color: #a8b2d8 !important; }
    section[data-testid="stSidebar"] .stMarkdown strong { color: #e94560; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CONEXIÓN GOOGLE SHEETS
# ─────────────────────────────────────────────
@st.cache_resource
def get_sheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    spreadsheet_id = st.secrets["SPREADSHEET_ID"]
    sh = gc.open_by_key(spreadsheet_id)
    return sh

def get_ws(nombre):
    sh = get_sheet()
    return sh.worksheet(nombre)


# ─────────────────────────────────────────────
# CATÁLOGO
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar_catalogo():
    ws = get_ws("catalogo")
    data = ws.get_all_records()
    df = pd.DataFrame(data)
    df.columns = df.columns.str.strip()
    df["CÓDIGO"] = df["CÓDIGO"].astype(str).str.strip()
    df["DESCRIPCIÓN PRODUCTO"] = df["DESCRIPCIÓN PRODUCTO"].astype(str).str.strip()
    df["FACTOR"] = pd.to_numeric(df["FACTOR"], errors="coerce").fillna(0)
    return df


# ─────────────────────────────────────────────
# REGISTRO
# ─────────────────────────────────────────────
COLS_REGISTRO = ["fecha", "codigo", "descripcion", "presentacion",
                 "cajas", "factor_ref", "factor", "total_unidades",
                 "turno", "operario"]

def cargar_registro():
    try:
        ws = get_ws("registro")
        data = ws.get_all_records()
        if not data:
            return []
        df = pd.DataFrame(data)
        # Asegurar columnas
        for c in COLS_REGISTRO:
            if c not in df.columns:
                df[c] = ""
        return df[COLS_REGISTRO].to_dict("records")
    except Exception as e:
        st.error(f"Error cargando registro: {e}")
        return []

def guardar_registro_fila(fila: dict):
    """Agrega UNA fila nueva al sheet de registro (no reescribe todo)."""
    try:
        ws = get_ws("registro")
        # Si la hoja está vacía, poner encabezados
        if ws.row_count == 0 or ws.acell("A1").value is None:
            ws.append_row(COLS_REGISTRO)
        row = [str(fila.get(c, "")) for c in COLS_REGISTRO]
        ws.append_row(row, value_input_option="USER_ENTERED")
    except Exception as e:
        st.error(f"Error guardando en Sheets: {e}")

def eliminar_ultima_fila_registro():
    """Elimina la última fila de datos del sheet registro."""
    try:
        ws = get_ws("registro")
        last = len(ws.get_all_values())  # incluye encabezado
        if last > 1:
            ws.delete_rows(last)
    except Exception as e:
        st.error(f"Error eliminando fila: {e}")

def eliminar_fila_por_indice(indice_global: int):
    """Elimina la fila del sheet por índice (0-based sobre los datos, fila 2 = índice 0)."""
    try:
        ws = get_ws("registro")
        ws.delete_rows(indice_global + 2)  # +2: 1 encabezado + base 1
    except Exception as e:
        st.error(f"Error eliminando fila: {e}")

def actualizar_fila_en_sheet(indice_global: int, fila: dict):
    """Reemplaza los valores de una fila existente en el sheet (índice 0-based sobre datos)."""
    try:
        ws = get_ws("registro")
        sheet_row = indice_global + 2  # +1 encabezado, +1 base 1
        row = [str(fila.get(c, "")) for c in COLS_REGISTRO]
        ws.update(f"A{sheet_row}", [row], value_input_option="USER_ENTERED")
    except Exception as e:
        st.error(f"Error actualizando fila: {e}")

def limpiar_registro_completo():
    try:
        ws = get_ws("registro")
        ws.clear()
        ws.append_row(COLS_REGISTRO)
    except Exception as e:
        st.error(f"Error limpiando registro: {e}")


# ─────────────────────────────────────────────
# OPERARIOS
# ─────────────────────────────────────────────
def cargar_operarios():
    try:
        ws = get_ws("operarios")
        data = ws.get_all_values()
        # Columna A: nombres, saltando encabezado si existe
        nombres = []
        for row in data:
            if row and row[0] and row[0].strip().lower() != "nombre":
                nombres.append(row[0].strip())
        return sorted(nombres)
    except Exception as e:
        st.error(f"Error cargando operarios: {e}")
        return []

def agregar_operario(nombre_completo: str):
    try:
        ws = get_ws("operarios")
        if ws.row_count == 0 or ws.acell("A1").value is None:
            ws.append_row(["nombre"])
        ws.append_row([nombre_completo])
    except Exception as e:
        st.error(f"Error agregando operario: {e}")

def eliminar_operario(nombre_completo: str):
    try:
        ws = get_ws("operarios")
        cell = ws.find(nombre_completo)
        if cell:
            ws.delete_rows(cell.row)
    except Exception as e:
        st.error(f"Error eliminando operario: {e}")


# ─────────────────────────────────────────────
# ESTADO DE SESIÓN
# ─────────────────────────────────────────────
if "registro" not in st.session_state:
    with st.spinner("Cargando registro desde Google Sheets..."):
        st.session_state.registro = cargar_registro()

if "operarios" not in st.session_state:
    with st.spinner("Cargando operarios..."):
        st.session_state.operarios = cargar_operarios()

if "producto_sel" not in st.session_state:
    st.session_state.producto_sel = None

if "editando_idx" not in st.session_state:
    st.session_state.editando_idx = None  # índice global en registro

# ─────────────────────────────────────────────
# CATÁLOGO
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
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📅 Fecha de Registro")
    fecha_sel = st.date_input("Selecciona la fecha", value=date.today(), format="DD/MM/YYYY")

    st.markdown("---")
    st.markdown("## 🔍 Buscar Producto")
    modo_busqueda = st.radio("Filtrar por:", ["CÓDIGO", "DESCRIPCIÓN"], horizontal=True)
    busqueda = st.text_input("Ingresa texto", placeholder="Ej: M6020039 o ALFAJOR", key="input_busqueda")

    st.markdown("---")
    st.markdown("## 📊 Resumen de Sesión")
    reg_df_side = pd.DataFrame(st.session_state.registro)
    if not reg_df_side.empty and "fecha" in reg_df_side.columns:
        fecha_str_side = fecha_sel.strftime("%Y-%m-%d")
        hoy = reg_df_side[reg_df_side["fecha"] == fecha_str_side]
        st.markdown(f"**Fecha:** {fecha_sel.strftime('%d/%m/%Y')}")
        st.markdown(f"**Líneas hoy:** {len(hoy)}")
        total_cajas_s = pd.to_numeric(hoy["cajas"], errors="coerce").sum() if not hoy.empty else 0
        total_unid_s  = pd.to_numeric(hoy["total_unidades"], errors="coerce").sum() if not hoy.empty else 0
        st.markdown(f"**Total cant.:** {total_cajas_s:,.0f}")
        st.markdown(f"**Total unid.:** {total_unid_s:,.2f}")
    else:
        st.markdown("Sin registros aún.")

    st.markdown("---")
    if st.button("🔄 Recargar desde Sheets", use_container_width=True):
        st.session_state.registro  = cargar_registro()
        st.session_state.operarios = cargar_operarios()
        st.cache_data.clear()
        st.rerun()

    if st.button("🗑️ Limpiar TODO el registro", type="secondary"):
        limpiar_registro_completo()
        st.session_state.registro = []
        st.success("Registro limpiado.")
        st.rerun()

# ─────────────────────────────────────────────
# BÚSQUEDA Y SELECCIÓN
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
        resultado_display = resultado[["CÓDIGO", "DESCRIPCIÓN PRODUCTO", "FACTOR"]].reset_index(drop=True)
        opciones_list = (resultado_display["CÓDIGO"] + " — " + resultado_display["DESCRIPCIÓN PRODUCTO"]).tolist()

        seleccion = st.selectbox("Selecciona un producto:", options=opciones_list, index=None, placeholder="Haz clic para elegir...")

        if seleccion:
            codigo_sel = seleccion.split(" — ")[0]
            fila = catalogo[catalogo["CÓDIGO"] == codigo_sel].iloc[0]
            st.session_state.producto_sel = {
                "codigo": fila["CÓDIGO"],
                "descripcion": fila["DESCRIPCIÓN PRODUCTO"],
                "factor": float(fila["FACTOR"]),
            }

        st.dataframe(
            resultado_display.style.format({"FACTOR": "{:,.2f}"}),
            use_container_width=True, height=300, hide_index=True,
        )

with col_der:
    st.markdown("### 📝 Ingreso de Cajas")

    if st.session_state.producto_sel:
        p = st.session_state.producto_sel

        st.markdown(f"""
        <div class="product-card">
            <div class="prod-code">🏷️ {p['codigo']}</div>
            <div class="prod-name">{p['descripcion']}</div>
            <div class="prod-factor">Factor referencial: <span>{p['factor']:,.2f}</span> unidades/caja</div>
        </div>
        """, unsafe_allow_html=True)

        # ── Presentación + Factor real ──
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
                help=f"Factor referencial: {p['factor']:,.2f} — modifícalo si cambió",
            )

        if abs(factor_real - p["factor"]) > 0.001:
            st.markdown(
                f"<small style='color:#f6c90e'>⚠️ Factor modificado "
                f"(ref: <b>{p['factor']:,.2f}</b> → actual: <b>{factor_real:,.2f}</b>)</small>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='margin-top:0.4rem'></div>", unsafe_allow_html=True)

        # ── Cantidad ──
        cajas = st.number_input(
            f"📦 Cantidad de {presentacion.lower()}s",
            min_value=0.0, value=0.0, step=1.0, format="%.2f",
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

        # ── Turno + Operario ──
        col_turno, col_op = st.columns(2)
        with col_turno:
            turno = st.selectbox("🕐 Turno", options=["Turno Mañana", "Turno Tarde", "Turno Noche"], index=0)
        with col_op:
            lista_ops = st.session_state.operarios
            if lista_ops:
                operario = st.selectbox("👷 Operario", options=["— Seleccionar —"] + lista_ops, index=0)
                operario = "" if operario == "— Seleccionar —" else operario
            else:
                st.selectbox("👷 Operario", options=["Sin operarios registrados"], disabled=True)
                operario = ""
                st.caption("⚠️ Registra operarios en la sección de abajo.")

        # ── Botones ──
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("✅ GRABAR", type="primary", use_container_width=True):
                if cajas <= 0:
                    st.error("⚠️ La cantidad debe ser mayor a 0.")
                else:
                    nuevo = {
                        "fecha"         : fecha_sel.strftime("%Y-%m-%d"),
                        "codigo"        : p["codigo"],
                        "descripcion"   : p["descripcion"],
                        "presentacion"  : presentacion,
                        "cajas"         : cajas,
                        "factor_ref"    : p["factor"],
                        "factor"        : factor_real,
                        "total_unidades": round(cajas * factor_real, 4),
                        "turno"         : turno,
                        "operario"      : operario,
                    }
                    with st.spinner("Guardando en Google Sheets..."):
                        guardar_registro_fila(nuevo)
                    st.session_state.registro.append(nuevo)
                    st.success(f"✅ {p['descripcion']} — {cajas:,.2f} {presentacion.lower()}s → {nuevo['total_unidades']:,.2f} unidades")
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
st.markdown('<div class="registro-header">📋 Registro del Día — Detalle de Cajas Empacadas</div>', unsafe_allow_html=True)

if st.session_state.registro:
    reg_df = pd.DataFrame(st.session_state.registro)
    fecha_str = fecha_sel.strftime("%Y-%m-%d")
    filtrado = reg_df[reg_df["fecha"] == fecha_str].copy()

    col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
    with col_f1:
        mostrar_todo = st.checkbox("Mostrar todos los días en el registro", value=False)
    with col_f3:
        if not filtrado.empty:
            csv_data = filtrado.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Exportar CSV", data=csv_data,
                               file_name=f"cajas_{fecha_str}.csv", mime="text/csv",
                               use_container_width=True)

    tabla = reg_df if mostrar_todo else filtrado

    if tabla.empty:
        st.info(f"Sin registros para {fecha_sel.strftime('%d/%m/%Y')}.")
    else:
        for campo in ["turno", "operario", "presentacion", "factor_ref"]:
            if campo not in tabla.columns:
                tabla[campo] = ""

        # Índices globales en session_state.registro para cada fila visible
        if mostrar_todo:
            indices_globales = list(range(len(st.session_state.registro)))
        else:
            indices_globales = [i for i, r in enumerate(st.session_state.registro)
                                if r.get("fecha") == fecha_str]

        tabla = tabla.reset_index(drop=True)
        tabla_show = tabla[["fecha","codigo","descripcion","presentacion","cajas",
                             "factor_ref","factor","total_unidades","turno","operario"]].copy()
        tabla_show.columns = ["Fecha","Código","Descripción","Presentación","Cantidad",
                               "Factor Ref.","Factor Real","Total Unidades","Turno","Operario"]
        for col_num in ["Cantidad","Factor Ref.","Factor Real","Total Unidades"]:
            tabla_show[col_num] = pd.to_numeric(tabla_show[col_num], errors="coerce")

        total_cajas_dia = tabla_show["Cantidad"].sum()
        total_unid_dia  = tabla_show["Total Unidades"].sum()

        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            st.markdown(f'<div class="metric-card"><div class="value">{len(tabla_show)}</div><div class="label">Líneas</div></div>', unsafe_allow_html=True)
        with mc2:
            st.markdown(f'<div class="metric-card"><div class="value">{total_cajas_dia:,.0f}</div><div class="label">Total Cantidad</div></div>', unsafe_allow_html=True)
        with mc3:
            st.markdown(f'<div class="metric-card"><div class="value">{total_unid_dia:,.0f}</div><div class="label">Total Unidades</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Selector de fila ──────────────────────────────────
        opciones_filas = {
            f"#{i+1} — {row['Descripción']} | {row['Presentación']} | {row['Cantidad']:,.0f} uds | {row['Turno']}": idx_global
            for i, (row, idx_global) in enumerate(zip(
                [tabla_show.iloc[j] for j in range(len(tabla_show))],
                indices_globales
            ))
        }
        fila_sel_label = st.selectbox(
            "✏️ Selecciona una línea para editar o eliminar:",
            options=["— Seleccionar línea —"] + list(opciones_filas.keys()),
            key="sel_fila_accion",
        )

        col_acc1, col_acc2, _ = st.columns([1, 1, 2])
        with col_acc1:
            btn_editar  = st.button("✏️ Editar línea",    type="primary",    use_container_width=True)
        with col_acc2:
            btn_eliminar = st.button("🗑️ Eliminar línea", type="secondary",  use_container_width=True)

        if fila_sel_label != "— Seleccionar línea —":
            idx_global_sel = opciones_filas[fila_sel_label]

            if btn_eliminar:
                with st.spinner("Eliminando en Google Sheets..."):
                    eliminar_fila_por_indice(idx_global_sel)
                st.session_state.registro.pop(idx_global_sel)
                st.session_state.editando_idx = None
                st.success("✅ Línea eliminada.")
                st.rerun()

            if btn_editar:
                st.session_state.editando_idx = idx_global_sel

        # ── Formulario de edición ─────────────────────────────
        if st.session_state.editando_idx is not None:
            idx_e = st.session_state.editando_idx
            # Verificar que el índice siga siendo válido
            if idx_e < len(st.session_state.registro):
                r = st.session_state.registro[idx_e]
                st.markdown("---")
                st.markdown(f"#### ✏️ Editando línea: **{r.get('descripcion','')}**")

                with st.container():
                    ec1, ec2 = st.columns(2)
                    with ec1:
                        e_presentacion = st.selectbox("📦 Presentación",
                            options=["Caja","Jaba","Bolsa"],
                            index=["Caja","Jaba","Bolsa"].index(r.get("presentacion","Caja"))
                                  if r.get("presentacion","Caja") in ["Caja","Jaba","Bolsa"] else 0,
                            key="e_pres")
                    with ec2:
                        e_factor = st.number_input("🔢 Factor real",
                            min_value=0.01,
                            value=float(r.get("factor", r.get("factor_ref", 1))),
                            step=1.0, format="%.2f", key="e_factor")

                    ec3, ec4 = st.columns(2)
                    with ec3:
                        e_cajas = st.number_input(f"📦 Cantidad de {e_presentacion.lower()}s",
                            min_value=0.01,
                            value=float(r.get("cajas", 1)),
                            step=1.0, format="%.2f", key="e_cajas")
                    with ec4:
                        e_turno = st.selectbox("🕐 Turno",
                            options=["Turno Mañana","Turno Tarde","Turno Noche"],
                            index=["Turno Mañana","Turno Tarde","Turno Noche"].index(r.get("turno","Turno Mañana"))
                                  if r.get("turno") in ["Turno Mañana","Turno Tarde","Turno Noche"] else 0,
                            key="e_turno")

                    lista_ops = st.session_state.operarios
                    op_actual = r.get("operario","")
                    op_opts = ["— Seleccionar —"] + lista_ops
                    op_idx = op_opts.index(op_actual) if op_actual in op_opts else 0
                    e_operario = st.selectbox("👷 Operario", options=op_opts, index=op_idx, key="e_operario")
                    e_operario = "" if e_operario == "— Seleccionar —" else e_operario

                    e_total = e_cajas * e_factor
                    st.markdown(f"""
                    <div style='background:#0d2137;border:1px solid #4cc9f0;border-radius:8px;
                                padding:0.6rem 1rem;margin:0.4rem 0;'>
                        <span style='color:#a8b2d8;font-size:0.8rem'>TOTAL UNIDADES (nuevo)</span><br>
                        <span style='color:#4cc9f0;font-size:1.5rem;font-weight:800'>{e_total:,.2f}</span>
                        <span style='color:#a8b2d8;font-size:0.8rem'> = {e_cajas:,.2f} × {e_factor:,.2f}</span>
                    </div>
                    """, unsafe_allow_html=True)

                    esb1, esb2 = st.columns(2)
                    with esb1:
                        if st.button("💾 Guardar cambios", type="primary", use_container_width=True, key="btn_guardar_edit"):
                            fila_actualizada = {
                                "fecha"         : r.get("fecha",""),
                                "codigo"        : r.get("codigo",""),
                                "descripcion"   : r.get("descripcion",""),
                                "presentacion"  : e_presentacion,
                                "cajas"         : e_cajas,
                                "factor_ref"    : r.get("factor_ref", r.get("factor",1)),
                                "factor"        : e_factor,
                                "total_unidades": round(e_total, 4),
                                "turno"         : e_turno,
                                "operario"      : e_operario,
                            }
                            with st.spinner("Actualizando en Google Sheets..."):
                                actualizar_fila_en_sheet(idx_e, fila_actualizada)
                            st.session_state.registro[idx_e] = fila_actualizada
                            st.session_state.editando_idx = None
                            st.success("✅ Registro actualizado correctamente.")
                            st.rerun()
                    with esb2:
                        if st.button("❌ Cancelar", use_container_width=True, key="btn_cancelar_edit"):
                            st.session_state.editando_idx = None
                            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Tabla visual ─────────────────────────────────────
        st.dataframe(
            tabla_show.style
                .format({"Cantidad":"{:,.2f}","Factor Ref.":"{:,.2f}","Factor Real":"{:,.2f}","Total Unidades":"{:,.2f}"})
                .background_gradient(subset=["Total Unidades"], cmap="Blues")
                .set_properties(**{"font-size":"13px"}),
            use_container_width=True,
            height=min(40 * len(tabla_show) + 80, 600),
            hide_index=True,
        )

else:
    st.info("📭 No hay registros todavía.")

# ─────────────────────────────────────────────
# GESTIÓN DE OPERARIOS
# ─────────────────────────────────────────────
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
st.markdown('<div class="registro-header" style="background:linear-gradient(90deg,#0f3460,#1a1a6e);">👷 Gestión de Operarios</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

col_op1, col_op2 = st.columns([1, 1.4], gap="large")

with col_op1:
    st.markdown("#### ➕ Registrar nuevo operario")
    nombre  = st.text_input("Nombre",   placeholder="Ej: Juan",   key="op_nombre")
    apellido = st.text_input("Apellido", placeholder="Ej: García", key="op_apellido")

    if st.button("💾 Agregar Operario", type="primary", use_container_width=True):
        nombre_l   = nombre.strip().title()
        apellido_l = apellido.strip().title()
        if not nombre_l or not apellido_l:
            st.error("⚠️ Ingresa nombre y apellido.")
        else:
            nombre_completo = f"{nombre_l} {apellido_l}"
            if nombre_completo in st.session_state.operarios:
                st.warning(f"⚠️ '{nombre_completo}' ya está registrado.")
            else:
                with st.spinner("Guardando..."):
                    agregar_operario(nombre_completo)
                st.session_state.operarios.append(nombre_completo)
                st.session_state.operarios.sort()
                st.success(f"✅ '{nombre_completo}' registrado.")
                st.rerun()

with col_op2:
    st.markdown("#### 📋 Operarios registrados")
    if st.session_state.operarios:
        ops_df = pd.DataFrame({"#": range(1, len(st.session_state.operarios)+1),
                               "Nombre completo": st.session_state.operarios})
        st.dataframe(ops_df, use_container_width=True, hide_index=True,
                     height=min(40*len(ops_df)+60, 350))

        op_eliminar = st.selectbox("Eliminar operario:", options=["— Seleccionar —"]+st.session_state.operarios, key="op_del_select")
        if st.button("🗑️ Eliminar operario seleccionado", type="secondary", use_container_width=True):
            if op_eliminar == "— Seleccionar —":
                st.warning("Selecciona un operario para eliminar.")
            else:
                with st.spinner("Eliminando..."):
                    eliminar_operario(op_eliminar)
                st.session_state.operarios.remove(op_eliminar)
                st.success(f"'{op_eliminar}' eliminado.")
                st.rerun()
    else:
        st.info("Sin operarios registrados aún.")

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
<div style='text-align:center;color:#444;font-size:0.8rem;padding:1rem 0;border-top:1px solid #222;'>
    María Almenara · White Rolling Pin S.R.L. · Registro de Cajas Empacadas
</div>
""", unsafe_allow_html=True)
