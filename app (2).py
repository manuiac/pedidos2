"""
Gestión de Pedidos y Control de Presupuestos Zonales
Conectada a la base: Gestión_zonal_DR_Julio.xlsx

CÓMO CORRER:
    pip install streamlit pandas openpyxl
    streamlit run app.py

CARGA DEL EXCEL (importante):
    - Si el archivo está en la MISMA carpeta que app.py, lo toma solo.
    - Si NO lo encuentra (ej. en Streamlit Cloud sin subir el archivo),
      la app te muestra un botón para subirlo y sigue funcionando igual.

LÓGICA DE NEGOCIO:
    - CC (numérico)   = llave del centro de costo (Insumos, Plan, Atributos)
    - Insumo (texto)  = llave del producto -> precio en hoja Precios
    - Fmes (1..12)    = mes;  julio = 7
    - Proyección  = suma de Total del pedido de insumos (en pesos)
    - Presupuesto = |Monto M$| del Plan x 1.000.000  (viene en millones y negativo)
    - Rol zonal   = "Responsable Nivel 2"; cada zonal ve solo SUS CC.
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path
import io

# ======================================================================
# CONFIGURACIÓN
# ======================================================================
NOMBRE_EXCEL = "Gestión_zonal_DR_Julio.xlsx"
BASE_DIR = Path(__file__).parent
RUTA_EXCEL = BASE_DIR / NOMBRE_EXCEL
MES_DEFAULT = 7                 # julio
PLAN_A_PESOS = 1_000_000        # Monto M$ está en millones de pesos

MESES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
         7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

st.set_page_config(page_title="Pedidos Zonales", page_icon="📦", layout="wide")


# ======================================================================
# ORIGEN DEL EXCEL: disco -> o subida manual desde la interfaz
# ======================================================================
def resolver_fuente_excel():
    """Devuelve la fuente del Excel (ruta en disco o bytes subidos) o None."""
    if RUTA_EXCEL.exists():
        return str(RUTA_EXCEL)
    if st.session_state.get("excel_bytes") is not None:
        return st.session_state["excel_bytes"]
    return None


def _abrir(src):
    return io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src


# ======================================================================
# CARGA DE DATOS  (cacheada por la fuente -> se refresca si cambia el archivo)
# ======================================================================
@st.cache_data(show_spinner="Cargando base de datos...")
def cargar_insumos(src) -> pd.DataFrame:
    df = pd.read_excel(_abrir(src), sheet_name="Insumos", header=4)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for c in ["CC", "Cantidad", "Total", "Precio Unitario"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=["CC", "Insumo"])
    df["CC"] = df["CC"].astype(int)
    return df


@st.cache_data
def cargar_precios(src) -> pd.DataFrame:
    df = pd.read_excel(_abrir(src), sheet_name="Precios")
    df["Precio Unitario"] = pd.to_numeric(df["Precio Unitario"], errors="coerce")
    return df.dropna(subset=["Insumo"])


@st.cache_data
def cargar_plan(src) -> pd.DataFrame:
    df = pd.read_excel(_abrir(src), sheet_name="Plan")
    df["CC"] = pd.to_numeric(df["CC"], errors="coerce")
    df["Monto M$"] = pd.to_numeric(df["Monto M$"], errors="coerce")
    df = df.dropna(subset=["CC"])
    df["CC"] = df["CC"].astype(int)
    df["presupuesto_pesos"] = df["Monto M$"].abs() * PLAN_A_PESOS
    return df


def _src():
    return st.session_state["src"]


def precio_map() -> dict:
    pr = cargar_precios(_src())
    return dict(zip(pr["Insumo"], pr["Precio Unitario"]))


if "pedidos_enviados" not in st.session_state:
    st.session_state.pedidos_enviados = {}   # (CC, mes) -> DataFrame editado


# ======================================================================
# ROLES
# ======================================================================
@st.cache_data
def lista_responsables(src) -> list:
    ins = cargar_insumos(src)
    return sorted(ins["Responsable Nivel 2"].dropna().unique().tolist())


def ccs_de(usuario: str, es_admin: bool) -> pd.DataFrame:
    ins = cargar_insumos(_src())
    cc = ins[["CC", "Nombre CC", "Cliente", "Región", "Responsable Nivel 2"]].drop_duplicates("CC")
    return cc if es_admin else cc[cc["Responsable Nivel 2"] == usuario]


def presupuesto_cc(cc: int, mes: int):
    pl = cargar_plan(_src())
    fila = pl[(pl["CC"] == cc) & (pl["Fmes"] == mes)]["presupuesto_pesos"]
    return float(fila.iloc[0]) if not fila.empty else None


# ======================================================================
# MÓDULO PEDIDOS
# ======================================================================
def vista_pedidos(usuario: str, es_admin: bool):
    st.header("📝 Generación de Pedidos")

    ccs = ccs_de(usuario, es_admin)
    if ccs.empty:
        st.warning("No tienes Centros de Costo asignados.")
        return

    c1, c2 = st.columns(2)
    mes = c1.selectbox("Mes", list(MESES.keys()), index=MES_DEFAULT - 1,
                       format_func=lambda m: MESES[m])
    etiquetas = (ccs["CC"].astype(str) + " — " + ccs["Nombre CC"].fillna("")).tolist()
    cc_label = c2.selectbox("Centro de Costo", etiquetas)
    cc = int(cc_label.split(" — ")[0])

    clave = (cc, mes)
    if clave in st.session_state.pedidos_enviados:
        base = st.session_state.pedidos_enviados[clave].copy()
        st.info("Mostrando el pedido ya enviado para este CC y mes. Puedes ajustarlo y reenviar.")
    else:
        ins = cargar_insumos(_src())
        precios = precio_map()
        base = ins[ins["CC"] == cc][["Insumo", "Categoria", "Proveedor", "Cantidad"]].copy()
        base = base.drop_duplicates("Insumo")
        base["Precio"] = base["Insumo"].map(precios)
        faltan = base["Precio"].isna().sum()
        if faltan:
            st.caption(f"⚠️ {faltan} insumo(s) sin precio en la hoja Precios (quedan en 0; revísalos en Catálogo).")
        base["Precio"] = base["Precio"].fillna(0)
        base["Cantidad"] = pd.to_numeric(base["Cantidad"], errors="coerce").fillna(0)

    if base.empty:
        st.info("Este CC no tiene insumos históricos. Agrega filas manualmente abajo.")
        base = pd.DataFrame(columns=["Insumo", "Categoria", "Proveedor", "Cantidad", "Precio"])

    st.subheader("Edita solo la columna Cantidad")
    editado = st.data_editor(
        base[["Insumo", "Categoria", "Proveedor", "Precio", "Cantidad"]],
        column_config={
            "Insumo": st.column_config.TextColumn("Insumo", disabled=True, width="large"),
            "Categoria": st.column_config.TextColumn("Categoría", disabled=True),
            "Proveedor": st.column_config.TextColumn("Proveedor", disabled=True),
            "Precio": st.column_config.NumberColumn("Precio unit.", format="$%d", disabled=True),
            "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=0, step=1),
        },
        hide_index=True, use_container_width=True, num_rows="dynamic",
        key=f"editor_{cc}_{mes}",
    )

    editado["Cantidad"] = pd.to_numeric(editado["Cantidad"], errors="coerce").fillna(0)
    editado["Total"] = editado["Cantidad"] * editado["Precio"].fillna(0)
    proyeccion = float(editado["Total"].sum())
    presup = presupuesto_cc(cc, mes)

    m1, m2, m3 = st.columns(3)
    m1.metric("Proyección pedido", f"${proyeccion:,.0f}")
    if presup is None:
        m2.metric("Presupuesto", "Sin plan")
        m3.metric("Saldo", "—")
        st.warning("Este CC no tiene presupuesto en el Plan para este mes.")
    else:
        saldo = presup - proyeccion
        m2.metric("Presupuesto (Plan)", f"${presup:,.0f}")
        m3.metric("Saldo", f"${saldo:,.0f}", delta=f"{saldo:,.0f}")
        if proyeccion > presup:
            st.error(f"⚠️ Sobregiro de ${proyeccion - presup:,.0f}. Requiere aprobación del administrador.")

    if st.button("✅ Enviar pedido", type="primary", use_container_width=True):
        st.session_state.pedidos_enviados[clave] = editado.copy()
        st.success(f"Pedido del CC {cc} ({MESES[mes]}) enviado. Proyección: ${proyeccion:,.0f}")


# ======================================================================
# MÓDULO CATÁLOGO  (admin)
# ======================================================================
def vista_catalogo():
    st.header("⚙️ Administración de Catálogo (Precios)")
    st.caption("Actualiza precios, agrega o elimina insumos. (En el MVP los cambios no persisten al recargar.)")
    pr = cargar_precios(_src())
    editado = st.data_editor(
        pr,
        column_config={"Precio Unitario": st.column_config.NumberColumn("Precio Unitario", format="$%d", min_value=0)},
        num_rows="dynamic", use_container_width=True, hide_index=True,
    )
    if st.button("💾 Guardar catálogo"):
        st.success(f"Catálogo guardado: {len(editado)} insumos (simulado).")


# ======================================================================
# MÓDULO DASHBOARD  (admin)
# ======================================================================
def vista_dashboard():
    st.header("📊 Dashboard y Control Presupuestario")
    mes = st.selectbox("Mes", list(MESES.keys()), index=MES_DEFAULT - 1,
                       format_func=lambda m: MESES[m])

    ins = cargar_insumos(_src())
    pl = cargar_plan(_src())

    base = ins.groupby(["CC", "Nombre CC", "Responsable Nivel 2"], as_index=False)["Total"].sum()
    base = base.rename(columns={"Total": "proyeccion"})
    for (cc, m), df in st.session_state.pedidos_enviados.items():
        if m == mes:
            base.loc[base["CC"] == cc, "proyeccion"] = float(df["Total"].sum())

    plan_mes = pl[pl["Fmes"] == mes][["CC", "presupuesto_pesos"]]
    tabla = base.merge(plan_mes, on="CC", how="left")
    tabla["saldo"] = tabla["presupuesto_pesos"] - tabla["proyeccion"]
    tabla["sobregiro"] = tabla["saldo"] < 0

    f1, f2 = st.columns(2)
    resp_op = sorted(tabla["Responsable Nivel 2"].dropna().unique().tolist())
    resp = f1.multiselect("Responsable", resp_op, default=resp_op)
    tabla = tabla[tabla["Responsable Nivel 2"].isin(resp)]
    cc_op = (tabla["CC"].astype(str) + " — " + tabla["Nombre CC"].fillna("")).tolist()
    cc_sel = f2.multiselect("Centro de Costo", cc_op, default=cc_op)
    ccs_filtro = [int(x.split(" — ")[0]) for x in cc_sel]
    tabla = tabla[tabla["CC"].isin(ccs_filtro)]

    k1, k2, k3 = st.columns(3)
    k1.metric("Proyección total", f"${tabla['proyeccion'].sum():,.0f}")
    k2.metric("Presupuesto total", f"${tabla['presupuesto_pesos'].sum():,.0f}")
    k3.metric("CC en sobregiro", int(tabla["sobregiro"].sum()))

    sobre = tabla[tabla["sobregiro"]]
    for _, r in sobre.iterrows():
        st.error(f"⚠️ {r['Nombre CC']} (CC {r['CC']}): sobregiro de ${-r['saldo']:,.0f}")

    st.dataframe(
        tabla[["CC", "Nombre CC", "Responsable Nivel 2", "proyeccion", "presupuesto_pesos", "saldo"]]
        .rename(columns={"Responsable Nivel 2": "Responsable", "proyeccion": "Proyección",
                         "presupuesto_pesos": "Presupuesto", "saldo": "Saldo"}),
        use_container_width=True, hide_index=True,
        column_config={
            "Proyección": st.column_config.NumberColumn(format="$%d"),
            "Presupuesto": st.column_config.NumberColumn(format="$%d"),
            "Saldo": st.column_config.NumberColumn(format="$%d"),
        },
    )


# ======================================================================
# MÓDULO EXPORTAR
# ======================================================================
def vista_exportar():
    st.header("📥 Exportar Consolidado de Pedidos Enviados")
    enviados = st.session_state.pedidos_enviados
    if not enviados:
        st.info("Aún no hay pedidos enviados en esta sesión.")
        return

    filas = []
    for (cc, mes), df in enviados.items():
        tmp = df.copy()
        tmp["CC"] = cc
        tmp["Mes"] = MESES[mes]
        filas.append(tmp)
    consolidado = pd.concat(filas, ignore_index=True)[
        ["CC", "Mes", "Insumo", "Categoria", "Proveedor", "Precio", "Cantidad", "Total"]]

    st.dataframe(consolidado, use_container_width=True, hide_index=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        consolidado.to_excel(writer, sheet_name="Consolidado", index=False)
    st.download_button(
        "⬇️ Descargar consolidado (Excel)", data=buffer.getvalue(),
        file_name=f"consolidado_pedidos_{datetime.now():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
    )


# ======================================================================
# PANTALLA DE CARGA DEL EXCEL (cuando no se encuentra en disco)
# ======================================================================
def pantalla_subir_excel():
    st.title("📦 Pedidos Zonales")
    st.warning(f"No encontré el archivo **{NOMBRE_EXCEL}** junto a la app.")
    st.write("Sube la base de datos para comenzar (hojas: Insumos, Precios, Plan).")
    archivo = st.file_uploader("Archivo Excel", type=["xlsx"])
    if archivo is not None:
        st.session_state["excel_bytes"] = archivo.getvalue()
        st.cache_data.clear()
        st.rerun()
    st.caption("Tip: para no subirlo cada vez, déjalo en el repositorio de GitHub "
               "junto al app.py y la app lo tomará automáticamente.")


# ======================================================================
# ROUTER + SIDEBAR
# ======================================================================
def main():
    fuente = resolver_fuente_excel()
    if fuente is None:
        pantalla_subir_excel()
        st.stop()
    st.session_state["src"] = fuente

    st.sidebar.title("📦 Pedidos Zonales")
    perfiles = ["Administrador"] + lista_responsables(_src())
    perfil = st.sidebar.selectbox("Perfil", perfiles)
    es_admin = (perfil == "Administrador")
    st.sidebar.caption(f"Sesión: **{perfil}**" + ("  ·  acceso total" if es_admin else "  ·  zonal"))
    st.sidebar.divider()

    opciones = (["Generación de Pedidos", "Catálogo", "Dashboard", "Exportar"]
                if es_admin else ["Generación de Pedidos"])
    seccion = st.sidebar.radio("Módulos", opciones)

    if seccion == "Generación de Pedidos":
        vista_pedidos(perfil, es_admin)
    elif seccion == "Catálogo":
        vista_catalogo()
    elif seccion == "Dashboard":
        vista_dashboard()
    elif seccion == "Exportar":
        vista_exportar()


if __name__ == "__main__":
    main()
