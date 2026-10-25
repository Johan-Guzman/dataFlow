"""Aplicación web: Normalizador de Excel (Streamlit)."""

import pandas as pd
import streamlit as st

import automatizacion_excel as ae
from utilidades_config import (
    COLUMNAS_CATALOGO, ENCABEZADOS, cfg_a_json, cfg_desde_json_bytes,
    construir_cfg, estado_desde_cfg, procesar_bytes,
)

st.set_page_config(page_title="Normalizador de Excel", page_icon="📊", layout="wide")

st.markdown(
    """
<style>
.block-container {padding-top: 3.6rem; max-width: 1180px;}
.hero {background: linear-gradient(120deg,#0F766E 0%,#0E7490 60%,#1D4ED8 100%);
  color:#fff; padding: 1.6rem 1.8rem; border-radius: 18px; margin-bottom: 1.2rem;}
.hero h1 {color:#fff; margin:0; font-size:1.9rem;}
.hero p {margin:.35rem 0 0; opacity:.92; font-size:1.02rem;}
[data-testid="stMetric"] {background:#fff; border:1px solid #E2E8F0; border-radius:14px;
  padding:.9rem 1rem; box-shadow:0 1px 3px rgba(15,23,42,.06);}
[data-testid="stMetricLabel"] {color:#64748B;}
.stTabs [data-baseweb="tab"] {font-weight:600; padding: .6rem 1rem;}
div[data-testid="stVerticalBlockBorderWrapper"] {border-radius:14px;}
.paso {color:#0F766E; font-weight:700; letter-spacing:.04em; font-size:.78rem; text-transform:uppercase;}
footer {visibility:hidden;}
</style>
<div class="hero">
  <h1>📊 Normalizador de Excel</h1>
  <p>Sube tu archivo, ajusta las reglas a tu manera y descarga el resultado con un reporte de auditoría.
     No se eliminan filas y el original nunca se modifica.</p>
</div>
""",
    unsafe_allow_html=True,
)

ss = st.session_state


def volcar(cfg):
    for k, v in estado_desde_cfg(cfg).items():
        ss[k] = v


if "iniciado" not in ss:
    volcar(ae.Configuracion())
    ss.iniciado = True


def _al_cargar_config():
    archivo = ss.get("cfg_upload")
    if archivo is None:
        return
    try:
        volcar(cfg_desde_json_bytes(archivo.getvalue()))
        ss.cfg_msg = ("ok", f"Configuración «{archivo.name}» cargada.")
    except Exception as e:
        ss.cfg_msg = ("error", f"No se pudo cargar la configuración: {e}")


def _restablecer():
    volcar(ae.Configuracion())
    ss.cfg_msg = ("ok", "Se restablecieron los valores por defecto.")


# ------------------------------ barra lateral ------------------------------
with st.sidebar:
    st.markdown('<span class="paso">Paso 1</span>', unsafe_allow_html=True)
    archivo = st.file_uploader("Archivo Excel", type=["xlsx", "xlsm"],
                               help="Se procesa en memoria; no se guarda en el servidor.")
    st.divider()
    st.markdown('<span class="paso">Configuración</span>', unsafe_allow_html=True)
    st.file_uploader("Cargar configuración (.json)", type=["json"], key="cfg_upload",
                     on_change=_al_cargar_config)
    if "cfg_msg" in ss:
        tipo, texto = ss.cfg_msg
        (st.success if tipo == "ok" else st.error)(texto)
    try:
        st.download_button("⬇️ Guardar configuración", cfg_a_json(construir_cfg(ss)),
                           "configuracion.json", "application/json", use_container_width=True)
    except ValueError:
        st.caption("Corrige la configuración para poder guardarla.")
    st.button("↺ Restablecer valores", on_click=_restablecer, use_container_width=True)
    st.divider()
    st.checkbox(f"Leer catálogos de la hoja «{ae.HOJA_REGLAS}»", key="leer_hoja",
                help="Si existe, reemplaza los catálogos de esta pantalla cuando los encabezados coinciden.")
    st.checkbox("Incluir celdas válidas en el reporte", key="reg_validos",
                help="Por defecto solo se cuentan en la hoja de resumen.")

# ------------------------------ pestañas ------------------------------
t_col, t_cat, t_per, t_run = st.tabs(
    ["⚙️ Columnas", "📚 Catálogos y equivalencias", "🗓️ Periodos", "▶️ Ejecutar y resultados"]
)

with t_col:
    st.caption("Escribe el encabezado (fila 1) de cada columna. Se ignoran mayúsculas, tildes, espacios y «_».")
    c1, c2 = st.columns(2)
    for i, (clave, etiqueta) in enumerate(ENCABEZADOS):
        (c1 if i % 2 == 0 else c2).text_input(etiqueta, key=clave)

with t_cat:
    st.caption("Lista oficial (un valor por línea) y equivalencias confirmadas con el formato "
               "`CC = Cédula de ciudadanía`. Con la lista vacía no se corrige nada.")
    for clave, etiqueta in COLUMNAS_CATALOGO:
        with st.container(border=True):
            st.markdown(f"**{etiqueta}**")
            a, b = st.columns(2)
            a.text_area("Valores oficiales", key=f"cat_{clave}", height=140,
                        placeholder="Un valor por línea")
            b.text_area("Equivalencias (clave = valor oficial)", key=f"eq_{clave}", height=140,
                        placeholder="CC = Cédula de ciudadanía")

with t_per:
    st.caption("Código de 2 dígitos = descripción. Los códigos que no estén aquí no reciben descripción.")
    a, b = st.columns([2, 1])
    a.text_area("Periodos académicos", key="periodos", height=150, placeholder="10 = Primer Semestre")
    b.number_input("Año mínimo válido", key="anio_min", step=1)
    b.number_input("Año máximo válido", key="anio_max", step=1)

with t_run:
    try:
        cfg = construir_cfg(ss)
        cfg_ok = True
    except (ValueError, TypeError) as e:
        cfg, cfg_ok = None, False
        st.error(f"Configuración inválida: {e}")

    if cfg_ok:
        vacios = [n for n in cfg.columnas_catalogo if not cfg.valores_validos.get(n)]
        if vacios:
            st.info("Catálogos sin definir: " + ", ".join(vacios) +
                    ". Esos valores se conservarán y se marcarán como SIN_CATALOGO.")
    if archivo is None:
        st.warning("Sube un archivo Excel en la barra lateral para continuar.")

    if st.button("▶ Ejecutar automatización", type="primary", disabled=not (cfg_ok and archivo)):
        try:
            with st.spinner("Procesando…"):
                ss.resultado = procesar_bytes(archivo.getvalue(), archivo.name, cfg)
        except Exception as e:
            ss.pop("resultado", None)
            st.error(f"Error al procesar: {e}")

    r = ss.get("resultado")
    if r:
        st.success(f"Listo: {r['nombre_salida']}")
        res = r["resumen"]
        m = st.columns(5)
        m[0].metric("Corregidos", res.get("CORREGIDO", 0))
        m[1].metric("Válidos", res.get("VALIDO", 0))
        m[2].metric("Inválidos", res.get("INVALIDO", 0))
        m[3].metric("Requieren revisión", res.get("REQUIERE_REVISION", 0))
        m[4].metric("Sin catálogo", res.get("SIN_CATALOGO", 0))
        for aviso in r["avisos"]:
            st.warning(aviso)
        st.download_button(
            "⬇️ Descargar Excel corregido", r["bytes"], r["nombre_salida"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        df = pd.DataFrame(r["reporte"])
        if df.empty:
            st.info("No hubo cambios ni inconsistencias.")
        else:
            st.subheader("Reporte de auditoría")
            f1, f2 = st.columns(2)
            estados = f1.multiselect("Resultado", sorted(df["Resultado"].unique()),
                                     default=sorted(df["Resultado"].unique()))
            columnas = f2.multiselect("Columna", sorted(df["Columna"].unique()),
                                      default=sorted(df["Columna"].unique()))
            filtrado = df[df["Resultado"].isin(estados) & df["Columna"].isin(columnas)]
            st.dataframe(filtrado, use_container_width=True, hide_index=True, height=360)
            st.caption(f"{len(filtrado)} de {len(df)} registros. El reporte completo también va "
                       f"dentro del Excel, en la hoja «{ae.HOJA_REPORTE}».")
