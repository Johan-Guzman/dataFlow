"""Aplicación web: Normalizador de Excel (Streamlit)."""

from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st

import automatizacion_excel as ae
from utilidades_config import (
    COLUMNAS_CATALOGO, ENCABEZADOS, cfg_a_json, cfg_desde_json_bytes,
    construir_cfg, estado_desde_cfg, procesar_bytes,
)

st.set_page_config(page_title="Normalizador de Excel", page_icon=":material/table_chart:", layout="wide")
st.markdown(f"<style>{(Path(__file__).with_name('estilos.css')).read_text(encoding='utf-8')}</style>",
            unsafe_allow_html=True)

ss = st.session_state

# ------------------------------ iconos y piezas visuales ------------------------------
_SVG = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="{w}" '
        'stroke-linecap="round" stroke-linejoin="round">{p}</svg>')
ICONO = {
    "check": _SVG.format(w=3, p='<path d="M5 12l5 5L20 7"/>'),
    "alert": _SVG.format(w=2.4, p='<path d="M12 8v5M12 17h.01"/>'),
    "dot": _SVG.format(w=2.4, p='<circle cx="12" cy="12" r="2"/>'),
    "tabla": _SVG.format(w=2, p='<rect x="3" y="4" width="18" height="16" rx="3"/><path d="M3 10h18M9 10v10"/>'),
}
ICONO_CAMPO = {
    "col_nivel": "school", "col_documento": "badge", "col_programa": "menu_book",
    "col_texto": "text_fields", "col_telefono": "call", "col_codigo_pais": "public", "col_periodo": "event",
}


def chip(texto, tipo, icono=None):
    ic = ICONO[icono] if icono else ""
    return f'<span class="chip chip-{tipo}">{ic}{escape(texto)}</span>'


def stepper_html(pasos):
    partes = []
    for i, (titulo, detalle, estado) in enumerate(pasos, 1):
        num = ICONO["check"] if estado == "done" else str(i)
        partes.append(f'<div class="paso {estado}"><span class="num">{num}</span>'
                      f'<span class="lbl">{titulo}<small>{escape(detalle)}</small></span></div>')
        if i < len(pasos):
            partes.append(f'<span class="linea {"done" if estado == "done" else ""}"></span>')
    return f'<div class="stepper">{"".join(partes)}</div>'


def check_html(filas):
    out = []
    for estado, titulo, detalle in filas:
        ic = {"ok": ICONO["check"], "warn": ICONO["alert"], "pend": ICONO["dot"]}[estado]
        out.append(f'<div class="chk {estado}"><span class="ic">{ic}</span>'
                   f'<span class="tx"><b>{escape(titulo)}</b><span>{escape(detalle)}</span></span></div>')
    return f'<div class="checks">{"".join(out)}</div>'


def seccion(titulo, descripcion):
    st.markdown(f'<div class="sec-t">{titulo}</div><div class="sec-d">{descripcion}</div>',
                unsafe_allow_html=True)


# ------------------------------ estado y configuración ------------------------------
def volcar(cfg):
    for k, v in estado_desde_cfg(cfg).items():
        ss[k] = v


if "iniciado" not in ss:
    volcar(ae.Configuracion())
    ss.iniciado = True


def _al_cargar_config():
    archivo_cfg = ss.get("cfg_upload")
    if archivo_cfg is None:
        return
    try:
        volcar(cfg_desde_json_bytes(archivo_cfg.getvalue()))
        ss.cfg_msg = ("ok", f"Configuración «{archivo_cfg.name}» cargada.")
    except Exception as e:
        ss.cfg_msg = ("error", f"No se pudo cargar la configuración: {e}")


def _restablecer():
    volcar(ae.Configuracion())
    ss.cfg_msg = ("ok", "Se restablecieron los valores por defecto.")


# Estado global (se calcula al inicio para mostrar el progreso en todas las pestañas)
archivo = ss.get("excel_upload")
try:
    cfg = construir_cfg(ss)
    cfg_error = None
except (ValueError, TypeError) as e:
    cfg, cfg_error = None, str(e)
cfg_ok = cfg is not None
vacios = [n for n in cfg.columnas_catalogo if not cfg.valores_validos.get(n)] if cfg_ok else []
r = ss.get("resultado")

# ------------------------------ barra lateral ------------------------------
with st.sidebar:
    st.markdown(f'<div class="brand"><span class="brandmark">{ICONO["tabla"].replace("<svg", "<svg width=18 height=18")}'
                '</span>Normalizador</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-label">CONFIGURACIÓN GUARDADA</div>', unsafe_allow_html=True)
    st.file_uploader("Cargar archivo .json", type=["json"], key="cfg_upload",
                     on_change=_al_cargar_config, label_visibility="collapsed")
    if "cfg_msg" in ss:
        tipo, texto = ss.cfg_msg
        (st.success if tipo == "ok" else st.error)(texto, icon=":material/check_circle:" if tipo == "ok" else ":material/error:")
    if cfg_ok:
        st.download_button("Guardar configuración", cfg_a_json(cfg), "configuracion.json",
                           "application/json", icon=":material/download:", use_container_width=True)
    else:
        st.caption("Corrige la configuración para poder guardarla.")
    st.button("Restablecer valores", on_click=_restablecer, icon=":material/restart_alt:",
              use_container_width=True)
    st.markdown('<div class="nav-label">OPCIONES</div>', unsafe_allow_html=True)
    st.toggle(f"Leer catálogos de la hoja «{ae.HOJA_REGLAS}»", key="leer_hoja",
              help="Si existe, reemplaza los catálogos de esta pantalla cuando los encabezados coinciden.")
    st.toggle("Incluir celdas válidas en el reporte", key="reg_validos",
              help="Por defecto solo se cuentan en la hoja de resumen.")

# ------------------------------ encabezado y progreso ------------------------------
st.markdown(
    '<div class="eyebrow">AUDITORÍA DE DATOS</div>'
    '<h1 class="titulo">Normalizador de Excel</h1>'
    '<p class="sub">Sube tu archivo, define las reglas y descarga el resultado con su reporte. '
    'No se eliminan filas y el original nunca se modifica.</p>',
    unsafe_allow_html=True,
)

s1 = archivo is not None
s2 = cfg_ok
s3 = r is not None
actual = next((i for i, hecho in enumerate((s1, s2, s3)) if not hecho), 2)
estados = ["done" if h else ("current" if i == actual else "pending") for i, h in enumerate((s1, s2, s3))]
st.markdown(stepper_html([
    ("Archivo", archivo.name if s1 else "Sin archivo", estados[0]),
    ("Reglas", "Configuración válida" if s2 else "Revisa los errores", estados[1]),
    ("Resultado", "Listo para descargar" if s3 else "Pendiente", estados[2]),
]), unsafe_allow_html=True)

if cfg_error:
    st.error(f"Configuración inválida: {cfg_error}", icon=":material/error:")

seccion("Archivo de entrada", "Arrastra aquí tu .xlsx o .xlsm, o pulsa Upload. Se procesa en memoria.")
archivo = st.file_uploader(
    "Archivo Excel", type=["xlsx", "xlsm"], key="excel_upload",
    label_visibility="collapsed",
)

# ------------------------------ pestañas ------------------------------
t_col, t_cat, t_per, t_run = st.tabs([
    ":material/view_column: Columnas",
    ":material/menu_book: Catálogos",
    ":material/event: Periodos",
    ":material/play_arrow: Ejecutar",
])

with t_col:
    seccion("Encabezados de las columnas",
            "Escribe el encabezado (fila 1) de cada columna. Se ignoran mayúsculas, tildes, espacios y «_».")
    with st.container(border=True, key="card_columnas"):
        c1, c2 = st.columns(2)
        for i, (clave, etiqueta) in enumerate(ENCABEZADOS):
            (c1 if i % 2 == 0 else c2).text_input(f":material/{ICONO_CAMPO[clave]}: {etiqueta}", key=clave)

with t_cat:
    seccion("Catálogos y equivalencias",
            "Lista oficial (un valor por línea) y equivalencias confirmadas con el formato "
            "<code>CC = Cédula de ciudadanía</code>. Con la lista vacía no se corrige nada.")
    iconos_cat = {"nivel": "school", "documento": "badge", "programa": "menu_book"}
    for clave, etiqueta in COLUMNAS_CATALOGO:
        n_val = len([x for x in ss[f"cat_{clave}"].splitlines() if x.strip()])
        n_eq = len([x for x in ss[f"eq_{clave}"].splitlines() if x.strip()])
        estado_chip = (chip(f"{n_val} valores · {n_eq} equivalencias", "ok", "check") if n_val
                       else chip("Pendiente", "warn", "alert"))
        with st.container(border=True, key=f"card_cat_{clave}"):
            st.markdown(f'<div class="card-h"><b>{etiqueta}</b>{estado_chip}</div>', unsafe_allow_html=True)
            a, b = st.columns(2)
            a.text_area("Valores oficiales", key=f"cat_{clave}", height=140, placeholder="Un valor por línea")
            b.text_area("Equivalencias (clave = valor oficial)", key=f"eq_{clave}", height=140,
                        placeholder="CC = Cédula de ciudadanía")

with t_per:
    seccion("Periodos académicos",
            "Código de 2 dígitos = descripción. Los códigos que no estén aquí no reciben descripción.")
    with st.container(border=True, key="card_periodos"):
        a, b = st.columns([2, 1])
        a.text_area("Códigos y descripciones", key="periodos", height=150, placeholder="10 = Primer Semestre")
        b.number_input("Año mínimo válido", key="anio_min", step=1)
        b.number_input("Año máximo válido", key="anio_max", step=1)

with t_run:
    with st.expander("Verificación previa", expanded=ss.get("resultado") is None, icon=":material/fact_check:"):
        st.markdown(check_html([
            ("ok", "Archivo cargado", archivo.name) if archivo else ("pend", "Archivo pendiente", "Súbelo en la parte superior."),
            ("ok", "Configuración válida", "Encabezados, equivalencias y periodos correctos.") if cfg_ok
            else ("warn", "Configuración con errores", cfg_error or ""),
            ("warn", "Catálogos sin definir", ", ".join(vacios) + ": se marcarán como SIN_CATALOGO.") if vacios
            else ("ok", "Catálogos definidos", "Todos los valores se compararán con la lista oficial.")
            if cfg_ok else ("pend", "Catálogos", "Pendiente de una configuración válida."),
        ]), unsafe_allow_html=True)

    if st.button("Ejecutar automatización", type="primary", icon=":material/play_arrow:",
                 disabled=not (cfg_ok and archivo), use_container_width=True):
        try:
            with st.spinner("Procesando…"):
                ss.resultado = procesar_bytes(archivo.getvalue(), archivo.name, cfg)
            ss.pop("error_proc", None)
        except Exception as e:
            ss.pop("resultado", None)
            ss.error_proc = str(e)
        st.rerun()  # refresca el progreso superior con el nuevo estado
    if ss.get("error_proc"):
        st.error(f"Error al procesar: {ss.error_proc}", icon=":material/error:")

    r = ss.get("resultado")
    if r:
        st.divider()
        seccion("Resultado", f"Archivo generado: <b>{escape(r['nombre_salida'])}</b>")
        res = r["resumen"]
        tarjetas = [("Corregidos", "CORREGIDO", "#5454e9"), ("Válidos", "VALIDO", "#4cb979"),
                    ("Inválidos", "INVALIDO", "#c15a3e"), ("Por revisar", "REQUIERE_REVISION", "#e0a458"),
                    ("Sin catálogo", "SIN_CATALOGO", "#b3b1bf")]
        st.markdown('<div class="stats">' + "".join(
            f'<div class="stat"><div class="k"><i style="background:{c}"></i>{t}</div>'
            f'<strong>{res.get(k, 0)}</strong></div>' for t, k, c in tarjetas) + "</div>",
            unsafe_allow_html=True)
        for aviso in r["avisos"]:
            st.warning(aviso, icon=":material/warning:")
        st.download_button(
            "Descargar Excel corregido", r["bytes"], r["nombre_salida"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", icon=":material/download:", use_container_width=True,
        )
        df = pd.DataFrame(r["reporte"])
        if df.empty:
            st.info("No hubo cambios ni inconsistencias.", icon=":material/info:")
        else:
            st.markdown('<div class="sec-t" style="margin-top:22px">Reporte de auditoría</div>',
                        unsafe_allow_html=True)
            f1, f2 = st.columns(2)
            estados_f = f1.multiselect("Resultado", sorted(df["Resultado"].unique()),
                                       default=sorted(df["Resultado"].unique()))
            columnas_f = f2.multiselect("Columna", sorted(df["Columna"].unique()),
                                        default=sorted(df["Columna"].unique()))
            filtrado = df[df["Resultado"].isin(estados_f) & df["Columna"].isin(columnas_f)]
            colores = {"CORREGIDO": ("#eaeaff", "#5454e9"), "VALIDO": ("#e9f5ee", "#4a9e75"),
                       "INVALIDO": ("#fdf1ee", "#c15a3e"), "REQUIERE_REVISION": ("#fff4e0", "#a6772a")}

            def pintar(v):
                bg, fg = colores.get(v, ("#f0f0f8", "#6d6d7c"))
                return f"background-color:{bg};color:{fg};font-weight:600"

            st.dataframe(filtrado.style.map(pintar, subset=["Resultado"]),
                         use_container_width=True, hide_index=True, height=360)
            st.caption(f"{len(filtrado)} de {len(df)} registros. El reporte completo también está en "
                       f"la hoja «{ae.HOJA_REPORTE}» del Excel.")
