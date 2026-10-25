"""Lógica compartida por las interfaces (sin dependencias de pantalla)."""

import json
import tempfile
from pathlib import Path

import automatizacion_excel as ae

COLUMNAS_CATALOGO = [
    ("nivel", "Nivel educativo"),
    ("documento", "Tipo de documento"),
    ("programa", "Programa de interés"),
]
ENCABEZADOS = [
    ("col_nivel", "Nivel educativo"),
    ("col_documento", "Tipo de documento"),
    ("col_programa", "Programa de interés"),
    ("col_texto", "Texto a normalizar (Title Case)"),
    ("col_telefono", "Teléfono"),
    ("col_codigo_pais", "Código de país"),
    ("col_periodo", "Periodo"),
]


def lineas(texto):
    return [l.strip() for l in texto.splitlines() if l.strip()]


def parsear_pares(texto, etiqueta):
    """'clave = valor' por línea -> dict. Rechaza líneas mal formadas y conflictos."""
    pares, vistos = {}, {}
    for linea in lineas(texto):
        if "=" not in linea:
            raise ValueError(f"{etiqueta}: la línea «{linea}» debe tener formato clave = valor")
        k, v = (x.strip() for x in linea.split("=", 1))
        if not k or not v:
            raise ValueError(f"{etiqueta}: la línea «{linea}» tiene la clave o el valor vacío")
        nk = ae.normalizar_comparacion(k)
        if nk in vistos and pares[vistos[nk]] != v:
            raise ValueError(
                f"{etiqueta}: «{vistos[nk]}» y «{k}» son equivalentes pero apuntan a valores distintos"
            )
        vistos[nk] = k
        pares[k] = v
    return pares


def pares_a_texto(d):
    return "\n".join(f"{k} = {v}" for k, v in d.items())


def estado_desde_cfg(cfg):
    """Configuracion -> dict plano con las claves que usa el formulario."""
    e = {clave: getattr(cfg, clave) for clave, _ in ENCABEZADOS}
    for clave, _ in COLUMNAS_CATALOGO:
        nombre = getattr(cfg, f"col_{clave}")
        e[f"cat_{clave}"] = "\n".join(cfg.valores_validos.get(nombre, []))
        e[f"eq_{clave}"] = pares_a_texto(cfg.equivalencias.get(nombre, {}))
    e.update(
        periodos=pares_a_texto(cfg.periodos),
        anio_min=int(cfg.anio_min),
        anio_max=int(cfg.anio_max),
        leer_hoja=bool(cfg.leer_catalogos_desde_hoja),
        reg_validos=bool(cfg.registrar_validos),
    )
    return e


def construir_cfg(e):
    """Dict/estado del formulario -> Configuracion validada (ValueError si hay errores)."""
    cfg = ae.Configuracion()
    nombres = []
    for clave, etiqueta in ENCABEZADOS:
        valor = str(e[clave]).strip()
        if not valor:
            raise ValueError(f"El encabezado «{etiqueta}» no puede estar vacío.")
        setattr(cfg, clave, valor)
        nombres.append(ae.normalizar_encabezado(valor))
    if len(set(nombres)) != len(nombres):
        raise ValueError("Hay dos columnas con el mismo encabezado.")

    cfg.valores_validos, cfg.equivalencias = {}, {}
    for clave, etiqueta in COLUMNAS_CATALOGO:
        nombre = getattr(cfg, f"col_{clave}")
        cfg.valores_validos[nombre] = lineas(e[f"cat_{clave}"])
        cfg.equivalencias[nombre] = parsear_pares(e[f"eq_{clave}"], f"Equivalencias de {etiqueta}")

    cfg.periodos = parsear_pares(e["periodos"], "Periodos")
    for cod in cfg.periodos:
        if not (cod.isdigit() and len(cod) == 2):
            raise ValueError(f"Periodos: el código «{cod}» debe tener exactamente 2 dígitos.")
    cfg.anio_min, cfg.anio_max = int(e["anio_min"]), int(e["anio_max"])
    if cfg.anio_min > cfg.anio_max:
        raise ValueError("El año mínimo no puede ser mayor que el máximo.")
    cfg.leer_catalogos_desde_hoja = bool(e["leer_hoja"])
    cfg.registrar_validos = bool(e["reg_validos"])
    return cfg


def cfg_a_json(cfg):
    from dataclasses import asdict
    return json.dumps(asdict(cfg), ensure_ascii=False, indent=2)


def cfg_desde_json_bytes(contenido):
    with tempfile.TemporaryDirectory() as d:
        ruta = Path(d) / "config.json"
        ruta.write_bytes(contenido)
        return ae.cargar_configuracion_json(ruta)


def procesar_bytes(contenido, nombre, cfg):
    """Procesa un Excel en memoria. Devuelve el resultado + bytes del archivo corregido."""
    entrada_nombre = Path(nombre).name
    sufijo = Path(entrada_nombre).suffix.lower()
    if sufijo not in (".xlsx", ".xlsm"):
        raise ValueError("Solo se aceptan archivos .xlsx o .xlsm.")
    with tempfile.TemporaryDirectory() as d:
        entrada = Path(d) / entrada_nombre
        entrada.write_bytes(contenido)
        salida = Path(d) / f"{entrada.stem}_corregido{sufijo}"
        r = ae.procesar_excel(entrada, salida, cfg)
        r["bytes"] = Path(r["salida"]).read_bytes()
        r["nombre_salida"] = Path(r["salida"]).name
    del r["salida"]
    return r
