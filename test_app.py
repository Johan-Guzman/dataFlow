from pathlib import Path

import pytest

import automatizacion_excel as ae
import utilidades_config as uc
from crear_excel_prueba import crear_excel_prueba

CONFIG = Path(__file__).parent / "config_prueba.json"


def test_procesar_bytes(tmp_path):
    ruta = crear_excel_prueba(tmp_path / "datos.xlsx")
    cfg = ae.cargar_configuracion_json(CONFIG)
    r = uc.procesar_bytes(ruta.read_bytes(), "datos.xlsx", cfg)
    assert r["nombre_salida"] == "datos_corregido.xlsx"
    assert r["bytes"][:2] == b"PK" and r["resumen"]["CORREGIDO"] > 0
    with pytest.raises(ValueError):
        uc.procesar_bytes(b"x", "datos.csv", cfg)


def test_ida_y_vuelta_json():
    cfg = ae.cargar_configuracion_json(CONFIG)
    cfg2 = uc.cfg_desde_json_bytes(uc.cfg_a_json(cfg).encode("utf-8"))
    assert uc.construir_cfg(uc.estado_desde_cfg(cfg2)).equivalencias == cfg.equivalencias


def test_app_se_renderiza_y_valida():
    at_mod = pytest.importorskip("streamlit.testing.v1")
    app = at_mod.AppTest.from_file("app.py", default_timeout=30).run()
    assert not app.exception
    assert len(app.tabs) == 4
    # editar un encabezado a vacío -> mensaje de configuración inválida
    app.text_input(key="col_periodo").set_value("").run()
    assert any("inválida" in e.value for e in app.error)
    # equivalencia mal formada
    app.text_input(key="col_periodo").set_value("periodo")
    app.text_area(key="eq_documento").set_value("sin igual").run()
    assert any("clave = valor" in e.value for e in app.error)
    # configuración válida editada desde la interfaz
    app.text_area(key="eq_documento").set_value("CC = Cédula de ciudadanía").run()
    assert not app.error and not app.exception


def test_convertir_bytes_y_perfil():
    ae_ = ae
    cfg = ae_.cargar_configuracion_json(Path(__file__).parent / "config_aspirantes.json")
    cfg2 = uc.construir_cfg(uc.estado_desde_cfg(cfg))
    assert cfg2.mapeo == cfg.mapeo and cfg2.periodos_permitidos == cfg.periodos_permitidos
    guia = Path(__file__).parent / "guia_anonimizada.xlsx"
    r = uc.convertir_bytes(guia.read_bytes(), "guia.xlsx",
                           Path(__file__).parent / "plantilla_importacion_Aspirante.xlsx", cfg2)
    assert r["filas_salida"] == 26 and r["nombre_salida"] == "guia_importacion.xlsx"
    assert r["bytes"][:2] == b"PK" and r["bytes_reporte"][:2] == b"PK"


def test_app_modo_plantilla():
    at_mod = pytest.importorskip("streamlit.testing.v1")
    app = at_mod.AppTest.from_file("app.py", default_timeout=30).run()
    app.radio(key="modo").set_value("Convertir a plantilla de importación").run()
    assert not app.exception and not app.error
    assert len([t for t in app.text_input if t.key and t.key.startswith("map_")]) == 12
    app.text_input(key="map_0").set_value("").run()  # mapeo vacío no rompe la app
    assert not app.exception


def test_construir_cfg_sin_campos_de_columnas():
    """En modo plantilla Streamlit descarta col_*: debe usar los valores por defecto."""
    cfg = ae.cargar_configuracion_json(Path(__file__).parent / "config_aspirantes.json")
    estado = uc.estado_desde_cfg(cfg)
    for clave, _ in uc.ENCABEZADOS:
        estado.pop(clave)
    assert uc.construir_cfg(estado).col_nivel == "nivel educativo"
