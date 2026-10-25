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
