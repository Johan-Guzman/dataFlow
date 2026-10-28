import json
from pathlib import Path

import pytest

import automatizacion_excel as ae

tk = pytest.importorskip("tkinter")
import interfaz_automatizacion as ui

CONFIG = Path(__file__).parent / "config_prueba.json"


def test_parsear_pares():
    assert ui.parsear_pares("CC = Cédula\n\nTI=Tarjeta = x", "t") == {"CC": "Cédula", "TI": "Tarjeta = x"}
    with pytest.raises(ValueError):
        ui.parsear_pares("sin igual", "t")
    with pytest.raises(ValueError):
        ui.parsear_pares("cc = A\nCC = B", "t")  # equivalentes con destinos distintos
    with pytest.raises(ValueError):
        ui.parsear_pares("CC = ", "t")


@pytest.fixture
def app():
    try:
        a = ui.App()
    except tk.TclError:
        pytest.skip("No hay pantalla disponible")
    a.withdraw()
    yield a
    a.destroy()


def test_ida_y_vuelta_de_configuracion(app):
    cfg0 = ae.cargar_configuracion_json(CONFIG)
    app.aplicar_cfg(cfg0)
    cfg1 = app.obtener_cfg()
    assert cfg1.valores_validos == cfg0.valores_validos
    assert cfg1.equivalencias == cfg0.equivalencias
    assert cfg1.periodos == cfg0.periodos


def test_validaciones(app):
    app.aplicar_cfg(ae.Configuracion())
    app.v["col_periodo"].set("")
    with pytest.raises(ValueError):
        app.obtener_cfg()
    app.v["col_periodo"].set("periodo")
    app._set_texto("periodos", "1 = Primer")
    with pytest.raises(ValueError):
        app.obtener_cfg()
    app._set_texto("periodos", "10 = Primer Semestre")
    app.v["anio_min"].set("3000")
    with pytest.raises(ValueError):
        app.obtener_cfg()


def test_ejecucion_desde_interfaz_equivale_al_script(app, tmp_path):
    from crear_excel_prueba import crear_excel_prueba
    entrada = crear_excel_prueba(tmp_path / "in.xlsx")
    app.aplicar_cfg(ae.cargar_configuracion_json(CONFIG))
    # edición hecha "en la interfaz": una equivalencia nueva
    app._set_texto("eq_documento", "CC = Cédula de ciudadanía\nRC = Pasaporte")
    r = ae.procesar_excel(entrada, tmp_path / "out.xlsx", app.obtener_cfg())
    assert any(x["Valor original"] == "RC" and x["Valor final"] == "Pasaporte" for x in r["reporte"])
    # guardado/lectura JSON
    app._guardar_json(tmp_path / "c.json", app.obtener_cfg())
    assert ae.cargar_configuracion_json(tmp_path / "c.json").equivalencias["tipo de documento"]["RC"] == "Pasaporte"
