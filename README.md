# Normalizador de Excel

Herramienta en Python para normalizar datos de un archivo Excel **sin eliminar registros ni tocar información ajena a las reglas**. Incluye una aplicación web (Streamlit), una ventana de escritorio opcional (Tkinter) y una línea de comandos. Todas usan el mismo motor.

El archivo original nunca se modifica: se genera un archivo nuevo con una hoja de auditoría que explica cada cambio.

---

## Qué hace

| Columna | Regla |
|---|---|
| Nivel educativo | Ajusta al valor oficial del catálogo (ignora mayúsculas, tildes y espacios). |
| Tipo de documento | Ajusta al catálogo y aplica equivalencias confirmadas (`CC` → `Cédula de ciudadanía`). |
| Programa de interés | Ajusta al nombre oficial (`Ingenieria de Sistemas` → `Ingeniería de Sistemas`). |
| Categoría de servicio | Reemplaza `_` por espacios y aplica Title Case (`CONSULTA__EXTERNA` → `Consulta Externa`). |
| Teléfono / Código de país | Separa `+57` del número nacional y lo guarda en `Codigo_Pais`. |
| Periodo | `202710` → `202710 - Primer Semestre 2027`. |

### Principios de seguridad de los datos
- **No se inventa nada.** Un valor que no está en el catálogo, una abreviatura sin equivalencia o un código de periodo sin descripción se **conservan** y se registran para revisión.
- **Sin coincidencias aproximadas.** Solo se corrige cuando hay una coincidencia exacta tras normalizar mayúsculas, tildes y espacios. La `ñ` se respeta (`Diseño` ≠ `Diseno`).
- **Los vacíos siguen vacíos.** No se completan ni se reportan.
- **Se conservan** hojas, encabezados, orden de filas, fórmulas y formatos. No se eliminan duplicados ni filas con inconsistencias.
- **Teléfonos:** no se convierten a número, se conservan los ceros iniciales y los espacios internos, y no se agrega `+57` a quien no lo tenga. Si `Codigo_Pais` ya trae otro valor, no se sobrescribe y se registra.
- Si el archivo es `.xlsm`, se guarda como `.xlsm` para no perder las macros.

---

## Contenido del paquete

```
app.py                      Aplicación web (Streamlit)
estilos.css                 Diseño de la app (paleta, tipografía, tarjetas)
automatizacion_excel.py     Motor de normalización y línea de comandos
utilidades_config.py        Lógica compartida por las interfaces
interfaz_automatizacion.py  Ventana de escritorio opcional (Tkinter)
config_prueba.json          Catálogos FICTICIOS para probar
crear_excel_prueba.py       Genera un Excel de prueba con datos ficticios
test_*.py                   Pruebas automáticas (pytest)
requirements.txt            Dependencias
.streamlit/config.toml      Tema base de Streamlit
```

---

## Instalación

Requiere Python 3.10 o superior.

```bash
pip install -r requirements.txt
```

---

## Uso

### 1. Aplicación web (recomendada)

```bash
streamlit run app.py
```

1. **Sube** tu Excel (`.xlsx` o `.xlsm`).
2. **Configura** las reglas en las pestañas *Columnas*, *Catálogos* y *Periodos*.
3. Ve a **Ejecutar**, revisa la verificación previa y pulsa *Ejecutar automatización*.
4. **Descarga** el Excel corregido y consulta el reporte de auditoría (con filtros).

La barra lateral permite cargar y guardar la configuración en JSON, restablecer valores y activar opciones. El paso a paso de la parte superior indica en qué punto estás y si hay errores.

### 2. Línea de comandos

```bash
python automatizacion_excel.py --entrada datos_originales.xlsx --salida datos_corregidos.xlsx --config mi_config.json
```

Sin `--config` se usa la configuración por defecto del script (catálogos vacíos).

### 3. Ventana de escritorio (opcional)

```bash
python interfaz_automatizacion.py
```

Requiere Tkinter (incluido en Python; en algunas instalaciones de Linux: `sudo apt install python3-tk`).

---

## Configuración

Los catálogos oficiales **no vienen definidos**: debes completarlos. Mientras un catálogo esté vacío, los valores de esa columna se conservan y se marcan `SIN_CATALOGO`.

Puedes configurar todo desde la app, desde un archivo JSON o editando la sección 1 de `automatizacion_excel.py`.

### Formato del JSON

```json
{
  "col_nivel": "nivel educativo",
  "col_documento": "tipo de documento",
  "col_programa": "programa de interés",
  "col_texto": "categoria_servicio",
  "col_telefono": "telefono",
  "col_codigo_pais": "Codigo_Pais",
  "col_periodo": "periodo",
  "valores_validos": {
    "nivel educativo": ["Bachillerato", "Técnico", "Tecnólogo", "Profesional"],
    "tipo de documento": ["Cédula de ciudadanía", "Tarjeta de identidad"],
    "programa de interés": ["Ingeniería de Sistemas"]
  },
  "equivalencias": {
    "tipo de documento": { "CC": "Cédula de ciudadanía", "TI": "Tarjeta de identidad" }
  },
  "periodos": { "10": "Primer Semestre", "20": "Segundo Semestre" },
  "anio_min": 2000,
  "anio_max": 2100,
  "leer_catalogos_desde_hoja": false,
  "registrar_validos": false
}
```

- Las claves de `valores_validos` y `equivalencias` deben coincidir con los encabezados definidos en `col_*`.
- Los encabezados se comparan ignorando mayúsculas, tildes, espacios y guiones bajos.
- Una equivalencia debe apuntar a un valor que exista en el catálogo; si no, el registro va a revisión.
- `leer_catalogos_desde_hoja`: si es `true` y el libro tiene una hoja llamada `Reglas de validación`, lee los catálogos de ella (solo columnas cuyo encabezado coincida exactamente). Está desactivado por defecto.
- `registrar_validos`: si es `true`, el reporte incluye también cada celda ya válida. Por defecto solo se cuentan en el resumen.

### Periodos
Se extrae el año de los 4 primeros dígitos y el código de los 2 últimos. Se valida que el año esté entre `anio_min` y `anio_max` y que el código esté configurado. Un periodo que ya viene con descripción solo se acepta si coincide con el código; si no, va a revisión.

---

## Reporte de auditoría

El Excel de salida incluye dos hojas nuevas:

- **`Reporte_validacion`**: `Hoja · Fila · Columna · Valor original · Valor final · Resultado · Motivo`
- **`Resumen_validacion`**: conteo por hoja, columna y resultado (incluye los válidos).

| Resultado | Significado |
|---|---|
| `CORREGIDO` | Se cambió el valor; el motivo explica por qué. |
| `VALIDO` | Ya estaba correcto (solo se cuenta en el resumen). |
| `INVALIDO` | No coincide con el catálogo; se conserva. |
| `REQUIERE_REVISION` | Ambiguo o inconsistente (p. ej. `+57` sin número, código de país distinto, periodo no configurado); se conserva. |
| `SIN_CATALOGO` | El catálogo de esa columna está vacío; una sola fila de resumen por hoja y columna. |
| `COLUMNA_NO_ENCONTRADA` | No se halló el encabezado configurado en esa hoja. |

---

## Pruebas

```bash
pip install pytest
pytest
```

Las pruebas usan datos ficticios y cubren los casos de normalización, los casos límite (vacíos, `+57` solo, código de país en conflicto, periodos no configurados), que no se eliminen filas, que las columnas ajenas y las fórmulas queden intactas, y que el original no cambie. Las pruebas de la ventana de escritorio se omiten si no hay Tkinter o pantalla.

---

### Privacidad
- Los archivos se procesan **en memoria** y no se guardan en el servidor.
- Cualquiera con el enlace puede usar la app. Si maneja datos de personas, configúrala como privada o usa un servidor propio.
- **No subas Excel reales ni configuraciones sensibles a un repositorio público.** `config_prueba.json` contiene solo valores ficticios.

---

## Limitaciones

- openpyxl guarda las fórmulas sin su valor calculado; Excel las recalcula al abrir el archivo.
- Gráficos e imágenes pueden no conservarse al guardar; el programa lo avisa por hoja.
- Solo se leen encabezados de la **fila 1**.
- Las celdas con fórmula o combinadas en las columnas procesadas no se modifican y se registran para revisión.
- Tamaño máximo de subida en la app: 50 MB (configurable en `.streamlit/config.toml`).