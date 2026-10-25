# Normalizador de Excel – despliegue

## Probar en tu equipo
    pip install -r requirements.txt
    streamlit run app.py

## Publicar gratis (Streamlit Community Cloud)
1. Crea un repositorio en GitHub y sube TODO el contenido de esta carpeta
   (incluida `.streamlit/config.toml`).
2. Entra a https://share.streamlit.io, inicia sesión con GitHub y elige
   "Create app" -> tu repositorio -> archivo principal `app.py` -> Deploy.
3. Comparte el enlace que te entrega.

## Archivos de diseño
- `estilos.css`: todos los colores, tamaños y estilos (paleta, tipografía, tarjetas, pestañas).
  Los colores están como variables al inicio del archivo (`:root`).
- `.streamlit/config.toml`: colores base del tema de Streamlit.

## Notas
- Los archivos se procesan en memoria; no se guardan en el servidor.
- Si el repositorio es público, NO subas Excel reales ni configuraciones con datos
  sensibles (el código y `config_prueba.json` sí pueden ser públicos). Para uso
  interno, usa un repositorio privado o un servidor propio.
- Cualquiera con el enlace puede usar la app. Si necesitas restringir el acceso,
  configura la app como privada en Streamlit Cloud.
- Para ejecutar las pruebas: `pip install pytest` y luego `pytest`.
