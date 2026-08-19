# UdeSA Horarios

Aplicación web para consulta integrada de aulas, ocupaciones y disponibilidad estudiantil.

**Desarrollado por Juan Ignacio Gutiérrez Julián.**

## Datos

Este repositorio no contiene fuentes institucionales. La aplicación sincroniza archivos privados mediante un bridge de Google Apps Script configurado en Streamlit Secrets.

## Deploy

- Entrypoint: `app.py`
- Framework: Streamlit 1.60.0
- Secrets requeridos:

```toml
[source_bridge]
url = "https://script.google.com/macros/s/.../exec"
token = "..."
```

Nunca subir `secrets.toml`, archivos `.xlsx` reales o configuraciones privadas del bridge.
