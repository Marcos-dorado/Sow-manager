# SOW Manager

Automatización de escritorio (Windows) que lee los correos de SOWs (Statements of Work) de
NEORIS/CHUBB Colombia desde Outlook, extrae la información de los documentos adjuntos, archiva los
archivos organizados por persona/año y actualiza el Excel maestro de seguimiento
(`Listado SOW's.xlsm`) sin corromper sus tablas dinámicas, Power Query, slicers ni macros VBA.

Se distribuye como ejecutable standalone (`SOW Manager.exe`, generado con PyInstaller) para que la
persona que lo opera no necesite tener Python instalado.

## Documentación

- **[MANUAL_DE_USO.md](MANUAL_DE_USO.md)** — para quien opera la app día a día: qué resuelve,
  requisitos, paso a paso de uso normal, errores comunes y buenas prácticas.
- **[MANUAL_TECNICO.md](MANUAL_TECNICO.md)** — para quien mantenga o extienda el código: stack,
  estructura del proyecto, flujo de datos completo, el motor de escritura OOXML sobre el `.xlsm`
  (la parte más compleja del proyecto), cómo hacer modificaciones comunes y cómo levantar el
  entorno desde cero.

## Setup rápido (desarrollo)

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

Requiere Windows, con Microsoft Outlook de escritorio instalado y con sesión iniciada — ver
requisitos completos en `MANUAL_TECNICO.md`.
