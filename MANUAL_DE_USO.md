# Manual de Uso — SOW Manager

Automatización de escritorio (Windows) que lee los correos de SOWs de CHUBB desde Outlook, extrae la información de los documentos adjuntos, los archiva en carpetas organizadas y actualiza el Excel maestro de seguimiento (`Listado SOW's.xlsm`) — sin necesidad de hacerlo a mano.

---

## 1. Qué problema resuelve

**Antes de esta herramienta**, el proceso era completamente manual: alguien debía revisar cada correo de CHUBB, abrir el documento Word o PDF adjunto para sacar fechas y valores, copiar todo a mano en el Excel `Listado SOW's.xlsm`, guardar el archivo en la carpeta correcta, y hacer seguimiento de cuáles SOWs estaban pendientes de firma. Es un proceso repetitivo y con riesgo de error de tipeo, filas duplicadas o archivos mal guardados.

**Con SOW Manager**, la aplicación hace todo eso automáticamente:

- Se conecta a Outlook y lee los correos de la carpeta "SOWs".
- Detecta si es un SOW nuevo (documento Word) o un PDF ya firmado.
- Archiva los documentos organizados por persona y año en la carpeta del proyecto.
- Extrae fechas, valor y hace un **análisis automático de 6 puntos** del documento (ver sección 8).
- Actualiza el Excel maestro con los datos extraídos — sin tocar Power Query, tablas dinámicas, slicers ni macros del archivo.
- Muestra los SOWs pendientes de firma para hacer seguimiento fácilmente.

> ℹ️ La aplicación nunca borra correos ni modifica el Excel de forma irreversible. Todo cambio pasa por la revisión de la persona antes de guardarse — sincronizar solo muestra lo encontrado, guardar es un paso aparte y deliberado.

---

## 2. Requisitos para correrlo

| Requisito | Detalle |
|---|---|
| **Windows** | La aplicación solo funciona en Windows (no en Mac ni Linux). |
| **Microsoft Outlook** | Debe estar instalado, configurado con la cuenta corporativa (`@epamneoris.com`) y **completamente sincronizado** antes de abrir la app. Tiene que ser el Outlook de escritorio clásico — no funciona con Outlook Web ni con la app "nueva" de Outlook. |
| **Outlook abierto** | Debe permanecer abierto mientras se usa SOW Manager. |
| **Acceso al Excel maestro** | El archivo `Listado SOW's.xlsm` debe estar accesible en la ruta del proyecto dentro de "EPAM Documents", con una hoja llamada exactamente `SOWS` (encabezados en la fila 7, datos desde la fila 8). |
| **Carpeta del proyecto sincronizada** | La carpeta `0. Documentacion Contractual` debe estar sincronizada con SharePoint/OneDrive en el equipo. |
| **El ejecutable `SOW Manager.exe`** | No requiere instalación ni tener Python — se ejecuta directamente desde la carpeta `SOW Manager` con doble clic. |

---

## 3. Primer uso — Configuración inicial

La primera vez que se usa SOW Manager en un computador nuevo:

**Paso 1 — Configurar Outlook.** Si Outlook nunca fue configurado en ese equipo, al abrirlo aparece el asistente de bienvenida:
1. Seleccionar Microsoft 365 e ingresar el correo corporativo (`@epamneoris.com`).
2. Si pide contraseña y la empresa usa autenticación por QR o SSO y eso falla: cerrar esa ventana, volver a abrir Outlook y elegir "Configuración manual o tipos de servidores adicionales", luego Microsoft 365.
3. Completar el proceso y hacer clic en Finalizar.
4. Esperar a que Outlook sincronice completamente antes de continuar.

> ⚠️ Si Outlook acaba de configurarse por primera vez, esperar a que la barra de estado diga **"Todas las carpetas están actualizadas"** antes de abrir SOW Manager. Puede tardar varios minutos según el tamaño del buzón.

**Paso 2 — Abrir SOW Manager.** Abrir la carpeta `SOW Manager` y ejecutar `SOW Manager.exe`.

**Paso 3 — Primera sincronización.** Al hacer clic en Sincronizar por primera vez, la aplicación crea automáticamente la carpeta "SOWs" dentro de la Bandeja de Entrada — no hace falta crearla a mano.

> ⚠️ Si la carpeta "SOWs" no aparece visible de inmediato en Outlook, no es un error: Outlook puede tardar unos segundos en refrescar la vista. Cerrar y volver a abrir Outlook la muestra.

---

## 4. La interfaz de la aplicación

| Área | Descripción |
|---|---|
| **Encabezado (arriba)** | Nombre de la app, fecha de última actualización, y los botones principales. |
| **Pestaña 🟡 Pendientes** | Lista todos los SOWs ya registrados en el Excel que todavía no tienen PDF firmado — los que están esperando firma. |
| **Pestaña 🔵 Novedades** | Aparece después de sincronizar. Muestra los SOWs nuevos encontrados en los correos que aún no se guardaron en el Excel. |
| **Barra de estado (abajo)** | Mensajes de progreso mientras la app trabaja (conectando, procesando, etc.). |

**Botones principales:**

| Botón | Qué hace |
|---|---|
| 🔄 **Sincronizar** | Lee los correos de la carpeta "SOWs" en Outlook y detecta novedades (SOWs nuevos o PDFs firmados). Tarda entre 15 y 35 segundos — fuerza una sincronización real con el servidor antes de leer nada. |
| 📥 **Guardar en Excel** | Guarda TODAS las novedades encontradas en el Excel. |
| ☑ **Guardar selección** | Guarda solo los ítems seleccionados de la lista de Novedades (clic + Ctrl para varios). |
| 📊 | Abre directamente el archivo Excel maestro. |
| 📂 | Abre la carpeta raíz donde se archivan los documentos SOW. |

---

## 5. Formato del correo de CHUBB

Para que SOW Manager procese correctamente los correos, CHUBB debe enviarlos siguiendo estas reglas:

- El correo debe incluir siempre una tabla con **3 columnas**: número de SOW, empresa, y la(s) persona(s) — si el SOW aplica para varias personas, se listan una por línea dentro de la misma celda. Ejemplo real de una fila de esa tabla:

  | SOW | Empresa | Persona(s) |
  |---|---|---|
  | NEO-SERV-2026-0179-CR01 | NEORIS Colombia SAS | Cardenas Calderon, Diego Fernando |

**Reglas importantes:**
- Siempre adjuntar el documento: Word la primera vez, PDF firmado cuando esté aprobado.
- La tabla debe estar presente en **ambos** casos: tanto cuando se envía el Word como cuando se envía el PDF.
- El número de SOW en la tabla debe coincidir **exactamente** con el nombre del archivo adjunto.
- Cada envío debe ser un **correo nuevo**, no una respuesta ni dentro de un hilo existente.
- El correo puede tener cualquier texto o saludo — la app solo lee la tabla y el adjunto.

> 🚫 Si el correo llega sin tabla, o el número de SOW no coincide con el nombre del archivo, la app no puede procesarlo y el SOW no aparecerá en Novedades (queda como "Sin match" en el registro de la sincronización).

---

## 6. Flujo de trabajo diario

**Cuando llega un correo con un SOW nuevo:**
1. CHUBB envía un correo con la tabla de datos y el documento Word adjunto.
2. Mover ese correo a la carpeta "SOWs" dentro de la Bandeja de Entrada en Outlook.
3. En SOW Manager, hacer clic en 🔄 Sincronizar.
4. El SOW nuevo aparece en la pestaña Novedades con ícono 🔵 (Word nuevo).
5. Doble clic sobre la fila para revisar el detalle y el análisis de 6 puntos del documento (sección 8).
6. Hacer clic en 📥 Guardar en Excel (o ☑ Guardar selección para guardar solo algunos).
7. El SOW queda registrado y aparece en Pendientes hasta que llegue el PDF firmado.

**Cuando CHUBB envía el PDF firmado:**
1. CHUBB envía un correo **nuevo** con la tabla de datos y el PDF firmado adjunto.
2. Mover ese correo a la carpeta "SOWs" en Outlook.
3. Sincronizar en SOW Manager.
4. El PDF aparece en Novedades con ícono ✅ (PDF firmado).
5. Guardar en Excel — el SOW desaparece de Pendientes (queda cerrado).

> ℹ️ Siempre mover el correo a la carpeta "SOWs" antes de sincronizar — la app solo lee lo que está en esa carpeta, no toda la Bandeja de Entrada.

Si alguna fila no tiene fecha de inicio ni valor (normalmente porque el Word no llegó adjunto o no se pudo leer), la app avisa antes de guardar y pregunta si continuar igual con datos incompletos; esos datos se pueden completar manualmente en el Excel después.

---

## 7. Pestaña Pendientes

La pestaña 🟡 Pendientes muestra todos los SOWs que ya están registrados en el Excel pero todavía no tienen PDF firmado — los que están "en espera de firma".

- Doble clic sobre cualquier SOW pendiente abre un panel con el análisis del documento Word: fechas, valor, cargos, vigencia y validación del contrato (sección 8).
- Un SOW sale de Pendientes automáticamente cuando se guarda su PDF firmado.
- Si la lista está vacía, significa que todos los SOWs están al día.

> ℹ️ Los Pendientes se cargan desde el Excel maestro al abrir la app y se actualizan automáticamente después de cada sincronización.

---

## 8. Análisis de 6 puntos del documento

Para cada SOW con Word disponible (adjunto o ya archivado), el panel de detalle muestra un análisis automático de 6 puntos, algo que antes había que revisar leyendo el documento completo:

1. **Contrato marco** — referencia al acuerdo del 14 de abril de 2020 con NEORIS Colombia.
2. **Tipo de contrato** — confirma que es prestación de servicios.
3. **Cargos contratados** — lista de perfiles/roles del documento.
4. **Vigencia del servicio** — vigente / por vencer / vencida.
5. **Servicios pactados** — tabla "DT Original" del documento.
6. **Cláusula de impuestos** — verifica que esté correctamente incluida.

---

## 9. Solución de problemas frecuentes

| Situación | Qué significa / qué hacer |
|---|---|
| **"No hay correos para procesar"** | Verificar que los correos de CHUBB estén dentro de la carpeta "SOWs" (no en la Bandeja de Entrada), que Outlook esté completamente sincronizado, y reintentar sincronizar después de unos segundos. |
| **El SOW aparece en Novedades pero sin fechas ni valor** | El correo no incluyó el Word adjunto, o el nombre del archivo no coincide con el número de SOW en la tabla. Se puede guardar igual con datos incompletos y completarlos luego a mano en el Excel. |
| **Un SOW no aparece aunque el correo llegó ("Sin match")** | El nombre del archivo adjunto no coincide exactamente con el número de SOW de la tabla del correo — revisar ambos. |
| **La app no se conecta a Outlook** | Verificar que Outlook esté abierto y con sesión iniciada. Si acaba de abrirse, esperar unos segundos antes de sincronizar. Si persiste, cerrar y volver a abrir SOW Manager. |
| **Error al guardar en Excel** | La causa más común es tener `Listado SOW's.xlsm` **abierto en Excel** (por la misma persona u otra) al mismo tiempo — cerrarlo y reintentar. También verificar que la carpeta del proyecto siga sincronizada con SharePoint. |
| **No se encontró el archivo Word en la carpeta archivada** (en el panel de detalle) | El Word no llegó adjunto en el correo, o no se guardó correctamente — revisar la carpeta de esa persona/año. |
| **Windows bloquea el `.exe` al abrirlo (SmartScreen)** | El ejecutable no está firmado digitalmente. En el aviso, elegir "Más información" → "Ejecutar de todas formas". |

---

## 10. Buenas prácticas y qué NO hacer

- 🚫 **No abrir el Excel maestro** mientras SOW Manager está sincronizando o guardando datos — puede causar conflictos o errores de escritura.
- 🚫 **No mover ni renombrar** la carpeta "SOWs" en Outlook, ni la hoja `SOWS` o las columnas del Excel — la app las busca/ubica por ese nombre y posición exactos.
- 🚫 **No eliminar** la carpeta "SOWs" de Outlook. Si se elimina, la app la vuelve a crear vacía y hay que volver a mover los correos.
- 🚫 **No cerrar Outlook** mientras la app está sincronizando.
- 🚫 **No mover los correos fuera de "SOWs"** después de procesados — si se eliminan o mueven, la app puede volver a mostrarlos como novedades en la siguiente sincronización.
- 🚫 **No editar manualmente la fila 7** (encabezados) del Excel.
- Está bien guardar SOWs con datos incompletos si hace falta — pero conviene revisar primero si el Word realmente llegó adjunto antes de hacerlo.
