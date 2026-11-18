# Biblioteca Digital de Instrucciones de Trabajo: Fase 3

Reto del Socio Formador **Impresos Múltiples**. Aplicación web hecha con Python (Flask) y MySQL (XAMPP + phpMyAdmin).

## Qué hace el sistema
Centraliza las instrucciones de trabajo de la planta, como "Preparación de prensa offset" o "Ajuste de troquel", en un solo lugar. Cada instrucción tiene:

- sus **pasos**, escritos directamente en la página, con un punto de control opcional en cada uno;
- su **documento** PDF o Word, con todas sus versiones guardadas en la base de datos;
- un **enlace a su video** (opcional).

Funciones principales:

- **Registrar, consultar, modificar y eliminar** instrucciones.
- **Subir documentos** PDF/Word (hasta 50 MB). Cada subida crea una versión nueva: la más reciente es **Vigente** y las anteriores quedan como **Obsoleta**. Nunca se reemplaza ni se borra una versión.
- **Descargar** cualquier versión y verla con **vista previa** en el navegador sin descargarla.
- **Nota de cambios**: desde la versión 2 es obligatorio explicar qué cambió.
- **Restaurar** una versión anterior: se copia como versión nueva y el historial queda completo.
- **Buscar** por código, título, área o nombre del documento, y **filtrar** por área y estado.
- **Archivar** las instrucciones que ya no se usan, en lugar de borrarlas, y **reactivarlas** después.
- **Resumen** de cuántas instrucciones tienen documento y video, con el avance por área.
- **Exportar** el listado a CSV para abrirlo en Excel.
- **Imprimir** la hoja de pasos de una instrucción.

## Cómo ejecutarlo
1. En XAMPP, iniciar **Apache** y **MySQL**.
2. (Una sola vez) Para subir documentos grandes, abrir `xampp/mysql/bin/my.ini` y, en la sección `[mysqld]`, poner `max_allowed_packet=64M`. Luego reiniciar MySQL.
3. En phpMyAdmin → **Importar** → elegir `crear_bd.sql` → **Continuar**. Esto crea la base `biblioteca_impresos` con 8 instrucciones de ejemplo y sus pasos.
4. Desde esta carpeta:

```
pip install -r requirements.txt
python app.py
```

5. Abrir http://127.0.0.1:5000

> Si un compañero cambia `crear_bd.sql`, hay que volver a importarlo. Al importarlo, los datos se reinician.

## Estructura de archivos
```
Fase3/
|-- app.py                  # rutas de Flask, validaciones y lógica (dividido por integrante)
|-- crear_bd.sql            # crea la base de datos, las 4 tablas y los datos de ejemplo
|-- requirements.txt        # flask, mysql-connector-python, python-docx
|-- .gitignore
|-- README.md
|-- templates/
|   |-- base.html           # encabezado, menú, mensajes y pie (lo heredan todas las páginas)
|   |-- listado.html        # tabla de instrucciones, búsqueda, filtros y exportar
|   |-- detalle.html        # datos, pasos, versiones del documento e historial
|   |-- nuevo.html          # registrar una instrucción
|   |-- editar.html         # modificar datos y pasos
|   |-- _editor_pasos.html  # editor de pasos que usan nuevo.html y editar.html
|   |-- subir.html          # subir una versión nueva del documento
|   |-- vista.html          # vista previa de un documento
|   |-- resumen.html        # contadores por documento y video, avance por área
|   |-- no_encontrado.html  # página 404
|   |-- error.html          # error de la base de datos
|-- static/
|   |-- style.css           # estilos con los colores de Impresos Múltiples
|   |-- pasos.js            # agregar, quitar y reordenar pasos
|   |-- img/                # logos
```

## Base de datos (MySQL)
| Tabla | Qué guarda |
|---|---|
| `instrucciones` | código (único), título, área, enlace de video, fecha de actualización, si está activa o archivada y desde cuándo |
| `pasos` | número, descripción y punto de control de cada paso |
| `archivos` | cada versión del documento: número de versión, nombre, extensión, tamaño, contenido (`LONGBLOB`), enlace de video y nota de cambios |
| `historial_versiones` | cada acción sobre los documentos, con fecha y hora |

**Decisión del equipo: el historial es restrictivo.** Las llaves foráneas de `archivos` y `historial_versiones` usan `ON DELETE RESTRICT`. MySQL no deja borrar una instrucción que tenga documentos, así que ninguna versión se pierde. Para retirarla se usa **Archivar**. Los pasos sí usan `ON DELETE CASCADE`, porque solo existen dentro de su instrucción.

La versión no se escribe a mano: se calcula con `SELECT MAX(version)` en la tabla `archivos`, y cada subida guarda `MAX + 1`. La restricción `UNIQUE (instruccion_id, version)` impide versiones repetidas.

## Rutas
| Ruta | Método | Qué hace |
|---|---|---|
| `/` | GET | Redirige al listado |
| `/instrucciones` | GET | Listado con búsqueda (`?q=`), filtros (`?area=`, `?estado=`) y archivadas (`?vista=archivadas`) |
| `/instrucciones/exportar.csv` | GET | Descarga el listado filtrado en CSV |
| `/instrucciones/<id>` | GET | Detalle: datos, pasos, versiones e historial |
| `/instrucciones/nuevo` | GET, POST | Registrar con pasos y, si se quiere, el primer documento |
| `/instrucciones/editar/<id>` | GET, POST | Modificar datos y pasos |
| `/instrucciones/eliminar/<id>` | POST | Eliminar, solo si no tiene documentos |
| `/instrucciones/<id>/subir` | GET, POST | Subir una versión nueva del documento |
| `/instrucciones/<id>/archivar` | POST | Archivar |
| `/instrucciones/<id>/reactivar` | POST | Reactivar |
| `/archivos/<id>/descargar` | GET | Descargar una versión (404 si no existe) |
| `/archivos/<id>/vista` | GET | Vista previa (PDF en el navegador, Word como texto) |
| `/archivos/<id>/ver` | GET | Abre el PDF dentro de la vista previa |
| `/archivos/<id>/restaurar` | POST | Restaurar una versión obsoleta como versión nueva |
| `/resumen` | GET | Resumen por documento y video |

## Validaciones
- **Documentos:** solo `.pdf`, `.doc` y `.docx`, sin importar mayúsculas. No pueden estar vacíos ni pasar de 50 MB. Se revisan los primeros bytes del archivo, así que se rechaza un archivo renombrado (por ejemplo, una imagen a la que le pusieron `.pdf`).
- **Nota de cambios:** obligatoria desde la versión 2, con un máximo de 300 caracteres.
- **Registrar y modificar:** código y título obligatorios, y el código no puede repetirse. El área debe ser de la lista. El enlace de video es opcional; si se escribe, debe empezar con `http://` o `https://`. Se necesita al menos un paso o un documento.
- **Pasos:** cada uno tiene hasta 500 caracteres de descripción y 255 de punto de control, con un máximo de 50 pasos.
- Si hay un error, no se guarda nada y el formulario conserva lo escrito.

## Cómo se decide cada estado
- **Versión** (`estado_archivo`, con `if / elif / else`): la versión más reciente es **Vigente** y las anteriores son **Obsoleta**.
- **Instrucción en el resumen** (`clasificar_documentacion`, con `if / elif / else`): **Completa** (tiene documento y video), **Falta video**, **Falta documento** o **Falta documento y video**. El resumen recorre todas las instrucciones activas con un `for` y cuenta cuántas hay en cada grupo y por área.

## Equipo: parte del código de cada integrante
> Reemplazar con los nombres y usuarios de GitHub reales.

En `app.py`, cada parte está marcada con un encabezado `# Integrante N`. `style.css` también se armó por partes: cada integrante agregó las secciones de sus pantallas.

| Integrante | Parte | Archivos |
|---|---|---|
| **Integrante 1** (@jdanielgregg) | Base de datos con historial restrictivo, configuración, mensajes de error, diseño general con la marca y **archivar / reactivar** | `crear_bd.sql`, `app.py` (bloque 1), `base.html`, `error.html`, `style.css` (estilos generales), `requirements.txt`, `.gitignore`, logos |
| **Integrante 2** (@bel3ndiaz) | Subida de documentos con validaciones y versiones, **nota de cambios** y **vista previa** | `app.py` (bloque 2), `subir.html`, `vista.html`, `style.css` (vista previa) |
| **Integrante 3** (@rafaelavalos1101) | Listado, detalle con versiones Vigente/Obsoleta e historial, **búsqueda y filtros** | `app.py` (bloque 3), `listado.html`, `detalle.html`, `style.css` (búsqueda y filtros) |
| **Integrante 4** (@A-858) | Registrar y modificar con **pasos editables**, eliminar restrictivo, descargas, 404 y **restaurar versiones** | `app.py` (bloque 4), `nuevo.html`, `editar.html`, `_editor_pasos.html`, `no_encontrado.html`, `pasos.js`, `style.css` (pasos e impresión) |
| **Integrante 5** (@juanmi-posada) | Resumen por documento y video, **exportar a CSV**, etiquetas de estado y este README | `app.py` (bloque 5), `resumen.html`, `style.css` (resumen y etiquetas), `README.md` |
