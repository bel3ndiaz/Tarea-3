"""
Biblioteca Digital de Instrucciones de Trabajo - Impresos Múltiples
Fase 3: documentos PDF/Word guardados en MySQL con historial de versiones
"""
import csv
import io
import mysql.connector

# python-docx lee el texto de los documentos Word (.docx) para la vista previa.
# Si no está instalado, la página sigue funcionando y solo avisa.
try:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    docx = None
from datetime import date, datetime

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, abort, send_file, Response)

app = Flask(__name__)
app.secret_key = "clave-solo-para-mensajes-flash"


# =====================================================================
# Integrante 1: base de datos, configuración y archivar/reactivar
# =====================================================================


# ---------------------------------------------------------------------
# Reglas de los documentos (Fase 3)
# ---------------------------------------------------------------------
EXTENSIONES_PERMITIDAS = ["pdf", "doc", "docx"]
TAMANO_MAXIMO_MB = 50
TAMANO_MAXIMO = TAMANO_MAXIMO_MB * 1024 * 1024  # en bytes
# Límite duro de toda la petición: si alguien envía algo enorme, Flask lo corta
# antes de leerlo (error 413). Los archivos entre 50 y 64 MB llegan y se rechazan
# en la validación con un mensaje en el mismo formulario.
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024


AREAS = [
    "Impresión Offset",
    "Impresión Digital",
    "Preparación de Tintas",
    "Troquelado",
    "Acabados",
    "Encuadernación",
    "Mantenimiento",
    "Control de Calidad",
]


def conectar():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="biblioteca_impresos",
        # buffered: cada consulta se lee completa al ejecutarla, así se pueden
        # hacer varias consultas seguidas con el mismo cursor sin el error
        # "Unread result found" (por ejemplo: buscar la instrucción y luego
        # calcular MAX(version) antes de insertar el documento).
        buffered=True,
    )


@app.template_filter("fecha")
def formato_fecha(valor):
    """Muestra una fecha como 15/10/2026. MySQL entrega objetos de fecha (DATE)."""
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y")
    partes = str(valor).split("-")
    if len(partes) == 3:
        return partes[2] + "/" + partes[1] + "/" + partes[0]
    return valor


def limite_paquete_mysql(cursor):
    """Bytes que MySQL acepta en una sola consulta (max_allowed_packet).
    En XAMPP suele ser 1 MB; si el archivo no cabe, MySQL corta la conexión."""
    cursor.execute("SELECT @@max_allowed_packet AS limite")
    fila = cursor.fetchone()
    if isinstance(fila, dict):
        return int(fila["limite"])
    return int(fila[0])


@app.context_processor
def datos_para_plantillas():
    """Variables disponibles en todas las plantillas (por ejemplo, el pie de página)."""
    return {"tamano_maximo_mb": TAMANO_MAXIMO_MB}


@app.errorhandler(413)
def archivo_demasiado_grande(error):
    """La petición superó MAX_CONTENT_LENGTH: no se guarda nada."""
    flash("El archivo es demasiado grande. El máximo permitido es "
          + str(TAMANO_MAXIMO_MB) + " MB.", "error")
    return redirect(request.referrer or url_for("listar"))


@app.errorhandler(mysql.connector.Error)
def error_base_datos(error):
    """Si MySQL está apagado o falla una consulta, se muestra un mensaje claro."""
    return render_template("error.html", detalle=str(error)), 500


@app.route("/")
def inicio():
    return redirect(url_for("listar"))


def contar_archivadas():
    """Cuántas instrucciones están archivadas (para el enlace del listado y el resumen)."""
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    cursor.execute("SELECT COUNT(*) AS total FROM instrucciones WHERE activa = 0")
    total = cursor.fetchone()["total"]
    cursor.close()
    conexion.close()
    return total


def cambiar_estado_instruccion(id, activa):
    """Archiva (activa=0) o reactiva (activa=1) una instrucción.
    No borra nada: los documentos y el historial de versiones se conservan."""
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    cursor.execute("SELECT codigo, activa FROM instrucciones WHERE id = %s", (id,))
    fila = cursor.fetchone()
    if fila is None:
        cursor.close()
        conexion.close()
        abort(404)

    if activa == 0 and fila["activa"] == 0:
        flash(fila["codigo"] + " ya estaba archivada.", "error")
    elif activa == 1 and fila["activa"] == 1:
        flash(fila["codigo"] + " ya estaba activa.", "error")
    else:
        fecha_archivado = None
        if activa == 0:
            fecha_archivado = datetime.now().replace(microsecond=0)
        cursor.execute("UPDATE instrucciones SET activa = %s, fecha_archivado = %s WHERE id = %s",
                       (activa, fecha_archivado, id))
        conexion.commit()
        if activa == 0:
            flash("Instrucción " + fila["codigo"] + " archivada. Ya no aparece en el listado "
                  "ni en el resumen, pero sus documentos y su historial se conservan.")
        else:
            flash("Instrucción " + fila["codigo"] + " reactivada.")
    cursor.close()
    conexion.close()
    return redirect(url_for("detalle", id=id))


@app.route("/instrucciones/<int:id>/archivar", methods=["POST"])
def archivar(id):
    return cambiar_estado_instruccion(id, 0)


@app.route("/instrucciones/<int:id>/reactivar", methods=["POST"])
def reactivar(id):
    return cambiar_estado_instruccion(id, 1)


# =====================================================================
# Integrante 2: subida de documentos, nota de cambios y vista previa
# =====================================================================


# Primeros bytes ("firma") de cada formato. Sirven para detectar un archivo
# renombrado: por ejemplo, una imagen a la que solo le cambiaron la extensión a .pdf.
FIRMAS = {
    "pdf": b"%PDF",                 # PDF
    "doc": b"\xd0\xcf\x11\xe0",     # Word 97-2003 (formato OLE)
    "docx": b"PK\x03\x04",          # Word moderno (es un ZIP por dentro)
}


@app.template_filter("tamano")
def formato_tamano(num_bytes):
    """Muestra un tamaño en bytes como 850 KB o 1.2 MB."""
    if num_bytes < 1024 * 1024:
        return str(max(1, round(num_bytes / 1024))) + " KB"
    return str(round(num_bytes / (1024 * 1024), 1)) + " MB"


def obtener_extension(nombre):
    """'Manual.PDF' -> 'pdf'. Si no tiene punto devuelve ''."""
    if "." not in nombre:
        return ""
    return nombre.rsplit(".", 1)[1].lower()


def validar_archivo(archivo):
    """Revisa el archivo recibido en request.files.
    Devuelve (errores, datos). Si hay errores, datos es None y no se guarda nada."""
    if archivo is None or archivo.filename == "":
        return ["Debes elegir un documento PDF o Word."], None

    # Algunos navegadores envían la ruta completa: solo se conserva el nombre
    nombre = archivo.filename.replace("\\", "/").split("/")[-1].strip()
    extension = obtener_extension(nombre)
    if extension not in EXTENSIONES_PERMITIDAS:
        return ["Tipo de archivo no permitido (" + nombre + "). "
                "Solo se aceptan .pdf, .doc o .docx."], None
    if len(nombre) > 255:
        return ["El nombre del archivo no puede tener más de 255 caracteres."], None

    contenido = archivo.read()
    if len(contenido) == 0:
        return ["El archivo " + nombre + " está vacío (0 bytes)."], None
    if len(contenido) > TAMANO_MAXIMO:
        return ["El archivo " + nombre + " pesa " + formato_tamano(len(contenido))
                + " y el máximo permitido es " + str(TAMANO_MAXIMO_MB) + " MB."], None
    if not contenido.startswith(FIRMAS[extension]):
        return ["El contenido de " + nombre + " no corresponde a un archivo ."
                + extension + " válido."], None

    return [], {"nombre": nombre, "extension": extension, "contenido": contenido}


def validar_nota(nota, version):
    """Desde la versión 2 es obligatorio explicar qué cambió (máximo 300 caracteres)."""
    if version >= 2 and len(nota) < 5:
        return ["Describe qué cambió respecto a la versión " + str(version - 1)
                + " (al menos 5 caracteres)."]
    if len(nota) > 300:
        return ["La nota de cambios no puede tener más de 300 caracteres."]
    return []


def version_siguiente(cursor, instruccion_id):
    """La primera versión es 1; después, la mayor que exista + 1."""
    cursor.execute("SELECT MAX(version) AS mayor FROM archivos WHERE instruccion_id = %s",
                   (instruccion_id,))
    fila = cursor.fetchone()
    mayor = fila["mayor"] if isinstance(fila, dict) else fila[0]
    if mayor is None:
        return 1
    return mayor + 1


def guardar_archivo(cursor, instruccion_id, datos, enlace, nota=None):
    """Inserta una versión NUEVA del documento (nunca reemplaza la anterior).
    Devuelve el número de versión asignado. El commit lo hace quien la llama."""
    version = version_siguiente(cursor, instruccion_id)
    ahora = datetime.now().replace(microsecond=0)
    cursor.execute(
        "INSERT INTO archivos (instruccion_id, version, nombre_original, extension, "
        "tamano_bytes, contenido, enlace_video, nota_cambio, fecha_subida) "
        "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)",
        (instruccion_id, version, datos["nombre"], datos["extension"],
         len(datos["contenido"]), datos["contenido"], enlace, nota, ahora))
    # Historial de versiones: queda registro de qué se subió, por qué y cuándo
    accion = "Subió la versión " + str(version) + ": " + datos["nombre"]
    if nota:
        accion = accion + " · " + nota
    cursor.execute(
        "INSERT INTO historial_versiones (instruccion_id, version, accion, fecha) "
        "VALUES (%s, %s, %s, %s)",
        (instruccion_id, version, accion[:255], ahora))
    # El enlace del video y la fecha de la instrucción quedan al día
    cursor.execute(
        "UPDATE instrucciones SET enlace_video = %s, fecha_actualizacion = %s WHERE id = %s",
        (enlace, ahora.date(), instruccion_id))
    return version


def error_si_no_cabe(cursor, datos):
    """Mensaje de error si MySQL no aceptaría el archivo; lista vacía si cabe."""
    limite = limite_paquete_mysql(cursor)
    if datos is not None and len(datos["contenido"]) > limite - 64 * 1024:
        return ["MySQL solo acepta archivos de hasta " + formato_tamano(limite - 64 * 1024)
                + " (max_allowed_packet). Revisa el README para subir ese límite."]
    return []


@app.route("/instrucciones/<int:id>/subir", methods=["GET", "POST"])
def subir(id):
    """Sube una versión nueva del documento de una instrucción."""
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    cursor.execute("SELECT id, codigo, titulo, area, enlace_video, activa FROM instrucciones WHERE id = %s", (id,))
    fila = cursor.fetchone()
    if fila is None:
        cursor.close()
        conexion.close()
        abort(404)
    if fila["activa"] == 0:
        cursor.close()
        conexion.close()
        flash(fila["codigo"] + " está archivada: reactívala para subir documentos.", "error")
        return redirect(url_for("detalle", id=id))

    errores = []
    datos = {"enlace_video": fila["enlace_video"] or "", "nota_cambio": ""}
    proxima = version_siguiente(cursor, id)
    if request.method == "POST":
        enlace = request.form.get("enlace_video", "").strip()
        nota = request.form.get("nota_cambio", "").strip()
        datos = request.form
        errores_archivo, datos_archivo = validar_archivo(request.files.get("archivo"))
        errores = errores_archivo + validar_enlace(enlace) + validar_nota(nota, proxima)
        if not errores:
            errores = error_si_no_cabe(cursor, datos_archivo)
        if not errores:
            if enlace == "":
                enlace = None
            if nota == "":
                nota = None
            version = guardar_archivo(cursor, id, datos_archivo, enlace, nota)
            conexion.commit()
            cursor.close()
            conexion.close()
            flash("Se guardó la versión " + str(version) + " del documento de " + fila["codigo"]
                  + " (" + datos_archivo["nombre"] + ", "
                  + formato_tamano(len(datos_archivo["contenido"])) + ").")
            return redirect(url_for("detalle", id=id))

    cursor.close()
    conexion.close()
    return render_template("subir.html", i=fila, datos=datos, errores=errores, proxima=proxima,
                           tamano_maximo=TAMANO_MAXIMO_MB)


@app.route("/archivos/<int:id>/ver")
def ver_archivo(id):
    """Abre un PDF dentro del navegador. Word no se puede mostrar, se descarga."""
    return enviar_archivo(id, False)


def leer_docx(contenido):
    """Convierte un .docx en una lista de bloques para mostrarlos en HTML, en el
    mismo orden del documento: títulos, párrafos, listas y tablas.
    Ejemplo: [{"tipo": "titulo", "nivel": 1, "texto": "..."},
              {"tipo": "lista", "elementos": ["...", "..."]}, ...]"""
    documento = docx.Document(io.BytesIO(contenido))
    bloques = []
    # Se recorre el cuerpo del documento en orden (párrafos y tablas mezclados)
    for elemento in documento.element.body.iterchildren():
        if elemento.tag == qn("w:p"):
            parrafo = Paragraph(elemento, documento)
            texto = parrafo.text.strip()
            estilo = parrafo.style.name if parrafo.style is not None else ""
            if texto == "":
                continue
            if estilo.startswith("Heading") or estilo == "Title":
                nivel = 1
                if estilo[-1:].isdigit():
                    nivel = min(int(estilo[-1]), 3)
                bloques.append({"tipo": "titulo", "nivel": nivel, "texto": texto})
            elif "List" in estilo:
                # "List Number" es una lista numerada; "List Bullet", con viñetas.
                # Los elementos seguidos del mismo tipo se juntan en una sola lista.
                numerada = "Number" in estilo
                if bloques and bloques[-1]["tipo"] == "lista" and bloques[-1]["numerada"] == numerada:
                    bloques[-1]["elementos"].append(texto)
                else:
                    bloques.append({"tipo": "lista", "numerada": numerada, "elementos": [texto]})
            else:
                bloques.append({"tipo": "parrafo", "texto": texto})
        elif elemento.tag == qn("w:tbl"):
            tabla = Table(elemento, documento)
            filas = []
            for fila in tabla.rows:
                filas.append([celda.text.strip() for celda in fila.cells])
            if filas:
                bloques.append({"tipo": "tabla", "filas": filas})
    return bloques


@app.route("/archivos/<int:id>/vista")
def vista_previa(id):
    """Muestra una versión del documento dentro del sistema, con sus datos y
    botones para pasar a la versión anterior o siguiente."""
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    # Primero solo los datos (sin el contenido): basta para el PDF y el .doc
    cursor.execute(
        "SELECT a.id, a.instruccion_id, a.version, a.nombre_original, a.extension, "
        "a.tamano_bytes, a.nota_cambio, a.fecha_subida, i.codigo, i.titulo "
        "FROM archivos a JOIN instrucciones i ON i.id = a.instruccion_id WHERE a.id = %s", (id,))
    archivo = cursor.fetchone()
    if archivo is None:
        cursor.close()
        conexion.close()
        abort(404)

    # Todas las versiones de la instrucción, para la navegación anterior/siguiente
    cursor.execute("SELECT id, version FROM archivos WHERE instruccion_id = %s ORDER BY version",
                   (archivo["instruccion_id"],))
    versiones = cursor.fetchall()
    anterior = None
    siguiente = None
    for posicion in range(len(versiones)):
        if versiones[posicion]["id"] == id:
            if posicion > 0:
                anterior = versiones[posicion - 1]
            if posicion < len(versiones) - 1:
                siguiente = versiones[posicion + 1]
    estado = estado_archivo(archivo["version"], versiones[-1]["version"])

    # Solo el .docx necesita leer el contenido para sacar su texto
    bloques = None
    aviso = None
    if archivo["extension"] == "docx":
        if docx is None:
            aviso = "Para ver documentos Word instala python-docx (pip install -r requirements.txt)."
        else:
            cursor.execute("SELECT contenido FROM archivos WHERE id = %s", (id,))
            contenido = bytes(cursor.fetchone()["contenido"])
            try:
                bloques = leer_docx(contenido)
            except Exception:
                aviso = "No se pudo leer el contenido de este documento Word. Puedes descargarlo."
            if bloques == []:
                aviso = "El documento no tiene texto para mostrar."
    elif archivo["extension"] == "doc":
        aviso = ("La vista previa no está disponible para el formato Word 97-2003 (.doc). "
                 "Descárgalo para abrirlo.")
    cursor.close()
    conexion.close()

    return render_template("vista.html", a=archivo, estado=estado, anterior=anterior,
                           siguiente=siguiente, total_versiones=len(versiones),
                           bloques=bloques, aviso=aviso)


# =====================================================================
# Integrante 3: listado, versiones, detalle y búsqueda con filtros
# =====================================================================


@app.template_filter("fecha_hora")
def formato_fecha_hora(valor):
    """Muestra fecha y hora como 15/10/2026 14:05."""
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y %H:%M")
    texto = str(valor)
    return formato_fecha(texto[:10]) + " " + texto[11:16]


def estado_archivo(version, version_mas_reciente):
    """Regla del reto: la versión más reciente es la vigente; las demás, obsoletas."""
    if version == version_mas_reciente:
        return "vigente"
    elif version < version_mas_reciente:
        return "obsoleta"
    else:
        return "por_revisar"  # no debería pasar: sería mayor que la más reciente


# Columnas del listado. La versión actual y el nombre del documento vigente se
# calculan desde la tabla archivos SIN leer la columna contenido (LONGBLOB).
COLUMNAS_LISTADO = (
    "SELECT i.id, i.codigo, i.titulo, i.area, i.enlace_video, i.fecha_actualizacion, i.activa, "
    "(SELECT MAX(a.version) FROM archivos a WHERE a.instruccion_id = i.id) AS version_actual, "
    "(SELECT a.nombre_original FROM archivos a WHERE a.instruccion_id = i.id "
    " ORDER BY a.version DESC LIMIT 1) AS archivo_vigente "
    "FROM instrucciones i "
)


def consultar_instrucciones(texto, area, estado, activas=True):
    """Lista de instrucciones filtrada. Devuelve [{"datos": fila, "estado": ...}].
    - texto: busca en código, título, área y en el nombre de CUALQUIER versión del documento
    - area: "" para todas
    - estado: "" para todos, o una clave de NOMBRES_ESTADO
    - activas: True para las que están en uso, False para las archivadas"""
    # Por defecto solo se muestran las activas; las archivadas tienen su propia vista
    condiciones = ["i.activa = %s"]
    parametros = [1 if activas else 0]
    if texto != "":
        patron = "%" + texto + "%"
        condiciones.append(
            "(i.titulo LIKE %s OR i.codigo LIKE %s OR i.area LIKE %s OR EXISTS ("
            "SELECT 1 FROM archivos a WHERE a.instruccion_id = i.id AND a.nombre_original LIKE %s))")
        parametros = parametros + [patron, patron, patron, patron]
    if area != "":
        condiciones.append("i.area = %s")
        parametros.append(area)

    consulta = COLUMNAS_LISTADO
    if condiciones:
        consulta = consulta + "WHERE " + " AND ".join(condiciones) + " "
    consulta = consulta + "ORDER BY i.codigo"

    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    cursor.execute(consulta, tuple(parametros))
    filas = cursor.fetchall()
    cursor.close()
    conexion.close()

    instrucciones = []
    for fila in filas:
        estado_fila = clasificar_documentacion(fila["version_actual"] is not None,
                                               fila["enlace_video"] is not None)
        # El estado se calcula en Python, así que el filtro por estado se aplica aquí
        if estado == "" or estado == estado_fila:
            instrucciones.append({"datos": fila, "estado": estado_fila})
    return instrucciones


def leer_filtros():
    """Lee los filtros de la URL (?q=...&area=...&estado=...) y descarta valores inválidos."""
    texto = request.args.get("q", "").strip()
    area = request.args.get("area", "")
    estado = request.args.get("estado", "")
    if area not in AREAS:
        area = ""
    if estado not in NOMBRES_ESTADO:
        estado = ""
    return texto, area, estado


@app.route("/instrucciones")
def listar():
    texto, area, estado = leer_filtros()
    ver_archivadas = request.args.get("vista") == "archivadas"
    instrucciones = consultar_instrucciones(texto, area, estado, activas=not ver_archivadas)
    hay_filtros = texto != "" or area != "" or estado != ""
    return render_template("listado.html", instrucciones=instrucciones, q=texto,
                           area=area, estado=estado, hay_filtros=hay_filtros,
                           areas=AREAS, nombres_estado=NOMBRES_ESTADO,
                           ver_archivadas=ver_archivadas, total_archivadas=contar_archivadas())


@app.route("/instrucciones/<int:id>")
def detalle(id):
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    # LEFT JOIN: une la instrucción con todas sus versiones del documento.
    # Si todavía no tiene documentos, igual aparece (con columnas de archivo en NULL).
    # Solo se piden las columnas que se muestran: nunca a.contenido.
    cursor.execute(
        "SELECT i.id, i.codigo, i.titulo, i.area, i.enlace_video, i.fecha_actualizacion, "
        "i.activa, i.fecha_archivado, "
        "a.id AS archivo_id, a.version, a.nombre_original, a.extension, a.tamano_bytes, a.nota_cambio, "
        "a.enlace_video AS video_version, a.fecha_subida "
        "FROM instrucciones i "
        "LEFT JOIN archivos a ON a.instruccion_id = i.id "
        "WHERE i.id = %s ORDER BY a.version DESC", (id,))
    filas = cursor.fetchall()

    if len(filas) == 0:
        cursor.close()
        conexion.close()
        abort(404)

    # Historial de versiones (tabla protegida con ON DELETE RESTRICT)
    cursor.execute("SELECT version, accion, fecha FROM historial_versiones "
                   "WHERE instruccion_id = %s ORDER BY fecha DESC, id DESC", (id,))
    historial = cursor.fetchall()
    pasos = obtener_pasos(cursor, id)
    cursor.close()
    conexion.close()

    fila = filas[0]  # los datos de la instrucción se repiten en cada fila del JOIN
    versiones = []
    if fila["archivo_id"] is not None:
        mas_reciente = filas[0]["version"]  # vienen ordenadas de mayor a menor
        for f in filas:
            versiones.append({"datos": f, "estado": estado_archivo(f["version"], mas_reciente)})

    return render_template("detalle.html", i=fila, versiones=versiones, historial=historial,
                           pasos=pasos)


# =====================================================================
# Integrante 4: registrar, modificar, pasos, eliminar, descargas y restaurar
# =====================================================================


# Tipo MIME con el que el navegador recibe cada formato al descargar
TIPOS_MIME = {
    "pdf": "application/pdf",
    "doc": "application/msword",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


def validar_enlace(enlace):
    if enlace != "" and not (enlace.startswith("http://") or enlace.startswith("https://")):
        return ["El enlace de video debe comenzar con http:// o https:// (o dejarse vacío)."]
    return []


# ---------------------------------------------------------------------
# Pasos de la instrucción (contenido editable desde Registrar y Modificar)
# ---------------------------------------------------------------------
MAX_PASOS = 50


def leer_pasos_formulario():
    """Lee los pasos que vienen del formulario. Cada paso llega en dos listas
    paralelas (paso_descripcion y paso_control), en el orden de la pantalla.
    Las filas vacías se ignoran. Devuelve (pasos, errores)."""
    descripciones = request.form.getlist("paso_descripcion")
    controles = request.form.getlist("paso_control")
    pasos = []
    errores = []
    for posicion in range(len(descripciones)):
        descripcion = descripciones[posicion].strip()
        control = ""
        if posicion < len(controles):
            control = controles[posicion].strip()
        if descripcion == "" and control == "":
            continue  # fila en blanco: no cuenta como paso
        numero = len(pasos) + 1
        if descripcion == "":
            errores.append("El paso " + str(numero) + " tiene punto de control pero le falta la descripción.")
        elif len(descripcion) > 500:
            errores.append("La descripción del paso " + str(numero) + " no puede tener más de 500 caracteres.")
        if len(control) > 255:
            errores.append("El punto de control del paso " + str(numero) + " no puede tener más de 255 caracteres.")
        pasos.append({"descripcion": descripcion, "punto_control": control})
    if len(pasos) > MAX_PASOS:
        errores.append("Una instrucción puede tener como máximo " + str(MAX_PASOS) + " pasos.")
    return pasos, errores


def guardar_pasos(cursor, instruccion_id, pasos):
    """Reemplaza los pasos de la instrucción: borra los anteriores e inserta
    los nuevos numerados 1, 2, 3... en el orden recibido. El commit lo hace quien la llama."""
    cursor.execute("DELETE FROM pasos WHERE instruccion_id = %s", (instruccion_id,))
    numero = 1
    for paso in pasos:
        control = paso["punto_control"]
        if control == "":
            control = None  # sin punto de control se guarda NULL
        cursor.execute(
            "INSERT INTO pasos (instruccion_id, numero, descripcion, punto_control) "
            "VALUES (%s, %s, %s, %s)",
            (instruccion_id, numero, paso["descripcion"], control))
        numero = numero + 1


def obtener_pasos(cursor, instruccion_id):
    """Pasos de una instrucción en orden."""
    cursor.execute("SELECT numero, descripcion, punto_control FROM pasos "
                   "WHERE instruccion_id = %s ORDER BY numero", (instruccion_id,))
    return cursor.fetchall()


@app.route("/instrucciones/nuevo", methods=["GET", "POST"])
def nuevo():
    errores = []
    if request.method == "POST":
        codigo = request.form["codigo"].strip().upper()
        titulo = request.form["titulo"].strip()
        area = request.form["area"]
        enlace = request.form["enlace_video"].strip()
        archivo = request.files.get("archivo")

        if codigo == "":
            errores.append("El código es obligatorio.")
        elif len(codigo) > 20:
            errores.append("El código no puede tener más de 20 caracteres.")
        if titulo == "":
            errores.append("El título es obligatorio.")
        elif len(titulo) > 100:
            errores.append("El título no puede tener más de 100 caracteres.")
        if area not in AREAS:
            errores.append("Selecciona un área de la lista.")
        errores = errores + validar_enlace(enlace)

        # Pasos escritos en el formulario
        pasos, errores_pasos = leer_pasos_formulario()
        errores = errores + errores_pasos

        # El documento es opcional al registrar, pero si se elige uno se valida
        datos_archivo = None
        hay_archivo = archivo is not None and archivo.filename != ""
        if hay_archivo:
            errores_archivo, datos_archivo = validar_archivo(archivo)
            errores = errores + errores_archivo

        # Una instrucción necesita contenido: al menos un paso o un documento
        if len(pasos) == 0 and not hay_archivo:
            errores.append("Agrega al menos un paso o adjunta el documento de la instrucción.")

        if not errores:
            hoy = date.today()
            if enlace == "":
                enlace = None  # sin enlace se guarda NULL en la base de datos
            conexion = conectar()
            cursor = conexion.cursor(dictionary=True)
            errores = error_si_no_cabe(cursor, datos_archivo)
            version = None
            if not errores:
                try:
                    cursor.execute(
                        "INSERT INTO instrucciones (codigo, titulo, area, enlace_video, fecha_actualizacion) "
                        "VALUES (%s, %s, %s, %s, %s)",
                        (codigo, titulo, area, enlace, hoy))
                    nuevo_id = cursor.lastrowid  # id que MySQL le asignó a la instrucción
                    guardar_pasos(cursor, nuevo_id, pasos)
                    if datos_archivo is not None:
                        version = guardar_archivo(cursor, nuevo_id, datos_archivo, enlace)
                    conexion.commit()  # instrucción, pasos y documento se guardan juntos o nada
                except mysql.connector.Error as error:
                    # 1062: el código ya existe (columna UNIQUE). Otro error se deja pasar.
                    if error.errno != 1062:
                        raise
                    conexion.rollback()
                    errores.append("Ya existe una instrucción con el código " + codigo + ".")
            cursor.close()
            conexion.close()
            if not errores:
                mensaje = ("Instrucción " + codigo + " registrada correctamente el " + formato_fecha(hoy)
                           + " con " + str(len(pasos)) + " paso(s)")
                if version is not None:
                    mensaje = mensaje + " con el documento " + datos_archivo["nombre"] + " (versión 1)"
                flash(mensaje + ".")
                return redirect(url_for("listar"))

    # Si hubo un error, el formulario vuelve a mostrar los pasos que se escribieron
    pasos_formulario = []
    if request.method == "POST":
        pasos_formulario, _ = leer_pasos_formulario()
    return render_template("nuevo.html", errores=errores, datos=request.form, areas=AREAS,
                           tamano_maximo=TAMANO_MAXIMO_MB, pasos=pasos_formulario)


@app.route("/instrucciones/editar/<int:id>", methods=["GET", "POST"])
def editar(id):
    """Modifica los datos y los pasos de la instrucción. La versión del documento
    no se escribe a mano: cambia sola cada vez que se sube un documento nuevo."""
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    cursor.execute("SELECT id, codigo, titulo, area, enlace_video, fecha_actualizacion "
                   "FROM instrucciones WHERE id = %s", (id,))
    fila = cursor.fetchone()
    if fila is None:
        cursor.close()
        conexion.close()
        abort(404)

    errores = []
    datos = fila  # valores que se muestran en el formulario
    pasos = obtener_pasos(cursor, id)
    if request.method == "POST":
        titulo = request.form["titulo"].strip()
        area = request.form["area"]
        enlace = request.form["enlace_video"].strip()
        datos = request.form
        pasos, errores_pasos = leer_pasos_formulario()

        if titulo == "":
            errores.append("El título es obligatorio.")
        elif len(titulo) > 100:
            errores.append("El título no puede tener más de 100 caracteres.")
        if area not in AREAS:
            errores.append("Selecciona un área de la lista.")
        errores = errores + validar_enlace(enlace) + errores_pasos

        # Sin pasos solo se permite si la instrucción ya tiene un documento subido
        if len(pasos) == 0:
            cursor.execute("SELECT COUNT(*) AS total FROM archivos WHERE instruccion_id = %s", (id,))
            if cursor.fetchone()["total"] == 0:
                errores.append("Agrega al menos un paso: esta instrucción todavía no tiene documento.")

        if not errores:
            hoy = date.today()
            if enlace == "":
                enlace = None
            cursor.execute(
                "UPDATE instrucciones SET titulo = %s, area = %s, enlace_video = %s, "
                "fecha_actualizacion = %s WHERE id = %s",
                (titulo, area, enlace, hoy, id))
            guardar_pasos(cursor, id, pasos)
            conexion.commit()  # datos y pasos se guardan juntos
            cursor.close()
            conexion.close()
            flash("Instrucción " + fila["codigo"] + " actualizada correctamente el "
                  + formato_fecha(hoy) + " (" + str(len(pasos)) + " paso(s)).")
            return redirect(url_for("detalle", id=id))

    cursor.close()
    conexion.close()
    return render_template("editar.html", i=fila, datos=datos, errores=errores, areas=AREAS,
                           pasos=pasos)


@app.route("/instrucciones/eliminar/<int:id>", methods=["POST"])
def eliminar(id):
    """Solo se puede eliminar una instrucción que todavía no tiene documentos.
    Las versiones están protegidas por ON DELETE RESTRICT: nunca se borran
    al eliminar un registro."""
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    cursor.execute("SELECT codigo FROM instrucciones WHERE id = %s", (id,))
    fila = cursor.fetchone()
    if fila is None:
        cursor.close()
        conexion.close()
        abort(404)

    # Se revisa ANTES de borrar, para dar un mensaje claro
    cursor.execute(
        "SELECT COUNT(*) AS total FROM archivos WHERE instruccion_id = %s", (id,))
    total_archivos = cursor.fetchone()["total"]

    if total_archivos > 0:
        flash("No se puede eliminar " + fila["codigo"] + ": tiene " + str(total_archivos)
              + " versión(es) del documento. El historial de versiones está protegido "
              "(ON DELETE RESTRICT). Si ya no se usa, puedes archivarla desde su detalle.",
              "error")
    else:
        try:
            cursor.execute("DELETE FROM instrucciones WHERE id = %s", (id,))
            conexion.commit()
            flash("Instrucción " + fila["codigo"] + " eliminada.")
        except mysql.connector.Error as error:
            # Segunda protección: si aun así tiene filas relacionadas, MySQL lo
            # rechaza con el error 1451 (o 1217 en versiones viejas)
            if error.errno not in (1451, 1217):
                raise
            conexion.rollback()
            flash("No se puede eliminar " + fila["codigo"]
                  + ": tiene historial relacionado (ON DELETE RESTRICT).", "error")
    cursor.close()
    conexion.close()
    return redirect(url_for("listar"))


def enviar_archivo(id, como_adjunto):
    """Busca el documento por su id y lo devuelve con su nombre original.
    Es la ÚNICA consulta que lee la columna contenido."""
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    cursor.execute("SELECT nombre_original, extension, contenido FROM archivos WHERE id = %s", (id,))
    fila = cursor.fetchone()
    cursor.close()
    conexion.close()
    if fila is None:
        abort(404)
    return send_file(io.BytesIO(fila["contenido"]),
                     mimetype=TIPOS_MIME[fila["extension"]],
                     as_attachment=como_adjunto,
                     download_name=fila["nombre_original"])


@app.route("/archivos/<int:id>/descargar")
def descargar(id):
    return enviar_archivo(id, True)


@app.route("/archivos/<int:id>/restaurar", methods=["POST"])
def restaurar(id):
    """Vuelve a poner en uso una versión obsoleta copiándola como versión NUEVA.
    No borra ni modifica ninguna versión: el historial queda completo."""
    conexion = conectar()
    cursor = conexion.cursor(dictionary=True)
    cursor.execute(
        "SELECT a.instruccion_id, a.version, a.nombre_original, a.extension, a.contenido, "
        "i.codigo, i.enlace_video, i.activa FROM archivos a "
        "JOIN instrucciones i ON i.id = a.instruccion_id WHERE a.id = %s", (id,))
    fila = cursor.fetchone()
    if fila is None:
        cursor.close()
        conexion.close()
        abort(404)

    instruccion_id = fila["instruccion_id"]
    mas_reciente = version_siguiente(cursor, instruccion_id) - 1
    if fila["activa"] == 0:
        flash(fila["codigo"] + " está archivada: reactívala para restaurar versiones.", "error")
    elif estado_archivo(fila["version"], mas_reciente) == "vigente":
        flash("La versión " + str(fila["version"]) + " ya es la vigente; no hace falta restaurarla.",
              "error")
    else:
        datos = {"nombre": fila["nombre_original"], "extension": fila["extension"],
                 "contenido": bytes(fila["contenido"])}
        nueva = guardar_archivo(cursor, instruccion_id, datos, fila["enlace_video"],
                                "Restaurada desde la versión " + str(fila["version"]))
        conexion.commit()
        flash("Se restauró la versión " + str(fila["version"]) + " de " + fila["codigo"]
              + " como la nueva versión vigente " + str(nueva) + ".")
    cursor.close()
    conexion.close()
    return redirect(url_for("detalle", id=instruccion_id))


@app.errorhandler(404)
def no_encontrado(error):
    return render_template("no_encontrado.html"), 404


if __name__ == "__main__":
    app.run(debug=True)
