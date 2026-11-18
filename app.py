"""
Biblioteca Digital de Instrucciones de Trabajo - Impresos Múltiples
Fase 3: documentos PDF/Word guardados en MySQL con historial de versiones
"""
import csv
import io
import mysql.connector

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



EXTENSIONES_PERMITIDAS = ["pdf", "doc", "docx"]
TAMANO_MAXIMO_MB = 50
TAMANO_MAXIMO = TAMANO_MAXIMO_MB * 1024 * 1024  
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


if __name__ == "__main__":
    app.run(debug=True)
FIRMAS = {
    "pdf": b"%PDF",                 
    "doc": b"\xd0\xcf\x11\xe0",     
    "docx": b"PK\x03\x04",
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
  
    accion = "Subió la versión " + str(version) + ": " + datos["nombre"]
    if nota:
        accion = accion + " · " + nota
    cursor.execute(
        "INSERT INTO historial_versiones (instruccion_id, version, accion, fecha) "
        "VALUES (%s, %s, %s, %s)",
        (instruccion_id, version, accion[:255], ahora))
   
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

    cursor.execute(
        "SELECT a.id, a.instruccion_id, a.version, a.nombre_original, a.extension, "
        "a.tamano_bytes, a.nota_cambio, a.fecha_subida, i.codigo, i.titulo "
        "FROM archivos a JOIN instrucciones i ON i.id = a.instruccion_id WHERE a.id = %s", (id,))
    archivo = cursor.fetchone()
    if archivo is None:
        cursor.close()
        conexion.close()
        abort(404)

 
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


if __name__ == "__main__":
    app.run(debug=True)
