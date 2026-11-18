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


if __name__ == "__main__":
    app.run(debug=True)
