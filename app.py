"""App web para armar la planilla de marmoleo.

Los archivos subidos se procesan en memoria y no se guardan en el servidor.

Variables de entorno (opcionales):
  APP_PASSWORD  si se define, la app pide usuario y contraseña (HTTP Basic).
  APP_USER      usuario para ese acceso (por defecto: marmoleo).
"""
import base64
import hmac
import os
from datetime import date

from flask import Flask, Response, jsonify, render_template, request

from marmoleo_core import DatosInvalidos, generar

app = Flask(__name__, template_folder=".")
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB por solicitud

APP_USER = os.environ.get("APP_USER", "marmoleo")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")


@app.before_request
def pedir_acceso():
    if not APP_PASSWORD or request.path == "/healthz":
        return None
    auth = request.authorization
    if auth and hmac.compare_digest(auth.username or "", APP_USER) \
            and hmac.compare_digest(auth.password or "", APP_PASSWORD):
        return None
    return Response("Acceso restringido.", 401, {"WWW-Authenticate": 'Basic realm="Marmoleo"'})


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/healthz")
def healthz():
    return "ok"


def _error(mensaje: str, codigo: int = 400):
    return jsonify(ok=False, error=mensaje), codigo


@app.post("/api/generar")
def api_generar():
    listing = request.files.get("listing")
    teins = request.files.get("teins")
    if not listing or not listing.filename:
        return _error("Falta el archivo del Listing Faena.")
    if not teins or not teins.filename:
        return _error("Falta el archivo del Teins.")
    for f in (listing, teins):
        if not f.filename.lower().endswith(".xlsx"):
            return _error(f"«{f.filename}» no es un archivo .xlsx.")

    serie = (request.form.get("serie") or "").strip()
    if not serie:
        return _error("Ingrese el número de serie inicial.")

    fecha = None
    fecha_txt = (request.form.get("fecha") or "").strip()
    if fecha_txt:
        try:
            fecha = date.fromisoformat(fecha_txt)
        except ValueError:
            return _error("La fecha de la planilla no es válida.")

    try:
        res = generar(listing.stream, teins.stream, serie, fecha, listing.filename, teins.filename)
    except DatosInvalidos as exc:
        return _error(str(exc), 422)

    return jsonify(
        ok=True,
        archivo=res.nombre_archivo,
        contenido=base64.b64encode(res.contenido).decode("ascii"),
        fecha_faena=res.fecha_faena.strftime("%d/%m/%Y") if res.fecha_faena else None,
        fecha_planilla=res.fecha_planilla.strftime("%d/%m/%Y"),
        hojas=[
            dict(tropa=h.tropa, animales=h.animales, serie=h.serie,
                 garron_desde=h.garron_desde, garron_hasta=h.garron_hasta, excluidos=h.excluidos)
            for h in res.hojas
        ],
        advertencias=res.advertencias,
        faltantes=res.faltantes,
    )


@app.errorhandler(413)
def demasiado_grande(_):
    return _error("Los archivos superan el tamaño permitido (10 MB).", 413)


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
