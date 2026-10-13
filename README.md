# Planilla de marmoleo

App web que arma la **planilla de evaluación de carcasas (cuota Reglamento CE 481/2012)** a partir de dos archivos descargados del sistema de faena.

## Uso

1. Descargar el **Listing Faena** (`GD_informe_ListingFaena_*.xlsx`).
2. Descargar el **Twins** (`DESCARGA_TWINS_*.xlsx`).
3. Abrir la app, cargar ambos archivos, indicar el **primer número de serie** y generar.

Se descarga un Excel con una hoja por tropa. Los archivos se procesan en memoria y no se guardan en el servidor.

## Reglas de armado

- **Hojas:** una por cada tropa con amparo **W** en el Listing Faena, en el orden del listing (por garrón inicial). El nombre de la hoja es el número de tropa.
- **Filas:** solo los animales con amparo **W** en el Twins, ordenados por número de garrón. Se completan GARRÓN, CLASIFICACIÓN (`CLAS`) y TIPIFICACIÓN (`TIP`). Los animales de la misma tropa con otro amparo (EXPORTACIÓN, NO HILTON, CONSUMO) no se incluyen y se informan en pantalla.
- **Fecha:** día siguiente a la faena. Puede indicarse otra en el formulario.
- **N° de serie:** se ingresa el de la primera hoja y las siguientes continúan en forma correlativa (por ejemplo `00018335`, `00018336`, …). Si se ingresa sin ceros a la izquierda se completa a 8 dígitos.
- Las columnas de osificación, marmoleo, color, consistencia, textura y apto quedan en blanco para completar. La fila **TOTAL AUTORIZADAS** suma la columna SACRAS.

La app avisa en pantalla si encuentra diferencias entre los archivos: cantidad de animales distinta a la del Listing, garrones fuera del rango, garrones repetidos, animales sin clasificación o tipificación, fechas de faena que no coinciden, o animales W en tropas que el Listing no marca como W.

## Archivos

```
app.py                servidor Flask (formulario y /api/generar)
marmoleo_core.py      lectura de los archivos y armado de la planilla
index.html            interfaz
plantilla_base.xlsx   encabezado y formato de la planilla
logo.jpeg             logo de la planilla
make_template.py      regenera la plantilla desde una planilla ya armada
verificar.py          compara el resultado con una planilla de referencia
render.yaml           configuración de despliegue en Render
```

Todos los archivos están en el mismo nivel, sin carpetas.

## Ejecutar en la computadora

```bash
python -m venv .venv
source .venv/bin/activate        # en Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py                    # http://localhost:5000
```

## Publicar en GitHub y Render

1. Crear un repositorio en GitHub (preferentemente **privado**) y subir el contenido de esta carpeta.
2. En Render: **New → Blueprint**, conectar el repositorio. Render lee `render.yaml` y crea el servicio.
3. Opcional, recomendado: en el servicio, **Environment**, definir `APP_PASSWORD` para que la app pida usuario (`marmoleo`) y contraseña. Sin esa variable, cualquier persona con el enlace puede usarla.
4. Cada `git push` a la rama principal vuelve a desplegar la app.

En el plan gratuito de Render el servicio se suspende tras unos minutos sin uso y la primera visita puede demorar cerca de un minuto en responder.

## Cambios en el diseño de la planilla

Si cambia el formato de la planilla, abrir una planilla armada con el nuevo diseño y ejecutar:

```bash
python make_template.py planilla_nueva.xlsx
```

## Verificación contra una planilla de referencia

```bash
python verificar.py listing.xlsx twins.xlsx 00018335 planilla_de_referencia.xlsx
```
