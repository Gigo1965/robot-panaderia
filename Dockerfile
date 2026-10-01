# Receta de la "caja" (contenedor) del bot de WhatsApp para Render.
# Robot Panaderia - fase de hosting (1-oct-2026).
#
# Que hace, en orden:
#   1. Parte de un Linux pequeno con Python 3.12.
#   2. Instala LibreOffice SIN pantalla (solo Calc), que el robot usa para
#      recalcular las formulas del Excel antes de subirlo a Drive.
#   3. Instala las librerias de los dos requirements.txt (robot + bot) y gunicorn.
#   4. Copia el codigo del proyecto (lo que excluye .dockerignore NO entra:
#      .env, credenciales, copias del Excel...).
#   5. Arranca el bot con gunicorn: 1 solo proceso (las conversaciones viven en
#      memoria) y 4 hilos, en el puerto que Render indica en la variable PORT.
#
# Las claves NUNCA van aqui: en Render se cargan como variables de entorno y la
# clave JSON de Google como "Secret File" (/etc/secrets/...).

FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    SAL_USE_VCLPLUGIN=svp \
    TZ=America/Costa_Rica

# LibreOffice Calc sin interfaz grafica + una fuente basica
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc-nogui fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Primero solo las listas de librerias (asi Render reutiliza este paso si el codigo cambia)
COPY requirements.txt ./requirements.txt
COPY whatsapp_bot/requirements.txt ./whatsapp_bot/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt -r whatsapp_bot/requirements.txt "gunicorn>=22"

# Luego el codigo
COPY . .

# --timeout 120: en el plan gratis (0.1 CPU) leer el catalogo puede tardar; con el
# limite normal (30 s) gunicorn reiniciaria el proceso y se perderian las conversaciones.
CMD gunicorn --workers 1 --threads 4 --timeout 120 --bind 0.0.0.0:${PORT:-10000} whatsapp_bot.app:app
