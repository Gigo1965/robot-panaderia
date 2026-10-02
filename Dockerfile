# Receta de la "caja" (contenedor) del bot de WhatsApp para Render.
# Robot Panaderia - fase de hosting (1-oct-2026).
#
# || V15 (1-oct-2026, noche): la 1a version partia de Debian 12 (python:3.12-slim-bookworm)
# con LibreOffice 7.4 sin pantalla. En Render ese LibreOffice se caia al recalcular
# ("Fatal exception: Signal 6" en SvpSalInstance) y ningun pedido se podia registrar.
# Ahora la caja parte de Ubuntu 24.04 con LibreOffice 24.2 + Python 3.12: el MISMO
# entorno donde el recalculo se probo con el Excel v107 (3.4 s, 0 errores nuevos).
#
# Que hace, en orden:
#   1. Parte de Ubuntu 24.04.
#   2. Instala Python 3.12 y LibreOffice Calc (sin recomendados), que el robot usa
#      para recalcular las formulas del Excel antes de subirlo a Drive.
#   3. Crea un entorno de Python aparte (venv) e instala ahi las librerias de los dos
#      requirements.txt (robot + bot) y gunicorn.
#   4. Copia el codigo del proyecto (lo que excluye .dockerignore NO entra:
#      .env, credenciales, copias del Excel...).
#   5. Arranca el bot con gunicorn: 1 solo proceso (las conversaciones viven en
#      memoria) y 4 hilos, en el puerto que Render indica en la variable PORT.
#
# Las claves NUNCA van aqui: en Render se cargan como variables de entorno y la
# clave JSON de Google como "Secret File" (/etc/secrets/...).

FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    TZ=America/Costa_Rica

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        python3 python3-venv \
        libreoffice-calc libreoffice-core \
        fonts-dejavu-core tzdata ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Entorno de Python propio (Ubuntu no deja instalar librerias en el Python del sistema)
RUN python3 -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

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
