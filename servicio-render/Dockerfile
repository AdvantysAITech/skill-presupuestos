# Servicio de render de presupuestos.
# Se construye desde la raíz del repo (contexto ".."), porque copia la skill tal cual:
# el render usa exactamente el mismo código que genera el presupuesto en Claude.
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    TZ=Europe/Madrid \
    SKILL_DIR=/srv/skill \
    DATOS_DIR=/datos

RUN apt-get update \
    && apt-get install -y --no-install-recommends tzdata \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /srv
COPY servicio-render/requirements.txt /srv/requirements.txt
RUN pip install --no-cache-dir -r /srv/requirements.txt

COPY presupuestos-scala-vertical/ /srv/skill/
COPY servicio-render/app/ /srv/app/

RUN useradd --system --uid 10001 render \
    && mkdir -p /datos \
    && chown render /datos
USER render

EXPOSE 8000
HEALTHCHECK --interval=60s --timeout=10s --start-period=20s --retries=3 \
    CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8000/salud', timeout=8).status == 200 else 1)"

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1", "--proxy-headers", "--forwarded-allow-ips", "*"]
