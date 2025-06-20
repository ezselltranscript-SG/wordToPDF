FROM python:3.9-slim

# Instalar LibreOffice y dependencias necesarias
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    fonts-liberation \
    procps \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*
    
# Verificar la instalación de LibreOffice
RUN libreoffice --version || echo "LibreOffice no está instalado correctamente"

# Establecer directorio de trabajo
WORKDIR /app

# Copiar archivos de requisitos primero para aprovechar la caché de Docker
COPY requirements.txt .

# Instalar dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el resto de la aplicación
COPY . .

# Crear directorios para archivos temporales
RUN mkdir -p uploads outputs

# Exponer el puerto que usa la aplicación (Render asignará el puerto a través de la variable PORT)
EXPOSE 10000

# Comando para ejecutar la aplicación usando la variable PORT de Render
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000}
