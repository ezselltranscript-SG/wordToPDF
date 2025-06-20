from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Depends
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import shutil
from pathlib import Path
import logging
import requests
import time
from typing import Optional

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuración para la API de Cloudmersive
CLOUDMERSIVE_API_KEY = os.environ.get('CLOUDMERSIVE_API_KEY', '')
# Si no hay API key en las variables de entorno, se usará un enfoque alternativo

app = FastAPI(
    title="Word to PDF Converter API",
    description="API sencilla para convertir documentos Word a PDF",
    version="1.0.0"
)

# Configurar CORS para permitir solicitudes desde cualquier origen
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Permite todos los orígenes
    allow_credentials=True,
    allow_methods=["*"],  # Permite todos los métodos
    allow_headers=["*"],  # Permite todos los headers
)


async def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Convierte un documento Word (.docx) a PDF usando la API de Cloudmersive.
    Esta implementación intenta preservar el formato original del documento Word.
    
    Args:
        docx_path: Ruta al archivo Word
        pdf_path: Ruta donde se guardará el PDF
    """
    try:
        # Usar la API de Cloudmersive para la conversión
        logger.info("Usando API de Cloudmersive para la conversión")
        return await convert_with_cloudmersive(docx_path, pdf_path)
    except Exception as e:
        logger.error(f"Error al convertir el documento: {str(e)}")
        return False

async def convert_with_cloudmersive(docx_path, pdf_path):
    """
    Convierte un documento Word a PDF usando la API de Cloudmersive o una API alternativa.
    """
    try:
        # Verificar si tenemos una clave API de Cloudmersive
        if CLOUDMERSIVE_API_KEY:
            # Configurar la API de Cloudmersive
            url = "https://api.cloudmersive.com/convert/docx/to/pdf"
            headers = {"Apikey": CLOUDMERSIVE_API_KEY}
            
            # Preparar el archivo para enviar
            with open(docx_path, 'rb') as file:
                files = {'inputFile': file}
                
                # Hacer la solicitud a la API
                logger.info("Enviando solicitud a la API de Cloudmersive")
                response = requests.post(url, headers=headers, files=files)
                
                # Verificar si la solicitud fue exitosa
                if response.status_code == 200:
                    # Guardar el PDF recibido
                    with open(pdf_path, 'wb') as output_file:
                        output_file.write(response.content)
                    logger.info(f"PDF creado exitosamente con Cloudmersive en: {pdf_path}")
                    return True
                else:
                    logger.error(f"Error en la API de Cloudmersive: {response.status_code} - {response.text}")
                    # Si hay error, intentar con la API alternativa
                    return await convert_with_alternative_api(docx_path, pdf_path)
        else:
            # Si no hay clave API, usar la API alternativa
            logger.info("No se encontró clave API de Cloudmersive, usando API alternativa")
            return await convert_with_alternative_api(docx_path, pdf_path)
    except Exception as e:
        logger.error(f"Error al usar la API de Cloudmersive: {str(e)}")
        # Si hay error, intentar con la API alternativa
        return await convert_with_alternative_api(docx_path, pdf_path)

async def convert_with_alternative_api(docx_path, pdf_path):
    """
    Convierte un documento Word a PDF usando una API alternativa gratuita.
    Esta es una implementación de respaldo que usa servicios públicos.
    """
    try:
        # Usar la API de ConvertAPI (tienen un plan gratuito limitado)
        url = "https://v2.convertapi.com/convert/docx/to/pdf"
        
        # Preparar el archivo para enviar
        with open(docx_path, 'rb') as file:
            files = {'File': file}
            params = {'Secret': 'free'}
            
            # Hacer la solicitud a la API
            logger.info("Enviando solicitud a la API alternativa")
            response = requests.post(url, files=files, params=params)
            
            # Verificar si la solicitud fue exitosa
            if response.status_code == 200:
                # Guardar el PDF recibido
                with open(pdf_path, 'wb') as output_file:
                    output_file.write(response.content)
                logger.info(f"PDF creado exitosamente con API alternativa en: {pdf_path}")
                return True
            else:
                logger.error(f"Error en la API alternativa: {response.status_code} - {response.text}")
                return False
    except Exception as e:
        logger.error(f"Error al usar la API alternativa: {str(e)}")
        return False

# Crear directorios para almacenar archivos temporales
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


@app.post("/convert/", summary="Convertir documento Word a PDF")
async def convert_word_to_pdf(file: UploadFile = File(...), background_tasks: BackgroundTasks = None):
    """
    Convierte un documento Word (.docx) a formato PDF.
    
    - **file**: Archivo Word (.docx) a convertir
    
    Retorna el archivo PDF convertido.
    """
    # Verificar que el archivo sea un documento Word
    if not file.filename.endswith(('.docx', '.doc')):
        logger.warning(f"Intento de convertir un archivo no válido: {file.filename}")
        raise HTTPException(status_code=400, detail="El archivo debe ser un documento Word (.docx o .doc)")
    
    # Generar nombres únicos para los archivos
    file_id = str(uuid.uuid4())
    input_filename = f"{file_id}_{file.filename}"
    input_path = UPLOAD_DIR / input_filename
    
    output_filename = f"{file_id}_{Path(file.filename).stem}.pdf"
    output_path = OUTPUT_DIR / output_filename
    
    logger.info(f"Iniciando conversión de {file.filename} a PDF")
    
    try:
        # Guardar el archivo subido
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        logger.info(f"Archivo guardado en {input_path}")
        
        # Convertir el documento Word a PDF
        conversion_result = await convert_docx_to_pdf(input_path, output_path)
        
        if not conversion_result or not output_path.exists():
            logger.error(f"Error al convertir {input_path} a PDF")
            raise HTTPException(status_code=500, detail="Error al convertir el documento")
        
        logger.info(f"Conversión exitosa: {output_path}")
        
        # Función para eliminar archivos temporales después de enviar la respuesta
        def cleanup_temp_files():
            try:
                if input_path.exists():
                    os.remove(input_path)
                    logger.info(f"Archivo temporal eliminado: {input_path}")
                # El archivo PDF se eliminará después de enviarse al cliente
                if output_path.exists():
                    os.remove(output_path)
                    logger.info(f"Archivo PDF temporal eliminado: {output_path}")
            except Exception as e:
                logger.error(f"Error al limpiar archivos temporales: {str(e)}")
        
        # Programar la limpieza de archivos temporales en segundo plano
        if background_tasks:
            background_tasks.add_task(cleanup_temp_files)
        
        # Devolver el archivo PDF
        return FileResponse(
            path=output_path,
            filename=f"{Path(file.filename).stem}.pdf",
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename={Path(file.filename).stem}.pdf"
            }
        )
    
    except Exception as e:
        # Manejar cualquier error durante la conversión
        logger.error(f"Error en la conversión: {str(e)}")
        
        # Limpiar archivos temporales en caso de error
        if input_path.exists():
            try:
                os.remove(input_path)
                logger.info(f"Archivo temporal eliminado después de error: {input_path}")
            except Exception as cleanup_error:
                logger.error(f"Error al limpiar archivo temporal: {str(cleanup_error)}")
        
        raise HTTPException(status_code=500, detail=f"Error en la conversión: {str(e)}")


@app.get("/", summary="Información de la API")
async def root():
    """
    Retorna información básica sobre la API.
    """
    return {
        "mensaje": "API de conversión de Word a PDF",
        "descripcion": "Esta API permite convertir documentos Word (.docx, .doc) a formato PDF",
        "uso": "Envía un archivo Word mediante POST a /convert/",
        "documentacion": "/docs",
        "version": "1.0.0"
    }


@app.get("/health", summary="Verificación de estado")
async def health_check():
    """
    Endpoint para verificar el estado del servicio.
    Util para monitoreo y health checks en servicios cloud.
    """
    return {"status": "ok", "message": "El servicio está funcionando correctamente"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8080, reload=True)
