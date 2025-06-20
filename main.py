from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import shutil
from pathlib import Path
import subprocess
import sys
import platform
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Importar docx2pdf para la conversión
from docx2pdf import convert

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
    Convierte un documento Word (.docx) a PDF usando docx2pdf.
    Esta biblioteca preserva el formato original del documento Word, incluyendo saltos de página.
    
    Args:
        docx_path: Ruta al archivo Word
        pdf_path: Ruta donde se guardará el PDF
    """
    try:
        # Usar docx2pdf para convertir el documento Word a PDF
        # Esta biblioteca preserva el formato original, incluyendo saltos de página
        convert(str(docx_path), str(pdf_path))
        
        # Verificar que el archivo PDF fue creado
        if os.path.exists(pdf_path):
            print(f"PDF creado exitosamente en: {pdf_path}")
            return True
        else:
            print(f"Error: No se pudo crear el PDF en {pdf_path}")
            return False
    except Exception as e:
        print(f"Error al convertir el documento: {str(e)}")
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
