from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Depends
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import shutil
from pathlib import Path
import logging
import time
from typing import Optional

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuración de variables de entorno

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
    Convierte un documento Word (.docx) a PDF usando LibreOffice.
    Esta implementación intenta preservar el formato original del documento Word.
    
    Args:
        docx_path: Ruta al archivo Word
        pdf_path: Ruta donde se guardará el PDF
    """
    try:
        # Usar LibreOffice para la conversión
        logger.info("Usando LibreOffice para la conversión")
        return await convert_with_libreoffice(docx_path, pdf_path)
    except Exception as e:
        logger.error(f"Error al convertir el documento: {str(e)}")
        return False

async def convert_with_libreoffice(docx_path, pdf_path):
    """
    Convierte un documento Word a PDF usando LibreOffice en modo headless.
    Requiere que LibreOffice esté instalado en el sistema.
    """
    import subprocess
    import platform
    
    try:
        # Determinar el comando según el sistema operativo
        if platform.system() == "Windows":
            # En Windows, buscar la instalación de LibreOffice o Microsoft Word
            libreoffice_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                # Agregar más rutas si es necesario
            ]
            
            # Encontrar la primera ruta válida
            soffice_path = None
            for path in libreoffice_paths:
                if os.path.exists(path):
                    soffice_path = path
                    break
            
            if not soffice_path:
                logger.error("No se encontró LibreOffice instalado")
                return False
            
            # Comando para Windows
            cmd = [
                soffice_path,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', os.path.dirname(pdf_path),
                str(docx_path)
            ]
        else:
            # Comando para Linux/Mac
            cmd = [
                'libreoffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', os.path.dirname(pdf_path),
                str(docx_path)
            ]
        
        # Ejecutar el comando
        logger.info(f"Ejecutando comando: {' '.join(cmd)}")
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = process.communicate()
        
        # Verificar si la conversión fue exitosa
        if process.returncode == 0:
            # LibreOffice guarda el archivo con el mismo nombre pero extensión .pdf
            # Necesitamos renombrarlo si el nombre de destino es diferente
            generated_pdf = os.path.splitext(docx_path)[0] + ".pdf"
            if generated_pdf != pdf_path:
                if os.path.exists(generated_pdf):
                    shutil.move(generated_pdf, pdf_path)
            
            logger.info(f"PDF creado exitosamente con LibreOffice en: {pdf_path}")
            return True
        else:
            logger.error(f"Error al ejecutar LibreOffice: {stderr.decode()}")
            return False
    except Exception as e:
        logger.error(f"Error al usar LibreOffice: {str(e)}")
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


# Verificar que LibreOffice esté instalado
def check_libreoffice_installation():
    import subprocess
    import platform
    
    try:
        if platform.system() == "Windows":
            # En Windows, verificar las rutas comunes
            libreoffice_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ]
            
            for path in libreoffice_paths:
                if os.path.exists(path):
                    logger.info(f"LibreOffice encontrado en: {path}")
                    return True
            
            logger.warning("No se encontró LibreOffice en las rutas comunes de Windows")
            return False
        else:
            # En Linux/Mac, verificar usando which
            result = subprocess.run(['which', 'libreoffice'], capture_output=True, text=True)
            if result.returncode == 0:
                logger.info(f"LibreOffice encontrado en: {result.stdout.strip()}")
                return True
            else:
                # Intentar verificar directamente
                try:
                    version_check = subprocess.run(['libreoffice', '--version'], capture_output=True, text=True)
                    if version_check.returncode == 0:
                        logger.info(f"LibreOffice versión: {version_check.stdout.strip()}")
                        return True
                except Exception:
                    pass
                
                logger.warning("No se encontró LibreOffice en el sistema")
                return False
    except Exception as e:
        logger.error(f"Error al verificar la instalación de LibreOffice: {str(e)}")
        return False

# Verificar la instalación al inicio
libreoffice_available = check_libreoffice_installation()
if not libreoffice_available:
    logger.warning("⚠️ LibreOffice no está disponible. La conversión de documentos puede fallar.")

if __name__ == "__main__":
    import uvicorn
    
    # Determinar el puerto desde la variable de entorno o usar 8080 por defecto
    port = int(os.environ.get("PORT", 8080))
    
    logger.info(f"Iniciando servidor en el puerto {port}")
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=True)
