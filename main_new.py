from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import subprocess
import logging
from pathlib import Path

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Crear directorios para archivos
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")

# Asegurarse de que los directorios existan
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(
    title="Word to PDF Converter API",
    description="API sencilla para convertir documentos Word a PDF",
    version="1.0.0"
)

# Configurar CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/convert/", summary="Convertir documento Word a PDF")
async def convert_word_to_pdf(file: UploadFile = File(...), background_tasks: BackgroundTasks = None):
    """
    Convierte un documento Word (.docx) a formato PDF.
    
    - **file**: Archivo Word (.docx) a convertir
    
    Retorna el archivo PDF convertido.
    """
    # Verificar que el archivo sea un documento Word
    if not file.filename.endswith(('.docx', '.doc')):
        logger.warning(f"Archivo no válido: {file.filename}")
        raise HTTPException(status_code=400, detail="El archivo debe ser un documento Word (.docx o .doc)")
    
    # Generar nombres únicos para los archivos
    file_id = str(uuid.uuid4())
    input_filename = f"{file_id}_{file.filename}"
    input_path = UPLOAD_DIR / input_filename
    
    try:
        # Guardar el archivo subido
        with open(input_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        logger.info(f"Archivo guardado en {input_path}")
        
        # Convertir a PDF usando LibreOffice
        pdf_filename = f"{Path(file.filename).stem}.pdf"
        output_pdf = await convert_to_pdf(str(input_path), str(OUTPUT_DIR))
        
        if not output_pdf:
            logger.error(f"Error al convertir {input_path}")
            raise HTTPException(status_code=500, detail="Error al convertir el documento")
        
        logger.info(f"Conversión exitosa: {output_pdf}")
        
        # Limpiar archivo temporal
        if background_tasks:
            background_tasks.add_task(lambda: os.remove(input_path) if os.path.exists(input_path) else None)
        
        # Devolver el archivo PDF
        return FileResponse(
            path=output_pdf,
            media_type="application/pdf",
            filename=pdf_filename
        )
        
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        # Limpiar archivo temporal en caso de error
        if os.path.exists(input_path):
            os.remove(input_path)
        raise HTTPException(status_code=500, detail="Error al convertir el documento")

async def convert_to_pdf(docx_path, output_dir):
    """
    Convierte un documento Word a PDF usando LibreOffice de manera simple.
    """
    try:
        # Nombre base del archivo sin extensión
        base_name = Path(docx_path).stem
        
        # Comando simple para convertir a PDF
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            docx_path
        ]
        
        logger.info(f"Ejecutando: {' '.join(cmd)}")
        
        # Ejecutar el comando
        process = subprocess.run(cmd, capture_output=True, text=True)
        
        # Registrar la salida
        if process.stdout:
            logger.info(f"Salida: {process.stdout}")
        if process.stderr:
            logger.warning(f"Error: {process.stderr}")
        
        # Verificar el archivo PDF generado
        expected_pdf = os.path.join(output_dir, f"{base_name}.pdf")
        
        if os.path.exists(expected_pdf):
            return expected_pdf
        else:
            # Listar archivos en el directorio para diagnóstico
            files = os.listdir(output_dir)
            logger.info(f"Archivos en directorio: {files}")
            
            # Buscar cualquier PDF generado
            for file in files:
                if file.endswith(".pdf") and file.startswith(Path(docx_path).name.split("_")[0]):
                    pdf_path = os.path.join(output_dir, file)
                    logger.info(f"PDF encontrado: {pdf_path}")
                    return pdf_path
            
            logger.error("No se encontró ningún PDF generado")
            return None
            
    except Exception as e:
        logger.error(f"Error en conversión: {str(e)}")
        return None

@app.get("/", summary="Información de la API")
async def root():
    """
    Retorna información básica sobre la API.
    """
    return {
        "mensaje": "API de conversión de Word a PDF",
        "descripcion": "Esta API permite convertir documentos Word a formato PDF",
        "uso": "Envía un archivo Word mediante POST a /convert/",
        "documentacion": "/docs",
        "version": "1.0.0"
    }

@app.get("/health", summary="Verificación de estado")
async def health_check():
    """
    Endpoint para verificar el estado del servicio.
    """
    return {"status": "ok", "message": "El servicio está funcionando correctamente"}

if __name__ == "__main__":
    import uvicorn
    
    # Determinar el puerto desde la variable de entorno o usar 8080 por defecto
    port = int(os.environ.get("PORT", 8080))
    
    logger.info(f"Iniciando servidor en el puerto {port}")
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=True)
