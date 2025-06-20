# Word to PDF Converter API

Una API sencilla desarrollada con FastAPI para convertir documentos Word (.docx, .doc) a formato PDF.

## Requisitos

- Python 3.7+
- Las dependencias listadas en `requirements.txt`

## Instalación

1. Instalar las dependencias:

```bash
pip install -r requirements.txt
```

## Ejecución

Para iniciar el servidor:

```bash
uvicorn main:app --reload
```

El servidor se iniciará en `http://localhost:8080`

## Uso de la API

### Convertir un documento Word a PDF

**Endpoint:** `POST /convert/`

**Parámetros:**
- `file`: Archivo Word (.docx o .doc) a convertir

**Ejemplo usando curl:**
```bash
curl -X POST "http://localhost:8080/convert/" -H "accept: application/json" -H "Content-Type: multipart/form-data" -F "file=@documento.docx"
```

**Ejemplo usando Python requests:**
```python
import requests

url = "http://localhost:8080/convert/"
files = {"file": open("documento.docx", "rb")}
response = requests.post(url, files=files)

# Guardar el PDF resultante
with open("documento.pdf", "wb") as f:
    f.write(response.content)
```

## Documentación de la API

La documentación interactiva de la API está disponible en:
- Swagger UI: `http://localhost:8080/docs`
- ReDoc: `http://localhost:8080/redoc`
