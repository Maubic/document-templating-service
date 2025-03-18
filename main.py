from fastapi import Body, FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pydantic import Json
from docxtpl import DocxTemplate, RichText
import aiofiles
from utils import remove_temporary_files, get_env
import requests
from bs4 import BeautifulSoup

app = FastAPI(
    title="Document Template Processing Service",
    description="""
        This is the documentation of the REST API exposed by the document template processing microservice.
        This will allow you to inject data in a specific word document template and get the pdf format as a result. 🚀🚀🚀
    """,
    version="1.0.0"
)

SERVICE_STATUS = {'status': 'Service is healthy !'}

@app.get('/')
async def livenessprobe():
    remove_temporary_files()
    return SERVICE_STATUS

@app.get('/health-check')
async def healthcheck():
    remove_temporary_files()
    return SERVICE_STATUS

def convert_html_to_richtext(html):
    """Convierte HTML en un objeto RichText compatible con docxtpl"""
    rt = RichText()
    soup = BeautifulSoup(html, "html.parser")

    for elem in soup.recursiveChildGenerator():
        if isinstance(elem, str):
            rt.add(elem)
        elif elem.name == "b":
            rt.add(elem.get_text(), bold=True)
        elif elem.name == "i":
            rt.add(elem.get_text(), italic=True)
        elif elem.name == "u":
            rt.add(elem.get_text(), underline=True)
        elif elem.name == "br":
            rt.add("\n")

    return rt

@app.post('/api/v1/process-template-document')
async def process_document_template(data: Json = Body(...), file: UploadFile = File(...)):
    if file.filename == '':
        return JSONResponse({'status': 'error', 'message': 'file is required'}, status_code=400)
    if data is None or len(data) == 0:
        return JSONResponse({'status': 'error', 'message': 'data is required'}, status_code=400)

    file_path = f'temp/{file.filename}'
    pdf_file_path = f'temp/{file.filename.split(".")[0]}.pdf'

    async with aiofiles.open(file_path, 'wb') as out_file:
        while content := await file.read(1024):
            await out_file.write(content)

    document = DocxTemplate(file_path)

    # Convertir solo las claves que contengan HTML
    for key, value in data.items():
        if isinstance(value, str) and ("<" in value and ">" in value):  # Detección básica de HTML
            data[key] = convert_html_to_richtext(value)

    document.render(data)
    document.save(file_path)

    # Enviar a Gotenberg para conversión a PDF
    resourceURL = f"{get_env('GOTENBERG_API_URL')}/forms/libreoffice/convert"
    response = requests.post(url=resourceURL, files={'file': open(file_path, 'rb')})

    async with aiofiles.open(pdf_file_path, 'wb') as out_file:
        await out_file.write(response.content)

    return FileResponse(pdf_file_path, media_type='application/pdf')