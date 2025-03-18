from fastapi import Body, FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pydantic import Json
from docxtpl import DocxTemplate
import aiofiles
from utils import remove_temporary_files, get_env
import requests
import re
from docx import Document
from bs4 import BeautifulSoup
from copy import deepcopy
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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

def insert_paragraph_after(paragraph):
    """Insert a new paragraph after the given paragraph."""
    p = paragraph._p
    parent = p.getparent()
    new_p = OxmlElement('w:p')
    parent.insert(parent.index(p) + 1, new_p)
    return docx.text.paragraph.Paragraph(new_p, paragraph._parent)

@app.post('/api/v1/process-template-document')
async def process_document_template(data: Json = Body(...), file: UploadFile = File(...)):
    if file.filename == '':
        return JSONResponse({'status': 'error', 'message': 'file is required'}, status_code=400)
    if data is None or len(data) == 0:
        return JSONResponse({'status': 'error', 'message': 'data is required'}, status_code=400)
    
    resourceURL = '{}/forms/libreoffice/convert'.format(get_env('GOTENBERG_API_URL')) 
    file_path = 'temp/{}'.format(file.filename)
    result_file_path = 'temp/result_{}'.format(file.filename)
    pdf_file_path = 'temp/{}.pdf'.format(file.filename.split('.')[0])
    
    async with aiofiles.open(file_path, 'wb') as out_file:
        while content := await file.read(1024):
            await out_file.write(content)
    
    # Paso 1: Identificar campos HTML y crear placeholders
    processed_data = {}
    html_fields = {}
    
    for key, value in data.items():
        if key.endswith('_html') and isinstance(value, str):
            # Guardar en un diccionario aparte
            html_fields[key] = value
            # Usar un placeholder único
            base_key = key[:-5]
            placeholder = f"##HTML_CONTENT_{base_key}##"
            processed_data[base_key] = placeholder
        else:
            processed_data[key] = value
    
    # Paso 2: Renderizar la plantilla con los placeholders
    document = DocxTemplate(file_path)
    document.render(processed_data)
    document.save(result_file_path)
    
    # Paso 3: Abrir el documento con python-docx y reemplazar los placeholders
    doc = Document(result_file_path)
    
    # Procesar cada párrafo
    i = 0
    while i < len(doc.paragraphs):
        paragraph = doc.paragraphs[i]
        found_placeholder = False
        placeholder_processed = False
        
        for html_key, html_value in html_fields.items():
            base_key = html_key[:-5]
            placeholder = f"##HTML_CONTENT_{base_key}##"
            
            if placeholder in paragraph.text:
                found_placeholder = True
                # Limpiar el párrafo actual manteniendo sus propiedades
                paragraph_style = paragraph.style
                current_p = paragraph
                current_p.clear()
                
                # Procesamos el HTML
                wrapped_html = f"<div>{html_value}</div>"
                soup = BeautifulSoup(wrapped_html, 'html.parser')
                
                # Extraer texto inicial (antes de p, ol, ul)
                initial_text = ""
                for node in soup.div.contents:
                    if isinstance(node, str):
                        initial_text += node
                    elif node.name not in ['p', 'ol', 'ul']:
                        initial_text += str(node)
                
                # Procesar texto inicial con formatos
                if initial_text.strip():
                    # Procesar negritas
                    bold_pattern = re.compile(r'<b>(.*?)</b>|<strong>(.*?)</strong>')
                    bold_matches = list(bold_pattern.finditer(initial_text))
                    
                    # Procesar cursivas
                    italic_pattern = re.compile(r'<i>(.*?)</i>|<em>(.*?)</em>')
                    italic_matches = list(italic_pattern.finditer(initial_text))
                    
                    # Extraer el texto plano
                    plain_text = BeautifulSoup(initial_text, 'html.parser').get_text()
                    
                    # Si hay formatos, procesarlos
                    if bold_matches or italic_matches:
                        # Textos procesados para no duplicar
                        processed_texts = []
                        
                        # Procesar negritas
                        for match in bold_matches:
                            content = match.group(1) or match.group(2)
                            run = current_p.add_run(content)
                            run.bold = True
                            processed_texts.append(content)
                        
                        # Procesar cursivas
                        for match in italic_matches:
                            content = match.group(1) or match.group(2)
                            if content not in processed_texts:  # Evitar duplicación
                                run = current_p.add_run(content)
                                run.italic = True
                                processed_texts.append(content)
                        
                        # Añadir el resto del texto plano
                        remaining_text = plain_text
                        for text in processed_texts:
                            remaining_text = remaining_text.replace(text, '', 1)
                        
                        if remaining_text.strip():
                            current_p.add_run(remaining_text.strip())
                    else:
                        # Si no hay formatos, añadir todo el texto plano
                        if plain_text.strip():
                            current_p.add_run(plain_text.strip())
                
                # Procesar elementos estructurales en orden
                curr_index = i
                for element in soup.div.children:
                    if element.name == 'p':
                        # Crear párrafo para contenido <p>
                        p_text = element.get_text().strip()
                        if p_text:
                            curr_index += 1
                            new_p = insert_paragraph_after(doc.paragraphs[curr_index-1])
                            new_p.text = p_text
                    
                    elif element.name in ['ol', 'ul']:
                        # Procesar listas
                        is_ordered = element.name == 'ol'
                        list_items = element.find_all('li', recursive=False)
                        
                        for idx, item in enumerate(list_items):
                            curr_index += 1
                            new_p = insert_paragraph_after(doc.paragraphs[curr_index-1])
                            new_p.style = 'List Number' if is_ordered else 'List Bullet'
                            prefix = f"{idx+1}. " if is_ordered else "• "
                            new_p.text = prefix + item.get_text().strip()
                
                placeholder_processed = True
                # Actualizar el índice para saltar los párrafos insertados
                i = curr_index
                break
        
        if not placeholder_processed:
            i += 1
    
    # Guardar el documento procesado
    doc.save(result_file_path)
    
    # Paso 4: Convertir a PDF
    response = requests.post(url=resourceURL, files={'file': open(result_file_path, 'rb')})
    async with aiofiles.open(pdf_file_path, 'wb') as out_file:
        await out_file.write(response.content)
    return FileResponse(pdf_file_path, media_type='application/pdf')