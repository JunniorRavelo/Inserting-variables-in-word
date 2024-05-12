from docx import Document
from docx.shared import Pt

def obtener_estilo(parrafo):
    estilo = {}
    if parrafo.style.name:
        estilo['name'] = parrafo.style.name
    if parrafo.style.font.bold:
        estilo['bold'] = True
    if parrafo.style.font.italic:
        estilo['italic'] = True
    if parrafo.style.font.underline:
        estilo['underline'] = True
    if parrafo.style.font.color.rgb:
        estilo['color'] = parrafo.style.font.color.rgb
    if parrafo.style.font.size:
        estilo['size'] = parrafo.style.font.size
    return estilo

def aplicar_estilo(texto, estilo):
    if 'name' in estilo:
        texto.style = estilo['name']
    if 'bold' in estilo:
        texto.bold = estilo['bold']
    if 'italic' in estilo:
        texto.italic = estilo['italic']
    if 'underline' in estilo:
        texto.underline = estilo['underline']
    if 'color' in estilo:
        texto.font.color.rgb = estilo['color']
    if 'size' in estilo:
        texto.font.size = estilo['size']

def remplazar_variables(documento, variables):
    for p in documento.paragraphs:
        for var, valor in variables.items():
            if var in p.text:
                partes = p.runs
                for parte in partes:
                    estilo = obtener_estilo(parte)
                    texto_reemplazo = parte.text.replace(var, valor)
                    parte.clear()
                    parte.text = texto_reemplazo  # Establecer el texto en la parte existente
                    aplicar_estilo(parte, estilo)

    for table in documento.tables:
        for row in table.rows:
            for cell in row.cells:
                for var, valor in variables.items():
                    if var in cell.text:
                        partes = cell.paragraphs[0].runs
                        for parte in partes:
                            estilo = obtener_estilo(parte)
                            texto_reemplazo = parte.text.replace(var, valor)
                            parte.clear()
                            parte.text = texto_reemplazo  # Establecer el texto en la parte existente
                            aplicar_estilo(parte, estilo)

# Ruta del archivo de Word
ruta_archivo = "spanish.docx"

# Cargar el documento de Word
doc = Document(ruta_archivo)

# Variables y sus valores
variables = {
    "[var1]": "Viaje",
    "[var2]": "Conejolandia",
    "[var3]": "Saltarín",
    "[var4]": "Fantasía",
    "[var5]": "Aladino"
}

# Reemplazar las variables en el documento
remplazar_variables(doc, variables)

# Guardar el documento modificado
doc.save("successfully_edited.docx")
