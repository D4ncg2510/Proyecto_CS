import openpyxl
import docx
import PyPDF2



def xlsx_meta(nombre):
    wb = openpyxl.load_workbook(nombre)
    metadata = {}

    # Obtener propiedades principales
    props = wb.properties
    metadata['Titulo'] = props.title
    metadata['Creador'] = props.creator
    metadata['keywords'] = props.keywords
    metadata['descripción'] = props.description
    metadata['Creación'] = props.created.isoformat() if props.created else None
    metadata['Modificado'] = props.modified.isoformat() if props.modified else None

    return metadata

def docx_meta(nombre):
    doc = docx.Document(nombre)
    metadata = {}

    # Obtener propiedades principales
    core_properties = doc.core_properties
    metadata['Titulo'] = core_properties.title
    metadata['Autor'] = core_properties.author
    metadata['Creación'] = core_properties.created
    metadata['Modificado'] = core_properties.modified
    metadata['Última modificación'] = core_properties.last_modified_by
    metadata['Revision'] = core_properties.revision
    metadata['version'] = core_properties.version

    return metadata

def pdf_meta(nombre):
    with open(nombre, 'rb') as f:
        pdf_reader = PyPDF2.PdfReader(f)
        metadata = {}
        info = pdf_reader.metadata
        for key in info:
            metadata[key] = info[key]
        return metadata

if __name__ == "__main__":

    # Extracción de datos del PDF
    print("\n∞∞∞ Metadatos de PDF ∞∞∞")
    #se llama a la función, en esta parte se debe de colocar la ruta al archivo del cual se quieren extraer datos
    pdf_metadata = pdf_meta("/home/kali/Desktop/ticketdecompra.pdf")
    for key, value in pdf_metadata.items():
        print(f"{key}: {value}")
        
    # Extracción de datos del DOCX
    print("\n∞∞∞ Metadatos de DOCX ∞∞∞")
    #se llama a la función, en esta parte se debe de colocar la ruta al archivo del cual se quieren extraer datos
    docx_metadata = docx_meta("/home/kali/Desktop/Doctos.docx")
    for key, value in docx_metadata.items():
        print(f"{key}: {value}")
        
    # Extracción de datos del XLSX
    print("\n∞∞∞ Metadatos de XLSX ∞∞∞")
    #se llama a la función, en esta parte se debe de colocar la ruta al archivo del cual se quieren extraer datos
    xlsx_metadata = xlsx_meta("/home/kali/Desktop/Tabla.xlsx")
    for key, value in xlsx_metadata.items():
        print(f"{key}: {value}")

# Esto solo de debe de compilar en la terminal, usando el comando python obtmet.py, las rutas al archivo se asignaran aqui en los apartados de extracción de datos

    

    
