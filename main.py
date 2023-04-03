"""
Python Docx
  pip install python-docx
  https://python-docx.readthedocs.io/en/latest/

Python PYPDF2
  pip install PyPDF2
  https://pypi.org/project/docx2pdf/

DOCX2PDF
  pip install docx2pdf
  https://pypi.org/project/docx2pdf/

rut:
, C.I. N° 
, con nota final 
"""
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import re 
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfWriter, PdfReader

def getFromDocx(doc, regex):
    """
    Busca parrafo por parrafo un string en base a una expresion regular.
    retorna una lista con todas la coicidencias y un numero que indica la pagina en la que fue encontrada
    """
    names = []
    i = 0
    for para in doc.paragraphs:
        text = para.text
        if not len(text):
            continue

        results = re.search(regex, text)
        if results:
            name = results.group(1).strip()
            names.append([i, name])
            i += 1
    return names

def splitPdf(reader, names, destDir):
    """
    Recibe el documento pdf, la lista con los nombres y el directorio de destino.
    Divide el pdf y le asigna el nombre correspondiente a la pagina
    """
    length = len(names)
    for index, name in names:
        try:
            page = reader.pages[index]
            pdfOutDir = f'{destDir}/{name}.pdf'
            pdfOut = open(pdfOutDir, 'wb') 
            
            writer = PdfWriter()
            writer.add_page(page)
            writer.write(pdfOut)
            print(f"{name}.pdf creado correctamente.")
        except:
            print(f"\n{name}.pdf no ha podido ser creado.\n")

def wordToPdf(options):
    print(f"""
    Archivo Original: {options["docDir"]}
    Destino: {options["destDir"]}
    Frase de inicio: {options["start"].get()}
    Frase de fin: {options["stop"].get()}\n""")
    #Generacion de la expresion regular
    start = options["start"].get()
    stop = options["stop"].get()
    regex = re.escape(start) + r'(.*?)' + re.escape(stop) 
    
    # Crear un pdf en base al docx
    docDir = options["docDir"]
    destDir = options["destDir"]
    pdfDir = destDir + "/" + docDir.replace('.docx', '').split("/")[-1] + ".pdf"
    convert(docDir, pdfDir)

    doc = Document(docDir)
    names = getFromDocx(doc, regex)
    
    reader = PdfReader(pdfDir)
    splitPdf(reader, names, destDir)

def selDoc(docDirLabel):
    docDir = filedialog.askopenfile(
        initialdir='/',
        title="Seleccione un documento...",
        filetypes=[("Documento Word", ".docx"),])
    
    options["docDir"] = docDir.name
    docDirLabel.configure(text=docDir.name)

def selDest(destDirLabel):
    destDir = filedialog.askdirectory(
        initialdir='/', 
        title="Seleccione una carpeta destino...")
    
    options["destDir"] = destDir
    destDirLabel.configure(text=destDir)

#########################################################################

root = Tk()

root.geometry("500x500")
root.title("Word a Pdfs")

start = StringVar()
stop = StringVar()
start.set("don(ña)")
stop.set(", C.I")

options = {
  "docDir" : "",
  "destDir" : "",
  "start" : start,
  "stop" : stop
}

ttk.Label(root, text="Indique la direccion del documento: ", 
          justify="left", font=("Arial", 13)).grid(column=0, row=0, sticky='w', pady=(10,0)) 
docDirLabel = Label(root, text="Seleccione un documento...", justify="left", font=("Arial", 12))
docDirLabel.grid(column=0, row=1, sticky='w')
ttk.Button(root, text="Examinar", command=lambda: selDoc(docDirLabel)).grid(column=0, row=2, sticky='w')

ttk.Label(root, text="Indique donde quiere crear los archivos: ", 
          justify="left", font=("Arial", 13)).grid(column=0, row=3, sticky='w', pady=(10,0)) 
destDirLabel = Label(root, text="Seleccione una carpeta...", justify="left", font=("Arial", 12))
destDirLabel.grid(column=0, row=4, sticky='w')
ttk.Button(root, text="Examinar", command=lambda: selDest(destDirLabel)).grid(column=0, row=5, sticky='w')

ttk.Label(root, text="""Ingrese una frase de inicio y una de fin. 
Cada palabra que esté dentro de estas dos frases 
será utilizada para nombrar los archivos""", 
  justify="left", font=("Arial", 13)).grid(column=0, row=6, sticky='w', pady=(10,0))
ttk.Entry(root, textvariable=start, width=25, ).grid(column=0, row=7, sticky='w', padx=5)
ttk.Entry(root, textvariable=stop, width=25).grid(column=0, row=8, sticky='w', padx=5, pady=(5,0))

ttk.Button(root, text="Confirmar", command=lambda: wordToPdf(options), width=40).grid(columnspan=2, column=0, row=9, sticky='w', pady=10)

root.mainloop()