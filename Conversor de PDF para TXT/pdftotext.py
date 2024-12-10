import fitz  # PyMuPDF

def extract_text_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ""
    
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text += page.get_text()

    return text

# Especifique o caminho do arquivo PDF
pdf_path = 'caminho.pdf'

# Extraia o texto
pdf_text = extract_text_from_pdf(pdf_path)

# Salve o texto extraído em um arquivo .txt (opcional)
with open("texto_extraido.txt", "w", encoding="utf-8") as text_file:
    text_file.write(pdf_text)

print("Texto extraído com sucesso!")
