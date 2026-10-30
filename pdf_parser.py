import pdfplumber


def extract_text(file_path):
    text = ""

    with pdfplumber.open(file_path) as pdf:

        for page in pdf.pages:

            if page.extract_text():
                text += page.extract_text()

    return text