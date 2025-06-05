import fitz
import json
import gspread
import re
from oauth2client.service_account import ServiceAccountCredentials

# Leer configuración
with open('config_BSC.json') as f:
    config = json.load(f)

KEYWORDS = config["keywords"]
SHEET_NAME = config["sheet_name"]
WORKSHEET_NAME = config["worksheet_name"]

def extract_text_from_pdf(file_path):
    text = ""
    with fitz.open(file_path) as pdf:
        for page in pdf:
            text += page.get_text()
    return text

def extract_values(text, keyword_cell_map):
    data = {}
    for keyword, info in keyword_cell_map.items():
        cells = info["cells"]
        mode = info.get("mode", "split")

        if mode == "until_dot":
            pattern = re.escape(keyword) + r"(.*?\.)(?:\s|$)"
            match = re.search(pattern, text, re.DOTALL)
            if match:
                data[cells[0]] = match.group(1).strip()
            else:
                data[cells[0]] = "No encontrado"

        elif mode == "until_newline":
            pattern = re.escape(keyword) + r"(.*?)\n"
            match = re.search(pattern, text)
            if match:
                data[cells[0]] = match.group(1).strip()
            else:
                data[cells[0]] = "No encontrado"

        elif mode == "first_number_after":
            pattern = re.escape(keyword) + r"(.*?)"
            match = re.search(pattern, text, re.DOTALL)
            if match:
                following_text = text[match.end():]
                found = re.search(r"\S+", following_text)
                data[cells[0]] = found.group(0) if found else "No encontrado"
            else:
                data[cells[0]] = "No encontrado"

        elif mode == "two_numbers_after":
            pattern = re.escape(keyword) + r"(.*?)"
            match = re.search(pattern, text, re.DOTALL)
            if match:
                following_text = text[match.end():]
                tokens = re.findall(r"\S+", following_text)

                unwanted_word = "Nota"

                numeric_values = []
                skip_count = 0  # To skip "Nota" + next two tokens
                for token in tokens:
                    if skip_count > 0:
                        skip_count -= 1
                        continue

                    if token == unwanted_word:
                        skip_count = 1  # skip this token + next 2 tokens
                        continue

                    cleaned = token.replace(".", "").replace(",", ".").strip("()")
                    try:
                        float(cleaned)
                        numeric_values.append(token)
                        if len(numeric_values) == 2:
                            break
                    except ValueError:
                        continue

                for i, cell in enumerate(cells):
                    data[cell] = numeric_values[i] if i < len(numeric_values) else "No encontrado"
            else:
                for cell in cells:
                    data[cell] = "No encontrado"    

        elif mode == "between_phrases":
            start_phrase = info.get("start_phrase", "")
            end_phrase = info.get("end_phrase", "")
            if start_phrase and end_phrase:
                pattern = re.escape(start_phrase) + r"(.*?)" + re.escape(end_phrase)
                match = re.search(pattern, text, re.DOTALL)
                if match:
                    extracted_text = match.group(1).strip()
                    cleaned_text = extracted_text.replace("\n", " ").strip()  # Eliminar saltos de línea
                    data[cells[0]] = cleaned_text
                else:
                    data[cells[0]] = "No encontrado"
            else:
                data[cells[0]] = "No encontrado"

        else:
            pattern = re.escape(keyword) + r"\s*([^\n]+)"
            match = re.search(pattern, text)
            if match:
                values = match.group(1).strip().split()
                for i, cell in enumerate(cells):
                    data[cell] = values[i] if i < len(values) else "No encontrado"
            else:
                for cell in cells:
                    data[cell] = "No encontrado"
    return data

def find_nif_in_text(text):
    nif_pattern = r"\b[A-Z]\d{8}\b"
    match = re.search(nif_pattern, text)
    
    if match:
        return match.group(0)  # Devolver el primer NIF encontrado
    else:
        return "No encontrado"

def convertir_valor(valor):
    if isinstance(valor, str):
        valor = valor.strip()
        if valor.startswith("'"):
            valor = valor[1:]
        if valor.startswith("(") and valor.endswith(")"):
            valor = "-" + valor[1:-1]
        valor = valor.replace(".", "").replace(",", ".")
        try:
            return float(valor)
        except ValueError:
            return valor  # Si no es número válido, se deja como texto
    return valor

def write_to_google_sheet(data_dict, nif):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open(SHEET_NAME)
    worksheet = sheet.worksheet(WORKSHEET_NAME)

    # Actualizar la celda B10 con el NIF
    worksheet.update('R10', [[nif]])

    # Escribir otros datos
    for cell, value in data_dict.items():
        valor_convertido = convertir_valor(value)
        worksheet.update(range_name=cell, values=[[valor_convertido]])

if __name__ == "__main__":
    pdf_path = "BSC.pdf"
    full_text = extract_text_from_pdf(pdf_path)
    
    # Buscar el primer NIF en el texto
    nif = find_nif_in_text(full_text)
    print(f"Primer NIF encontrado: {nif}")
    
    # Extraer otros datos según la configuración
    extracted_data = extract_values(full_text, KEYWORDS)
    
    # Escribir los datos y el NIF en Google Sheets
    write_to_google_sheet(extracted_data, nif)
    
    print("Datos escritos correctamente en Google Sheets.")
