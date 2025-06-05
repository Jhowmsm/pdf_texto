// balances.js

const fs = require('fs');
const path = require('path');
const { PDFExtract } = require('pdf.js-extract');
const { google } = require('googleapis');

// Leer configuración
const configPath = path.join(__dirname, 'config_balance.json');
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

const KEYWORDS = config.keywords;
const SHEET_NAME = config.sheet_name;
const WORKSHEET_NAME = config.worksheet_name;
const SPREADSHEET_ID = config.spreadsheet_id; // Asegúrate de incluir este campo en tu archivo de configuración

// Función para extraer texto del PDF
async function extractTextFromPDF(filePath) {
  const pdfExtract = new PDFExtract();
  const options = {}; // Opciones por defecto

  return new Promise((resolve, reject) => {
    pdfExtract.extract(filePath, options, (err, data) => {
      if (err) return reject(err);
      const text = data.pages
        .map(page => page.content.map(item => item.str).join(' '))
        .join('\n');
      resolve(text);
    });
  });
}

// Función para extraer valores según las palabras clave
function extractValues(text, keywordCellMap) {
  const data = {};
  for (const [keyword, info] of Object.entries(keywordCellMap)) {
    const cells = info.cells;
    const mode = info.mode || 'split';

    let pattern, match;
    switch (mode) {
      case 'until_dot':
        pattern = new RegExp(`${keyword}(.*?\\.)`, 's');
        match = text.match(pattern);
        data[cells[0]] = match ? match[1].trim() : 'No encontrado';
        break;

      case 'until_newline':
        pattern = new RegExp(`${keyword}(.*?)\\n`);
        match = text.match(pattern);
        data[cells[0]] = match ? match[1].trim() : 'No encontrado';
        break;

      case 'first_number_after':
        pattern = new RegExp(`${keyword}(.*?)`);
        match = text.match(pattern);
        if (match) {
          const followingText = text.slice(match.index + match[0].length);
          const found = followingText.match(/\S+/);
          data[cells[0]] = found ? found[0] : 'No encontrado';
        } else {
          data[cells[0]] = 'No encontrado';
        }
        break;

      case 'two_numbers_after':
        pattern = new RegExp(`${keyword}(.*?)`);
        match = text.match(pattern);
        if (match) {
          const followingText = text.slice(match.index + match[0].length);
          const tokens = followingText.match(/\S+/g) || [];

          const numericValues = [];
          let skipCount = 0;
          for (const token of tokens) {
            if (skipCount > 0) {
              skipCount--;
              continue;
            }

            if (token === 'Nota') {
              skipCount = 2;
              continue;
            }

            const cleaned = token.replace(/\./g, '').replace(',', '.').replace(/[()]/g, '');
            if (!isNaN(parseFloat(cleaned))) {
              numericValues.push(token);
              if (numericValues.length === 2) break;
            }
          }

          cells.forEach((cell, i) => {
            data[cell] = numericValues[i] || 'No encontrado';
          });
        } else {
          cells.forEach(cell => {
            data[cell] = 'No encontrado';
          });
        }
        break;

      case 'between_phrases':
        const startPhrase = info.start_phrase || '';
        const endPhrase = info.end_phrase || '';
        if (startPhrase && endPhrase) {
          pattern = new RegExp(`${startPhrase}(.*?)${endPhrase}`, 's');
          match = text.match(pattern);
          const extractedText = match ? match[1].replace(/\n/g, ' ').trim() : 'No encontrado';
          data[cells[0]] = extractedText;
        } else {
          data[cells[0]] = 'No encontrado';
        }
        break;

      default:
        pattern = new RegExp(`${keyword}\\s*([^\\n]+)`);
        match = text.match(pattern);
        if (match) {
          const values = match[1].trim().split(/\s+/);
          cells.forEach((cell, i) => {
            data[cell] = values[i] || 'No encontrado';
          });
        } else {
          cells.forEach(cell => {
            data[cell] = 'No encontrado';
          });
        }
        break;
    }
  }
  return data;
}

// Función para encontrar el primer NIF en el texto
function findNIFInText(text) {
  const nifPattern = /\b[A-Z]\d{8}\b/;
  const match = text.match(nifPattern);
  return match ? match[0] : 'No encontrado';
}

// Función para convertir valores
function convertirValor(valor) {
  if (typeof valor === 'string') {
    valor = valor.trim();
    if (valor.startsWith("'")) {
      valor = valor.slice(1);
    }
    if (valor.startsWith('(') && valor.endsWith(')')) {
      valor = '-' + valor.slice(1, -1);
    }
    valor = valor.replace(/\./g, '').replace(',', '.');
    const num = parseFloat(valor);
    return isNaN(num) ? valor : num;
  }
  return valor;
}

// Función para escribir en Google Sheets
async function writeToGoogleSheet(dataDict, nif) {
  const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, 'credentials.json'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

  // Actualizar la celda R10 con el NIF
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${WORKSHEET_NAME}!R10`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[nif]],
    },
  });

  // Escribir otros datos
  for (const [cell, value] of Object.entries(dataDict)) {
    const valorConvertido = convertirValor(value);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${WORKSHEET_NAME}!${cell}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[valorConvertido]],
      },
    });
  }
}

// Función principal
(async () => {
  try {
    const pdfPath = path.join(__dirname, 'documento.pdf'); // Reemplaza con la ruta a tu PDF
    const fullText = await extractTextFromPDF(pdfPath);

    const nif = findNIFInText(fullText);
    console.log(`Primer NIF encontrado: ${nif}`);

    const extractedData = extractValues(fullText, KEYWORDS);

    await writeToGoogleSheet(extractedData, nif);

    console.log('Datos escritos correctamente en Google Sheets.');
  } catch (error) {
    console.error('Error:', error);
  }
})();
