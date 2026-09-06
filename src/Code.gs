/**
 * ============================================================================
 * PRESCRITOR BIOLÓGICO SYNGENTA - BACKEND (GOOGLE APPS SCRIPT)
 * ============================================================================
 * Arquivo: src/Code.gs
 * Função: Servidor API JSON e leitor da planilha "Prescritor Biologico - Base"
 * ============================================================================
 */

function doGet(e) {
  // Retorno da API JSON para sincronização offline (GitHub Pages / PWA)
  if (e && e.parameter && e.parameter.action === 'getDados') {
    const dados = getDadosProdutos();
    return ContentService.createTextOutput(JSON.stringify(dados))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Renderização padrão Apps Script
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Prescritor Biológico | Syngenta')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDadosProdutos() {
  try {
    const NOME_PLANILHA = "Prescritor Biologico - Base";
    const NOME_ABA = "Base_Produtos";

    let spreadsheet;

    try {
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active && active.getName() === NOME_PLANILHA) {
        spreadsheet = active;
      }
    } catch (err) {}

    if (!spreadsheet) {
      const arquivos = DriveApp.getFilesByName(NOME_PLANILHA);
      if (!arquivos.hasNext()) {
        throw new Error(`A planilha "${NOME_PLANILHA}" não foi encontrada no seu Google Drive.`);
      }
      spreadsheet = SpreadsheetApp.open(arquivos.next());
    }

    const sheet = spreadsheet.getSheetByName(NOME_ABA);
    if (!sheet) {
      throw new Error(`A aba "${NOME_ABA}" não foi encontrada dentro da planilha.`);
    }

    const range = sheet.getDataRange();
    const data = range.getValues();

    if (data.length <= 1) return [];

    data.shift(); // Remove cabeçalho

    return data
      .filter(row => row[0] !== "" && row[1] !== "" && row[3] !== "")
      .map((row, index) => {
        const rawDose = String(row[5]).replace(',', '.');
        const rawVolCalda = String(row[7]).replace(',', '.');

        return {
          id: String(row[0] || index + 1).trim(),
          cultura: String(row[1] || '').trim(),
          alvo: String(row[2] || '').trim(),
          produto: String(row[3] || '').trim(),
          tipoAplicacao: String(row[4] || '').trim(),
          doseHa: parseFloat(rawDose) || 0,
          unidadeDose: String(row[6] || '').trim(),
          volCalda: parseFloat(rawVolCalda) || 0,
          unidadeCalda: String(row[8] || '').trim(),
          observacoes: String(row[9] || '').trim()
        };
      });

  } catch (error) {
    Logger.log("❌ Erro em getDadosProdutos(): " + error.toString());
    throw new Error(error.message);
  }
}