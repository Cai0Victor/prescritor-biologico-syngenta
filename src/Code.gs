/**
 * Renderiza o WebApp do Prescritor Biológico Syngenta.
 * Função reservada do Google Apps Script para requisições HTTP GET.
 * @param {Object} e Parâmetros da requisição HTTP.
 * @returns {HtmlOutput} Página HTML renderizada.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Prescritor Biológicos | Syngenta')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Auxiliar para inclusão modular de arquivos HTML (CSS e JavaScript).
 * @param {string} filename Nome do arquivo a ser incluído.
 * @returns {string} Conteúdo HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Converte strings/números do formato PT-BR para float válido do JS.
 * @param {any} valor Valor vindo da célula.
 * @returns {number} Número convertido.
 */
function parseNumeroPTBR(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;
  if (typeof valor === 'number') return isNaN(valor) ? 0 : valor;
  
  let str = String(valor).trim();
  if (str.includes(',')) {
    str = str.replace(/\./g, '').replace(',', '.');
  }
  
  let num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

/**
 * Lê os dados da aba 'Base_Produtos' na planilha ativa.
 * @returns {Array<Object>} Lista de produtos estruturada.
 */
function getDadosPrescricao() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("O script não está vinculado a uma planilha ativa do Google Sheets.");
    }

    const sheet = ss.getSheetByName('Base_Produtos');
    if (!sheet) {
      throw new Error("A aba 'Base_Produtos' não foi encontrada na planilha.");
    }

    const data = sheet.getDataRange().getDisplayValues();
    if (!data || data.length <= 1) return [];

    data.shift(); // Remove a linha de cabeçalho

    return data
      .filter(row => row[1] && String(row[1]).trim() !== '')
      .map((row, index) => ({
        id: row[0] || (index + 1),
        cultura: String(row[1]).trim(),
        alvo: String(row[2] || '').trim(),
        produto: String(row[3] || '').trim(),
        doseHa: parseNumeroPTBR(row[4]),
        unidadeDose: String(row[5] || 'L/ha').trim(),
        volCaldaHa: parseNumeroPTBR(row[6]),
        unidadeCalda: String(row[7] || 'L/ha').trim(),
        observacoes: String(row[8] || '').trim()
      }));
  } catch (erro) {
    throw new Error("Erro no backend: " + erro.message);
  }
}