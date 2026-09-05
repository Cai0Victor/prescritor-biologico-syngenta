/**
 * SENHA DE ACESSO RESTRITO DA APLICAÇÃO
 */
const SENHA_ACESSO = '7410';

/**
 * Renderiza o WebApp e valida preliminarmente se a chave foi passada via URL (?chave=7410).
 * @param {Object} e Parâmetros HTTP GET.
 * @returns {HtmlOutput}
 */
function doGet(e) {
  const chaveUrl = (e && e.parameter && e.parameter.chave) ? String(e.parameter.chave).trim() : '';
  const autorizacaoPrevia = (chaveUrl === SENHA_ACESSO);

  const template = HtmlService.createTemplateFromFile('Index');
  template.autorizacaoPrevia = autorizacaoPrevia;
  
  return template.evaluate()
    .setTitle('Prescritor Biológicos | Syngenta')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Auxiliar para inclusão modular de arquivos HTML (CSS e JavaScript).
 * @param {string} filename Nome do arquivo na pasta src.
 * @returns {string} Conteúdo HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Valida o PIN informado no frontend.
 * @param {string} senhaInformada 
 * @returns {boolean}
 */
function validarSenha(senhaInformada) {
  return String(senhaInformada).trim() === SENHA_ACESSO;
}

/**
 * Converte valores no formato brasileiro (ex: "1,5" ou "1.250,50") para float.
 * @param {any} valor 
 * @returns {number}
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
 * Busca e filtra os produtos na planilha do Google Sheets.
 * Exige a senha de acesso válida para retornar os dados.
 * @param {string} senha Senha enviada pelo cliente.
 * @returns {Array<Object>} Lista de prescrições cadastradas.
 */
function getDadosPrescricao(senha) {
  if (String(senha).trim() !== SENHA_ACESSO) {
    throw new Error("Acesso não autorizado. Chave incorreta.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Script não está vinculado a uma planilha ativa.");
    }

    const sheet = ss.getSheetByName('Base_Produtos');
    if (!sheet) {
      throw new Error("A aba 'Base_Produtos' não foi encontrada na planilha.");
    }

    const data = sheet.getDataRange().getDisplayValues();
    if (!data || data.length <= 1) return [];

    // Remove cabeçalho
    data.shift();

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
    throw new Error("Erro no servidor: " + erro.message);
  }
}