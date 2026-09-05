/**
 * PRESCRITOR BIOLÓGICO SYNGENTA - BACKEND
 * Tratamento rigoroso de tipos e leitura da aba 'Base_Produtos'
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.autorizacaoPrevia = validarChaveUrl(e);
  
  return template.evaluate()
    .setTitle('Prescritor Biológico | Syngenta')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function validarChaveUrl(e) {
  if (!e || !e.parameter) return false;
  const chave = e.parameter.chave;
  const pinCorreto = ScriptProperties.getProperty('APP_PIN') || '7410';
  return chave === pinCorreto;
}

function validarSenha(pin) {
  const pinCorreto = ScriptProperties.getProperty('APP_PIN') || '7410';
  return String(pin).trim() === pinCorreto;
}

/**
 * Converte entradas da planilha para números válidos (trata vírgulas e pontos decimais)
 */
function parseSheetNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  
  let str = String(val).trim().replace(/\s/g, '');
  if (str.includes(',') && !str.includes('.')) {
    str = str.replace(',', '.');
  } else if (str.includes(',') && str.includes('.')) {
    str = str.replace(/\./g, '').replace(',', '.');
  }
  
  const parsed = parseFloat(str);
  return isNaN(parsed) ? 0 : parsed;
}

function getDadosPrescricao(chaveInput) {
  const pinCorreto = ScriptProperties.getProperty('APP_PIN') || '7410';
  
  if (String(chaveInput).trim() !== pinCorreto) {
    throw new Error("PIN_INCORRETO");
  }

  const sheetId = ScriptProperties.getProperty('SPREADSHEET_ID');
  let ss;

  if (sheetId && !sheetId.includes('<') && !sheetId.includes('ID_DA_SUA_PLANILHA')) {
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (err) {
      throw new Error(`Falha ao conectar no ID configurado: ${sheetId}.`);
    }
  } else {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (err) {
      throw new Error("A propriedade 'SPREADSHEET_ID' não foi configurada nas Script Properties.");
    }
  }

  const sheet = ss.getSheetByName('Base_Produtos') || ss.getSheets()[0];
  if (!sheet) throw new Error("Aba de produtos não localizada na planilha.");

  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return [];

  const cabecalhos = dados[0].map(c => String(c).trim().toLowerCase());
  
  const idxId = cabecalhos.findIndex(c => c === 'id');
  const idxCultura = cabecalhos.findIndex(c => c.includes('cultura'));
  const idxAlvo = cabecalhos.findIndex(c => c.includes('alvo') || c.includes('praga') || c.includes('doenca'));
  const idxProduto = cabecalhos.findIndex(c => c.includes('produto'));
  const idxTipo = cabecalhos.findIndex(c => c.includes('tipo') || c.includes('aplicacao'));
  const idxDoseHa = cabecalhos.findIndex(c => c.includes('dose'));

  let idxUnidDose = cabecalhos.findIndex((c, i) => c.includes('unidade') && i > idxDoseHa);
  if (idxUnidDose === -1) idxUnidDose = 6;

  const idxVolCalda = cabecalhos.findIndex(c => c.includes('vol') || c.includes('calda'));

  let idxUnidCalda = cabecalhos.findIndex((c, i) => c.includes('unidade') && i > idxVolCalda);
  if (idxUnidCalda === -1) idxUnidCalda = 8;

  const idxObs = cabecalhos.findIndex(c => c.includes('obs') || c.includes('tecnica') || c.includes('observacao'));

  const linhas = dados.slice(1);

  return linhas.map((linha, index) => {
    return {
      id: idxId !== -1 && linha[idxId] !== "" ? linha[idxId] : index + 1,
      cultura: idxCultura !== -1 ? String(linha[idxCultura]).trim() : '',
      alvo: idxAlvo !== -1 ? String(linha[idxAlvo]).trim() : '',
      produto: idxProduto !== -1 ? String(linha[idxProduto]).trim() : '',
      tipoAplicacao: idxTipo !== -1 && linha[idxTipo] ? String(linha[idxTipo]).trim() : 'Pulverização',
      doseHa: idxDoseHa !== -1 ? parseSheetNumber(linha[idxDoseHa]) : 0,
      unidadeDose: idxUnidDose < linha.length && linha[idxUnidDose] ? String(linha[idxUnidDose]).trim() : 'L/ha',
      volCaldaHa: idxVolCalda !== -1 ? parseSheetNumber(linha[idxVolCalda]) : 0,
      unidadeCalda: idxUnidCalda < linha.length && linha[idxUnidCalda] ? String(linha[idxUnidCalda]).trim() : 'L/ha',
      observacoes: idxObs !== -1 && linha[idxObs] ? String(linha[idxObs]).trim() : ''
    };
  }).filter(item => item.cultura !== '' && item.produto !== '');
}