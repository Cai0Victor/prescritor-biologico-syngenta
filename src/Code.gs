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

    data.shift(); // Remove a linha do cabeçalho

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
    throw new Error(erro.message);
  }
}

/**
 * Gera um PDF oficial da prescrição técnica e armazena no Google Drive.
 * @param {Object} dados Dados formatados da prescrição.
 * @returns {string} URL pública do PDF gerado.
 */
function gerarPDFPrescricao(dados) {
  try {
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial, sans-serif; color: #333; margin: 30px; }
          .header { background-color: #008751; color: white; padding: 20px; border-radius: 6px; }
          .title { font-size: 20px; font-weight: bold; margin: 0; }
          .subtitle { font-size: 13px; margin-top: 4px; opacity: 0.9; }
          .section { margin-top: 20px; border-bottom: 1px solid #cbd5e1; padding-bottom: 12px; }
          .field { margin: 6px 0; font-size: 13px; }
          .label { font-weight: bold; color: #475569; }
          .box { background-color: #f0fdf4; border-left: 4px solid #008751; padding: 12px; margin-top: 15px; }
          .footer { margin-top: 30px; font-size: 10px; text-align: center; color: #94a3b8; }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="title">SYNGENTA BIOLOGICALS</div>
          <div class="subtitle">Prescrição Técnica de Produtos Biológicos</div>
        </div>

        <div class="section">
          <div class="field"><span class="label">Cultura:</span> ${dados.cultura}</div>
          <div class="field"><span class="label">Alvo / Praga / Doença:</span> ${dados.alvo}</div>
          <div class="field"><span class="label">Produto Biológico:</span> <strong>${dados.produto}</strong></div>
          <div class="field"><span class="label">Área Aplicada:</span> ${dados.area} ha</div>
        </div>

        <div class="section">
          <div class="field"><span class="label">Dose por Hectare:</span> ${dados.doseHa} ${dados.unidadeDose}</div>
          <div class="field"><span class="label">Volume de Calda/ha:</span> ${dados.volCaldaHa} ${dados.unidadeCalda}</div>
        </div>

        <div class="box">
          <div class="field"><span class="label">PRODUTO TOTAL:</span> <strong>${dados.prodTotal}</strong></div>
          <div class="field"><span class="label">CALDA TOTAL:</span> <strong>${dados.caldaTotal}</strong></div>
        </div>

        ${dados.observacoes ? `
          <div class="section">
            <div class="label">Observações Técnicas:</div>
            <p style="font-size: 12px; color: #334155;">${dados.observacoes}</p>
          </div>
        ` : ''}

        <div class="footer">
          Documento gerado via Prescritor Biológico Syngenta em ${new Date().toLocaleDateString('pt-BR')}.
        </div>
      </body>
      </html>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
    blob.setName(`Prescricao_${dados.produto.replace(/\s+/g, '_')}.pdf`);

    const arquivo = DriveApp.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return arquivo.getUrl();
  } catch (erro) {
    throw new Error("Erro na geração do PDF: " + erro.message);
  }
}