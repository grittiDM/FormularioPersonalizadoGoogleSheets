function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle("Conferência Mensal de Andaimes e Formas")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Ou .DEFAULT, se preferir
}

/**
 * Inclui o conteúdo de outro arquivo HTML no template atual.
 * @param {string} filename O nome do arquivo HTML (sem a extensão .html) a ser incluído.
 * @return {string} O conteúdo do arquivo HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function loadFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const itens = ss.getSheetByName("Cadastro de Itens").getRange("A2:A").getValues().flat().filter(String);

  const obrasSheet = ss.getSheetByName("Cadastro de Obras").getDataRange().getValues();
  const obras = {};
  for (let i = 1; i < obrasSheet.length; i++) {
    const row = obrasSheet[i];
    const numero = row[0];
    if (!numero) continue;
    obras[numero] = {
      nome: row[1],
      enderecos: [row[2], row[3], row[4], row[5]].filter(Boolean),
      contratos: []
    };
  }

  // Buscar contratos existentes na aba Respostas por nº da obra
  const respostasSheet = ss.getSheetByName("Respostas").getDataRange().getValues();
  for (let i = 1; i < respostasSheet.length; i++) {
    const row = respostasSheet[i];
    const numeroObra = row[1];
    const contrato = row[6];
    if (obras[numeroObra] && contrato && !obras[numeroObra].contratos.includes(contrato)) {
      obras[numeroObra].contratos.push(contrato);
    }
  }

  const fornecedoresSheet = ss.getSheetByName("Cadastro Fornecedores").getDataRange().getValues();
  const fornecedores = {};
  for (let i = 1; i < fornecedoresSheet.length; i++) {
    const row = fornecedoresSheet[i];
    const nome = row[0];
    if (!nome) continue;
    fornecedores[nome] = {
      cnpj: row[1] || "",
    };
  }

  return { itens, obras, fornecedores };
}

function submitData(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Respostas");

  const timestamp = new Date();
  dados.itens.forEach(item => {
    sheet.appendRow([
      timestamp, // Coluna Data Registro
      dados.numeroObra,
      dados.nomeObra,
      dados.endereco,
      dados.fornecedor,
      dados.tipoRegistro,
      dados.contrato,
      item.item,
      item.quantidade,
      dados.conferente
    ]);
  });

  return "Dados enviados com sucesso!";
}
