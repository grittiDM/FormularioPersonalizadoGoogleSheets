function atualizarMenusConsulta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRespostas = ss.getSheetByName("Respostas");
  const abaConsulta = ss.getSheetByName("Consulta Conferencia");

  if (!abaRespostas || !abaConsulta) {
    Logger.log("Erro: Abas 'Respostas' ou 'Consulta Conferencia' não encontradas.");
    return;
  }

  const dados = abaRespostas.getDataRange().getValues();
  const idSelecionado = String(abaConsulta.getRange("C1").getValue()).trim(); 
  let fornecedorSelecionado = String(abaConsulta.getRange("C2").getValue()).trim(); 

  // FILTRA C2 (fornecedor) com base em C1 (ID)
  let fornecedoresFiltrados = [];
  if (idSelecionado) {
    fornecedoresFiltrados = [...new Set(dados
      .filter(row => String(row[1]).trim() == idSelecionado) // Coluna B = index 1 (ID na aba Respostas)
      .map(row => String(row[6]).trim()) // Coluna G = index 6 (Fornecedor na aba Respostas)
      .filter(val => val) // Remove valores vazios
    )].sort();
  }

  // Lógica para C2 (Fornecedor)
  if (fornecedorSelecionado && !fornecedoresFiltrados.includes(fornecedorSelecionado)) {
    abaConsulta.getRange("C2").clearContent(); 
    fornecedorSelecionado = ""; 
  }

  const validacaoC2 = SpreadsheetApp.newDataValidation()
    .requireValueInList(fornecedoresFiltrados, true)
    .setAllowInvalid(false)
    .build();
  abaConsulta.getRange("C2").setDataValidation(validacaoC2); 

  // Se C2 for limpo ou alterado, é bom limpar G2 e K2.
  if (!fornecedorSelecionado && abaConsulta.getRange("G2").getValue()) {
      abaConsulta.getRange("G2").clearContent();
  }
  if (!fornecedorSelecionado && abaConsulta.getRange("K2").getValue()) {
      abaConsulta.getRange("K2").clearContent();
  }
}

// Executa ao abrir
function onOpen() {
  atualizarMenusAno(); // Mantém a validação de G2 e K2
  atualizarMenusConsulta(); // Atualiza C2
  // A filtragem principal agora será acionada por um botão
  // filtrarDadosEPreencherDatas(); // REMOVIDO: Não mais no onOpen
}

// Executa ao editar C1 ou C2 (e para as chamadas de atualização do menu)
function onEdit(e) {
  const aba = e.range.getSheet();
  const celula = e.range.getA1Notation();

  if (aba.getName() === "Consulta Conferencia") {
    if (celula === "C1") { // ID do Contrato
      // Se C1 foi editado, limpa C2, G2 e K2 antes de atualizar os menus
      e.range.getSheet().getRange("C2").clearContent(); 
      e.range.getSheet().getRange("G2").clearContent(); 
      e.range.getSheet().getRange("K2").clearContent(); // Limpa K2 também
      atualizarMenusConsulta(); // Atualiza C2 dropdown
      // filtrarDadosEPreencherDatas(); // REMOVIDO: Não mais automático ao editar C1
    } else if (celula === "C2") { // Fornecedor
      // Se C2 foi editado, limpa G2 e K2, atualiza menus
      e.range.getSheet().getRange("G2").clearContent();
      e.range.getSheet().getRange("K2").clearContent(); // Limpa K2 também
      atualizarMenusConsulta(); // Atualiza C2, G2, K2
      // filtrarDadosEPreencherDatas(); // REMOVIDO: Não mais automático ao editar C2
    } 
    // REMOVIDO: O bloco 'else if (celula === "G2" || celula === "K2") { ... }'
    // A filtragem principal agora será acionada por um botão quando G2 ou K2 forem editados.
  }
}