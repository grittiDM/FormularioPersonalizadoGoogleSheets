function atualizarMenusAno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRespostas = ss.getSheetByName("Respostas");
  const abaConsulta = ss.getSheetByName("Consulta Conferencia");
  
  if (!abaRespostas || !abaConsulta) {
    Logger.log("Erro: Abas 'Respostas' ou 'Consulta Conferencia' não encontradas.");
    return;
  }

  // Pega os valores da coluna A (datas) da aba Respostas
  // Ignora a primeira linha (cabeçalho)
  const datas = abaRespostas.getRange("A2:A").getValues().flat().filter(String);
  
  // Extrai os anos únicos e os ordena
  const anos = [...new Set(datas.map(data => {
    const d = new Date(data);
    return d.getFullYear();
  }).filter(ano => !isNaN(ano)))].sort((a, b) => a - b); // Ordena numericamente

  // Se não houver anos, use uma lista vazia ou um placeholder
  const anosParaValidacao = anos.length > 0 ? anos.map(String) : ["Sem Anos Disponíveis"];

  // Cria a validação com os anos
  const regraValidacaoAnos = SpreadsheetApp.newDataValidation()
    .requireValueInList(anosParaValidacao, true)
    .setAllowInvalid(false) // Garante que apenas valores da lista sejam aceitos
    .build();
  
  // Aplica a validação na célula G2 (Ano Início) da aba Consulta Conferencia
  abaConsulta.getRange("G2").setDataValidation(regraValidacaoAnos);

  // Aplica a validação na célula K2 (Ano Fim) da aba Consulta Conferencia
  abaConsulta.getRange("K2").setDataValidation(regraValidacaoAnos);
}

// Executa automaticamente ao abrir a planilha
function onOpen() {
  atualizarMenusAno(); // Renomeei a função para ser mais genérica
  // Chame a função de atualização de menus de consulta também, se necessário
  atualizarMenusConsulta(); 
  // O filtro principal deve ser chamado no onOpen também, para garantir que a tabela inicie filtrada
  filtrarDadosEPreencherDatas();
}