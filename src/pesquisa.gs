function filtrarDadosEPreencherDatas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRespostas = ss.getSheetByName("Respostas");
  const abaConsulta = ss.getSheetByName("Consulta Conferencia");

  if (!abaRespostas || !abaConsulta) {
    Logger.log("Erro: Abas 'Respostas' ou 'Consulta Conferencia' não encontradas.");
    return;
  }

  // Definir as linhas fixas da tabela conforme o layout
  const LINHA_CABECALHO_TABELA = 6;
  const LINHA_SINALIZACAO = 5; // <--- NOVA LINHA PARA SINALIZAÇÃO
  const LINHA_INICIO_DADOS_EQUIPAMENTOS = 8;
  const LINHA_RODAPE_TABELA_ORIGINAL = 10; 

  // --- LEITURA DOS FILTROS ---
  const idSelecionado = String(abaConsulta.getRange("C1").getValue()).trim();
  const fornecedorSelecionado = String(abaConsulta.getRange("C2").getValue()).trim();
  let anoInicio = parseInt(abaConsulta.getRange("G2").getValue()); 
  let anoFim = parseInt(abaConsulta.getRange("K2").getValue());   

  // Ajuste: Se Ano Fim (K2) não for preenchido, limitar a busca ao Ano Início (G2)
  if (isNaN(anoFim) && !isNaN(anoInicio)) {
    anoFim = anoInicio;
  } else if (isNaN(anoInicio) && !isNaN(anoFim)) {
    anoInicio = anoFim; 
  }


  const rangeMenorDataSaida = abaConsulta.getRange("D3");
  const rangeMaiorDataSaida = abaConsulta.getRange("D4");
  const rangeDataConsultaSaida = abaConsulta.getRange("G1");

  const dadosRespostas = abaRespostas.getDataRange().getValues();

  // Se não há dados na aba Respostas (além do cabeçalho), limpa a tabela e sai
  if (dadosRespostas.length <= 1) { 
    Logger.log("Aba 'Respostas' não possui dados para filtrar. Limpando tabela.");
    limparTabelaConsulta(abaConsulta, LINHA_INICIO_DADOS_EQUIPAMENTOS, LINHA_RODAPE_TABELA_ORIGINAL, LINHA_SINALIZACAO); // <--- ATUALIZADO
    rangeMenorDataSaida.clearContent();
    rangeMaiorDataSaida.clearContent();
    rangeDataConsultaSaida.clearContent();
    return;
  }

  // --- BUSCA INICIAL DE DATAS PARA AJUSTE DO PERÍODO ---
  let menorDataRecebimentoGlobal = null;
  let maiorDataGeralGlobal = null;

  dadosRespostas.slice(1).forEach(row => {
      const dataAtual = new Date(row[0]);
      const tipoTransacao = String(row[5]).trim();

      if (!isNaN(dataAtual.getTime())) {
          if (maiorDataGeralGlobal === null || dataAtual.getTime() > maiorDataGeralGlobal.getTime()) {
              maiorDataGeralGlobal = dataAtual;
          }
          if (tipoTransacao === "Recebimento") {
              if (menorDataRecebimentoGlobal === null || dataAtual.getTime() < menorDataRecebimentoGlobal.getTime()) {
                  menorDataRecebimentoGlobal = dataAtual;
              }
          }
      }
  });

  // --- AJUSTE DE PERÍODO (CHAMADA PARA FUNÇÃO DO PERIODO.GS) ---
  const periodoAjustado = getPeriodoDeBuscaAjustado(anoInicio, anoFim, menorDataRecebimentoGlobal, maiorDataGeralGlobal);

  const filtroStartDate = periodoAjustado.adjustedStartDate;
  const filtroEndDate = periodoAjustado.adjustedEndDate;

  // --- SINALIZAÇÃO VISUAL (CHAMADA PARA FUNÇÃO DO PERIODO.GS) ---
  sinalizarPeriodoVisualmente(abaConsulta, filtroStartDate, filtroEndDate, LINHA_SINALIZACAO); // <--- ATUALIZADO


  // --- FILTRAGEM PRINCIPAL DOS DADOS (AGORA USANDO O PERÍODO AJUSTADO) ---
  const dadosFiltradosParaBuscaPrincipal = dadosRespostas.slice(1).filter(row => { 
    const dataRegistro = new Date(row[0]);
    const id = String(row[1]).trim();
    const fornecedor = String(row[6]).trim();

    const condicaoID = idSelecionado ? (id === idSelecionado) : true;
    const condicaoFornecedor = fornecedorSelecionado ? (fornecedor === fornecedorSelecionado) : true;
    
    // Condição de data baseada no período ajustado
    let condicaoData = true;
    if (filtroStartDate && filtroEndDate) {
      condicaoData = (dataRegistro.getTime() >= filtroStartDate.getTime() && dataRegistro.getTime() <= filtroEndDate.getTime());
    } else if (filtroStartDate) { 
        condicaoData = (dataRegistro.getTime() >= filtroStartDate.getTime());
    } else if (filtroEndDate) { 
        condicaoData = (dataRegistro.getTime() <= filtroEndDate.getTime());
    } else if (!isNaN(anoInicio) || !isNaN(anoFim)) { // Fallback se não houve ajuste, mas anos foram preenchidos
        const anoDaData = !isNaN(dataRegistro.getFullYear()) ? dataRegistro.getFullYear() : null;
        if (!isNaN(anoInicio) && !isNaN(anoFim)) {
            condicaoData = (anoDaData >= anoInicio && anoDaData <= anoFim);
        } else if (!isNaN(anoInicio)) {
            condicaoData = (anoDaData === anoInicio);
        } else if (!isNaN(anoFim)) {
            condicaoData = (anoDaData <= anoFim);
        }
    }

    return condicaoID && condicaoFornecedor && condicaoData;
  });

  // Preencher D3 (Menor Data Saída) e D4 (Maior Data Saída) com base nos dados FILTRADOS
  let menorDataRecebimentoFiltrada = null;
  let maiorDataDevolucaoFiltrada = null;
  let maiorDataGeralFiltrada = null;

  if (dadosFiltradosParaBuscaPrincipal.length > 0) {
    dadosFiltradosParaBuscaPrincipal.forEach(row => {
      const dataAtual = new Date(row[0]);
      const tipoTransacao = String(row[5]).trim();

      if (!isNaN(dataAtual.getTime())) {
        if (maiorDataGeralFiltrada === null || dataAtual.getTime() > maiorDataGeralFiltrada.getTime()) {
          maiorDataGeralFiltrada = dataAtual;
        }
        if (tipoTransacao === "Recebimento") {
          if (menorDataRecebimentoFiltrada === null || dataAtual.getTime() < menorDataRecebimentoFiltrada.getTime()) {
            menorDataRecebimentoFiltrada = dataAtual;
          }
        }
        if (tipoTransacao === "Devolução") {
          if (maiorDataDevolucaoFiltrada === null || dataAtual.getTime() > maiorDataDevolucaoFiltrada.getTime()) {
            maiorDataDevolucaoFiltrada = dataAtual;
          }
        }
      }
    });
  }

  // Preencher D3 e D4 com base nas datas dos dados REALMENTE filtrados
  if (menorDataRecebimentoFiltrada) {
    rangeMenorDataSaida.setValue(Utilities.formatDate(menorDataRecebimentoFiltrada, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"));
  } else {
    rangeMenorDataSaida.clearContent();
  }

  if (maiorDataDevolucaoFiltrada || maiorDataGeralFiltrada) {
    rangeMaiorDataSaida.setValue(Utilities.formatDate(maiorDataDevolucaoFiltrada || maiorDataGeralFiltrada, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"));
  } else {
    rangeMaiorDataSaida.clearContent();
  }

  rangeDataConsultaSaida.setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss"));

  // --- PARTE DE GERENCIAMENTO DE LINHAS E PREENCHIMENTO DE EQUIPAMENTOS ---

  let equipamentosUnicos = [];
  if (idSelecionado && fornecedorSelecionado && (!isNaN(anoInicio) || !isNaN(anoFim))) {
    equipamentosUnicos = [...new Set(
      dadosFiltradosParaBuscaPrincipal
        .filter(row => String(row[1]).trim() === idSelecionado && 
                        String(row[6]).trim() === fornecedorSelecionado) 
        .map(row => row[7]) 
        .filter(val => val)
    )].sort();
  }

  const numEquipamentos = equipamentosUnicos.length;
  const linhasDeDadosRequeridas = Math.max(numEquipamentos, 2); 

  let linhaRodapeDetectadaRange = abaConsulta.getRange("A:A").createTextFinder("TOTAL").matchEntireCell(true).findNext();
  let linhaRodapeDetectada = linhaRodapeDetectadaRange ? linhaRodapeDetectadaRange.getRow() : abaConsulta.getLastRow() + 1;

  const NUM_COLUNAS_DA_TABELA = 28; 

  const rangeRodapeFormulas = abaConsulta.getRange(linhaRodapeDetectada, 2, 1, NUM_COLUNAS_DA_TABELA - 1);
  rangeRodapeFormulas.clearContent();

  const linhaFimDaAreaDeDadosAtual = linhaRodapeDetectada - 1;
  if (linhaFimDaAreaDeDadosAtual >= LINHA_INICIO_DADOS_EQUIPAMENTOS) {
    const rangeParaLimparDados = abaConsulta.getRange(
      LINHA_INICIO_DADOS_EQUIPAMENTOS,
      1,
      linhaFimDaAreaDeDadosAtual - LINHA_INICIO_DADOS_EQUIPAMENTOS + 1,
      NUM_COLUNAS_DA_TABELA
    );
    rangeParaLimparDados.clearContent();
    rangeParaLimparDados.clearDataValidations();
    rangeParaLimparDados.clearNote();
  }

  const linhasAtualmenteNaAreaDeDados = linhaRodapeDetectada - LINHA_INICIO_DADOS_EQUIPAMENTOS;

  if (linhasAtualmenteNaAreaDeDados > linhasDeDadosRequeridas) {
    const linhasParaRemover = linhasAtualmenteNaAreaDeDados - linhasDeDadosRequeridas;
    const linhaDeRemocaoInicio = LINHA_INICIO_DADOS_EQUIPAMENTOS + linhasDeDadosRequeridas;
    if (linhasParaRemover > 0 && linhaDeRemocaoInicio < linhaRodapeDetectada) {
      abaConsulta.deleteRows(linhaDeRemocaoInicio, linhasParaRemover);
      // Re-detecta a linha do rodapé após a exclusão
      linhaRodapeDetectada = abaConsulta.getRange("A:A").createTextFinder("TOTAL").matchEntireCell(true).findNext().getRow();
    }
  } else if (linhasDeDadosRequeridas > linhasAtualmenteNaAreaDeDados) {
    const linhasParaInserir = linhasDeDadosRequeridas - linhasAtualmenteNaAreaDeDados;
    if (linhasParaInserir > 0) {
      abaConsulta.insertRowsBefore(linhaRodapeDetectada, linhasParaInserir);
      // Re-detecta a linha do rodapé após a inserção
      linhaRodapeDetectada = abaConsulta.getRange("A:A").createTextFinder("TOTAL").matchEntireCell(true).findNext().getRow();
    }
  }

  const linhaFinalParaSoma = linhaRodapeDetectada - 1;

  if (numEquipamentos > 0 && idSelecionado && fornecedorSelecionado && (!isNaN(anoInicio) || !isNaN(anoFim))) {
    const dadosParaTabela = [];
    const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

    equipamentosUnicos.forEach(equipamento => {
      let qtdInicial = ""; 
      const quantidadesMensais = {}; 
      MESES.forEach(mes => {
          quantidadesMensais[mes] = { obra: 0, dev: 0 }; 
      });

      const primeiroRecebimentoNoPeriodo = dadosFiltradosParaBuscaPrincipal.find(row => {
          const tipo = String(row[5]).trim();
          const equip = String(row[7]).trim();
          const id = String(row[1]).trim();
          const fornecedor = String(row[6]).trim();
          
          return id === idSelecionado &&
                 fornecedor === fornecedorSelecionado &&
                 equip === String(equipamento).trim() &&
                 tipo === "Recebimento";
      });

      if (primeiroRecebimentoNoPeriodo) {
          const valQtdInicial = parseFloat(primeiroRecebimentoNoPeriodo[8]);
          if (!isNaN(valQtdInicial) && valQtdInicial !== 0) {
              qtdInicial = valQtdInicial;
          }
      }
      
      dadosFiltradosParaBuscaPrincipal.forEach(row => {
          const dataTransacao = new Date(row[0]);
          const tipo = String(row[5]).trim();
          const equip = String(row[7]).trim();
          const quantity = parseFloat(row[8]);

          if (!isNaN(dataTransacao.getTime()) && !isNaN(quantity) && equip === String(equipamento).trim()) {
              const mesIndex = dataTransacao.getMonth(); 
              const mesNome = MESES[mesIndex];

              // Verificar se a data está dentro do período ajustado para a contagem mensal
              if (filtroStartDate && filtroEndDate) {
                  if (dataTransacao.getTime() >= filtroStartDate.getTime() && dataTransacao.getTime() <= filtroEndDate.getTime()) {
                      if (quantidadesMensais[mesNome]) {
                          if (tipo === "Recebimento" || tipo === "Renovação") {
                              quantidadesMensais[mesNome].obra += quantity;
                          } else if (tipo === "Devolução") {
                              quantidadesMensais[mesNome].dev += quantity;
                          }
                      }
                  }
              } else { 
                 if (quantidadesMensais[mesNome]) {
                          if (tipo === "Recebimento" || tipo === "Renovação") {
                              quantidadesMensais[mesNome].obra += quantity;
                          } else if (tipo === "Devolução") {
                              quantidadesMensais[mesNome].dev += quantity;
                          }
                      }
              }
          }
      });

      const linhaTabela = Array(NUM_COLUNAS_DA_TABELA).fill("");
      linhaTabela[0] = equipamento; 
      linhaTabela[1] = qtdInicial; 

      let currentColIndex = 2; 
      MESES.forEach(mes => {
          linhaTabela[currentColIndex] = quantidadesMensais[mes].obra === 0 ? "" : quantidadesMensais[mes].obra; 
          linhaTabela[currentColIndex + 1] = quantidadesMensais[mes].dev === 0 ? "" : quantidadesMensais[mes].dev; 
          currentColIndex += 2; 
      });

      // --- CHAMADA PARA AS NOVAS FUNÇÕES (PASSANDO AS DATAS AJUSTADAS) ---
      // Coluna AA (índice 26)
      linhaTabela[26] = calcularQtdAnual(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate);

      // Coluna AB (índice 27)
      linhaTabela[27] = calcularUltimaQtdDevolvida(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate);

      dadosParaTabela.push(linhaTabela);
    });

    if (dadosParaTabela.length === 0) {
        dadosParaTabela.push(Array(NUM_COLUNAS_DA_TABELA).fill("")); 
        dadosParaTabela.push(Array(NUM_COLUNAS_DA_TABELA).fill("")); 
    } else if (dadosParaTabela.length === 1) {
        dadosParaTabela.push(Array(NUM_COLUNAS_DA_TABELA).fill("")); 
    }


    abaConsulta.getRange(
      LINHA_INICIO_DADOS_EQUIPAMENTOS,
      1,
      dadosParaTabela.length,
      dadosParaTabela[0].length
    ).setValues(dadosParaTabela);

    const numLinhasPreenchidas = dadosParaTabela.length;
    for (let i = 0; i < numLinhasPreenchidas; i++) {
        const linhaOrigemFormatacao = LINHA_INICIO_DADOS_EQUIPAMENTOS + (i % 2);
        const linhaDestino = LINHA_INICIO_DADOS_EQUIPAMENTOS + i;
        abaConsulta.getRange(linhaOrigemFormatacao, 1, 1, NUM_COLUNAS_DA_TABELA)
            .copyTo(abaConsulta.getRange(linhaDestino, 1, 1, NUM_COLUNAS_DA_TABELA), { formatOnly: true });
    }

    const colunaQtdInicial = 2; 
    abaConsulta.getRange(linhaRodapeDetectada, colunaQtdInicial)
        .setFormulaR1C1(`=SUM(R${LINHA_INICIO_DADOS_EQUIPAMENTOS}C:R${linhaFinalParaSoma}C)`);

    // As colunas de soma mensal (C até Z)
    for (let col = 3; col <= 26; col++) { 
      abaConsulta.getRange(linhaRodapeDetectada, col)
          .setFormulaR1C1(`=SUM(R${LINHA_INICIO_DADOS_EQUIPAMENTOS}C:R${linhaFinalParaSoma}C)`);
    }

    // --- REINTRODUZIR FÓRMULAS SUM PARA AA E AB NO RODAPÉ ---
    // Coluna AA (índice 26, coluna 27)
    abaConsulta.getRange(linhaRodapeDetectada, 27).setFormulaR1C1(`=SUM(R${LINHA_INICIO_DADOS_EQUIPAMENTOS}C:R${linhaFinalParaSoma}C)`);
    // Coluna AB (índice 27, coluna 28)
    abaConsulta.getRange(linhaRodapeDetectada, 28).setFormulaR1C1(`=SUM(R${LINHA_INICIO_DADOS_EQUIPAMENTOS}C:R${linhaFinalParaSoma}C)`);
    

  } else {
    limparTabelaConsulta(abaConsulta, LINHA_INICIO_DADOS_EQUIPAMENTOS, linhaRodapeDetectada, LINHA_SINALIZACAO); // <--- ATUALIZADO
  }
}


/**
 * Função auxiliar para limpar a área de dados da tabela e aplicar formatação padrão.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet A aba onde a tabela está.
 * @param {number} linhaInicioDados A linha onde os dados da tabela começam.
 * @param {number} linhaRodape A linha onde o rodapé "TOTAL" está.
 * @param {number} linhaSinalizacao A linha onde a sinalização dos anos está para limpeza de formatação.
 */
function limparTabelaConsulta(sheet, linhaInicioDados, linhaRodape, linhaSinalizacao) { // <--- NOVO PARAMETRO
    const NUM_COLUNAS_DA_TABELA = 28; 
    const LINHA_RODAPE_TABELA_ORIGINAL = 10; 
    const linhaFimDaAreaDeDadosAtual = linhaRodape - 1;

    if (linhaFimDaAreaDeDadosAtual >= linhaInicioDados) {
        const rangeParaLimparDados = sheet.getRange(
            linhaInicioDados,
            1,
            linhaFimDaAreaDeDadosAtual - linhaInicioDados + 1,
            NUM_COLUNAS_DA_TABELA
        );
        rangeParaLimparDados.clearContent();
        rangeParaLimparDados.clearDataValidations();
        rangeParaLimparDados.clearNote();
    }

    const linhasDeDadosRequeridas = 2; 
    for (let i = 0; i < linhasDeDadosRequeridas; i++) {
        const linhaOrigemFormatacao = linhaInicioDados + (i % 2); 
        const linhaDestino = linhaInicioDados + i;
        sheet.getRange(linhaOrigemFormatacao, 1, 1, NUM_COLUNAS_DA_TABELA)
            .copyTo(sheet.getRange(linhaDestino, 1, 1, NUM_COLUNAS_DA_TABELA), { formatOnly: true });
    }
    
    let linhaRodapeAtualizadaRange = sheet.getRange("A:A").createTextFinder("TOTAL").matchEntireCell(true).findNext();
    let linhaRodapeAtualizada;
    if (linhaRodapeAtualizadaRange) {
        linhaRodapeAtualizada = linhaRodapeAtualizadaRange.getRow();
        sheet.getRange(LINHA_RODAPE_TABELA_ORIGINAL, 1, 1, NUM_COLUNAS_DA_TABELA)
            .copyTo(sheet.getRange(linhaRodapeAtualizada, 1, 1, NUM_COLUNAS_DA_TABELA), { formatOnly: true });
        
        // Limpa conteúdo das somas no rodapé ao limpar a tabela
        sheet.getRange(linhaRodapeAtualizada, 27).clearContent(); // Coluna AA
        sheet.getRange(linhaRodapeAtualizada, 28).clearContent(); // Coluna AB
    }

    // Chama a sinalização visual para remover as cores e anos do cabeçalho de meses
    sinalizarPeriodoVisualmente(sheet, null, null, linhaSinalizacao); // <--- ATUALIZADO
}