/**
 * Calcula a quantidade inicial do equipamento com base na menor data de "Recebimento".
 * Ignora o filtro de datas para garantir que sempre pegue o primeiro recebimento histórico.
 * @param {Array<Array<any>>} dadosRespostas Todos os dados da aba "Respostas".
 * @param {string} idSelecionado O ID filtrado.
 * @param {string} fornecedorSelecionado O fornecedor filtrado.
 * @param {string} equipamento O nome do equipamento.
 * @returns {number|string} A quantidade inicial ou "" se não houver.
 */
function calcularQtdInicial(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento) {
  let qtdInicial = "";
  
  const recebimentos = dadosRespostas.slice(1).filter(row => {
    const id = String(row[1]).trim();
    const fornecedor = String(row[6]).trim();
    const equip = String(row[7]).trim();
    const tipo = String(row[5]).trim();
    
    return id === idSelecionado &&
           fornecedor === fornecedorSelecionado &&
           equip === String(equipamento).trim() &&
           tipo === "Recebimento";
  });
  
  if (recebimentos.length > 0) {
    // Ordena por data crescente para pegar o primeiro recebimento
    recebimentos.sort((a, b) => new Date(a[0]).getTime() - new Date(b[0]).getTime());
    
    const qtd = parseFloat(recebimentos[0][8]);
    if (!isNaN(qtd) && qtd !== 0) {
      qtdInicial = qtd;
    }
  }
  
  return qtdInicial;
}

/**
 * Calcula a quantidade de devolução do equipamento na data mais recente encontrada
 * dentro do período de filtro, considerando apenas devoluções marcadas como finais.
 * @param {Array<Array<any>>} dadosRespostas Todos os dados da aba "Respostas".
 * @param {string} idSelecionado O ID filtrado.
 * @param {string} fornecedorSelecionado O fornecedor filtrado.
 * @param {string} equipamento O nome do equipamento.
 * @param {Date | null} filtroStartDate A data de início do período ajustado para filtro.
 * @param {Date | null} filtroEndDate A data de fim do período ajustado para filtro.
 * @returns {number|string} A quantidade devolvida na última data ou "" se não houver.
 */
function calcularUltimaQtdDevolvida(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate) {
  let ultimaQtdDevolvida = ""; 

  const devolucoesEquipamento = dadosRespostas.slice(1).filter(row => { 
    const dataTransacao = new Date(row[0]);
    const id = String(row[1]).trim();
    const fornecedor = String(row[6]).trim();
    const equip = String(row[7]).trim();
    const tipo = String(row[5]).trim();
    const devolucaoFinal = row[10] === true; // Verifica se é devolução final

    const condicaoGeral = id === idSelecionado &&
                         fornecedor === fornecedorSelecionado &&
                         equip === String(equipamento).trim() &&
                         tipo === "Devolução" &&
                         devolucaoFinal === true; // Apenas devoluções finais
    
    // Condição de data baseada no período ajustado
    let condicaoData = true;
    if (filtroStartDate && filtroEndDate) {
      condicaoData = (dataTransacao.getTime() >= filtroStartDate.getTime() && dataTransacao.getTime() <= filtroEndDate.getTime());
    } else if (filtroStartDate) { // Apenas data de início ajustada
        condicaoData = (dataTransacao.getTime() >= filtroStartDate.getTime());
    } else if (filtroEndDate) { // Apenas data de fim ajustada
        condicaoData = (dataTransacao.getTime() <= filtroEndDate.getTime());
    }

    return condicaoGeral && condicaoData;
  });

  if (devolucoesEquipamento.length > 0) {
    // Ordena por data decrescente para pegar a devolução mais recente
    devolucoesEquipamento.sort((a, b) => new Date(b[0]).getTime() - new Date(a[0]).getTime());
    
    const qtd = parseFloat(devolucoesEquipamento[0][8]);
    if (!isNaN(qtd) && qtd !== 0) {
      ultimaQtdDevolvida = qtd;
    }
  }
  
  return ultimaQtdDevolvida;
}

/**
 * Verifica se existe "Devolução Final" no período filtrado para um equipamento específico.
 * @param {Array<Array<any>>} dadosRespostas Todos os dados da aba "Respostas".
 * @param {string} idSelecionado O ID filtrado.
 * @param {string} fornecedorSelecionado O fornecedor filtrado.
 * @param {string} equipamento O nome do equipamento.
 * @param {Date | null} filtroStartDate A data de início do período ajustado para filtro.
 * @param {Date | null} filtroEndDate A data de fim do período ajustado para filtro.
 * @returns {boolean} True se houver devolução final no período filtrado.
 */
function temDevolucaoFinalNoPeriodo(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate) {
  const devolucoesFinal = dadosRespostas.slice(1).filter(row => {
    const dataTransacao = new Date(row[0]);
    const id = String(row[1]).trim();
    const fornecedor = String(row[6]).trim();
    const equip = String(row[7]).trim();
    const tipo = String(row[5]).trim();
    const devolucaoFinal = row[10] === true;

    const condicaoGeral = id === idSelecionado &&
                         fornecedor === fornecedorSelecionado &&
                         equip === String(equipamento).trim() &&
                         tipo === "Devolução" &&
                         devolucaoFinal === true;
    
    // Condição de data baseada no período ajustado
    let condicaoData = true;
    if (filtroStartDate && filtroEndDate) {
      condicaoData = (dataTransacao.getTime() >= filtroStartDate.getTime() && dataTransacao.getTime() <= filtroEndDate.getTime());
    } else if (filtroStartDate) {
        condicaoData = (dataTransacao.getTime() >= filtroStartDate.getTime());
    } else if (filtroEndDate) {
        condicaoData = (dataTransacao.getTime() <= filtroEndDate.getTime());
    }

    return condicaoGeral && condicaoData;
  });

  return devolucoesFinal.length > 0;
}

/**
 * Calcula a quantidade anual de um equipamento para a coluna AA, considerando
 * o último mês com "Recebimento" ou "Renovação" e subtraindo as "Devoluções" do mesmo mês.
 * IMPORTANTE: Retorna "" (vazio) se houver "Devolução Final" no período filtrado.
 * @param {Array<Array<any>>} dadosRespostas Todos os dados da aba "Respostas".
 * @param {string} idSelecionado O ID filtrado.
 * @param {string} fornecedorSelecionado O fornecedor filtrado.
 * @param {string} equipamento O nome do equipamento.
 * @param {Date | null} filtroStartDate A data de início do período ajustado para filtro.
 * @param {Date | null} filtroEndDate A data de fim do período ajustado para filtro.
 * @returns {number|string} A quantidade calculada ou "" se houver devolução final no período ou não houver dados.
 */
function calcularQtdAnual(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate) {
  // Verifica se há devolução final no período - se houver, retorna vazio
  if (temDevolucaoFinalNoPeriodo(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate)) {
    return "";
  }

  let qtdAnual = ""; 

  const transacoesRelevantes = dadosRespostas.slice(1).filter(row => {
    const dataTransacao = new Date(row[0]);
    const id = String(row[1]).trim();
    const fornecedor = String(row[6]).trim();
    const equip = String(row[7]).trim();
    const tipo = String(row[5]).trim();
    const devolucaoFinal = row[10] === true;

    // Para devoluções, considera apenas as finais
    const condicaoTipo = (tipo === "Recebimento" || tipo === "Renovação") || 
                        (tipo === "Devolução" && devolucaoFinal === true);

    const condicaoGeral = id === idSelecionado &&
                         fornecedor === fornecedorSelecionado &&
                         equip === String(equipamento).trim() &&
                         condicaoTipo;
    
    // Condição de data baseada no período ajustado
    let condicaoData = true;
    if (filtroStartDate && filtroEndDate) {
      condicaoData = (dataTransacao.getTime() >= filtroStartDate.getTime() && dataTransacao.getTime() <= filtroEndDate.getTime());
    } else if (filtroStartDate) { 
        condicaoData = (dataTransacao.getTime() >= filtroStartDate.getTime());
    } else if (filtroEndDate) { 
        condicaoData = (dataTransacao.getTime() <= filtroEndDate.getTime());
    }

    return condicaoGeral && condicaoData;
  });

  if (transacoesRelevantes.length === 0) {
    return qtdAnual; 
  }

  let ultimoMesComRecebimentoOuRenovacao = -1; // Mês 0-11
  let anoDoUltimoMes = -1;

  // Encontra o último mês com recebimento ou renovação
  transacoesRelevantes.forEach(row => {
    const dataTransacao = new Date(row[0]);
    const tipo = String(row[5]).trim();

    if ((tipo === "Recebimento" || tipo === "Renovação") && !isNaN(dataTransacao.getTime())) {
      const mesAtual = dataTransacao.getMonth();
      const anoAtual = dataTransacao.getFullYear();

      if (anoAtual > anoDoUltimoMes || (anoAtual === anoDoUltimoMes && mesAtual > ultimoMesComRecebimentoOuRenovacao)) {
        ultimoMesComRecebimentoOuRenovacao = mesAtual;
        anoDoUltimoMes = anoAtual;
      }
    }
  });

  if (ultimoMesComRecebimentoOuRenovacao === -1) {
    return qtdAnual; 
  }

  let totalRecebimentoRenovacao = 0;
  let totalDevolucao = 0;

  // Calcula totais para o mês específico
  transacoesRelevantes.forEach(row => {
    const dataTransacao = new Date(row[0]);
    const tipo = String(row[5]).trim();
    const quantity = parseFloat(row[8]);
    const devolucaoFinal = row[10] === true;

    if (!isNaN(dataTransacao.getTime()) && !isNaN(quantity) &&
        dataTransacao.getMonth() === ultimoMesComRecebimentoOuRenovacao &&
        dataTransacao.getFullYear() === anoDoUltimoMes) {
      
      if (tipo === "Recebimento" || tipo === "Renovação") {
        totalRecebimentoRenovacao += quantity;
      } else if (tipo === "Devolução" && devolucaoFinal === true) {
        // Apenas devoluções finais são consideradas
        totalDevolucao += quantity;
      }
    }
  });

  const resultado = totalRecebimentoRenovacao - totalDevolucao;
  if (!isNaN(resultado) && resultado !== 0) {
    qtdAnual = resultado;
  }

  return qtdAnual;
}

/**
 * Função auxiliar para validar se uma linha contém dados válidos
 * @param {Array<any>} row Linha dos dados
 * @returns {boolean} True se a linha é válida
 */
function validarLinha(row) {
  return row && 
         row.length >= 11 && 
         row[0] && // Data existe
         row[1] && // ID existe
         row[5] && // Tipo existe
         row[6] && // Fornecedor existe
         row[7] && // Equipamento existe
         row[8] !== undefined && row[8] !== null; // Quantidade existe
}

/**
 * Função auxiliar para obter estatísticas gerais do equipamento
 * @param {Array<Array<any>>} dadosRespostas Todos os dados da aba "Respostas"
 * @param {string} idSelecionado O ID filtrado
 * @param {string} fornecedorSelecionado O fornecedor filtrado
 * @param {string} equipamento O nome do equipamento
 * @param {Date | null} filtroStartDate Data de início do filtro
 * @param {Date | null} filtroEndDate Data de fim do filtro
 * @returns {Object} Objeto com estatísticas completas
 */
function obterEstatisticasEquipamento(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate) {
  return {
    qtdInicial: calcularQtdInicial(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento),
    ultimaQtdDevolvida: calcularUltimaQtdDevolvida(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate),
    qtdAnual: calcularQtdAnual(dadosRespostas, idSelecionado, fornecedorSelecionado, equipamento, filtroStartDate, filtroEndDate)
  };
}