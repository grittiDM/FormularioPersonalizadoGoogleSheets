/**
 * Calcula e retorna o período de busca ajustado para 12 meses
 * quando o Ano Início e Ano Fim são diferentes e há dados para ajuste.
 */
function getPeriodoDeBuscaAjustado(anoInicio, anoFim, menorDataRecebimento, maiorDataGeral) {
  let adjustedStartDate = null;
  let adjustedEndDate = null;
  const temAnoInicio = !isNaN(anoInicio);
  const temAnoFim = !isNaN(anoFim);
  
  if ((temAnoInicio && temAnoFim && anoInicio === anoFim) || (temAnoInicio && !temAnoFim) || (!temAnoInicio && temAnoFim)) {
    const anoReferencia = temAnoInicio ? anoInicio : (temAnoFim ? anoFim : new Date().getFullYear());
    adjustedStartDate = new Date(anoReferencia, 0, 1);
    adjustedEndDate = new Date(anoReferencia, 11, 31);
    return { adjustedStartDate: adjustedStartDate, adjustedEndDate: adjustedEndDate };
  }
  
  if (temAnoInicio && temAnoFim && anoInicio !== anoFim) {
    let dataReferencia = null;
    if (menorDataRecebimento instanceof Date && !isNaN(menorDataRecebimento.getTime())) {
      dataReferencia = menorDataRecebimento;
    } else if (maiorDataGeral instanceof Date && !isNaN(maiorDataGeral.getTime())) {
      dataReferencia = maiorDataGeral;
    } else {
      adjustedStartDate = new Date(anoInicio, 0, 1);
      adjustedEndDate = new Date(anoFim, 11, 31);
      return { adjustedStartDate: adjustedStartDate, adjustedEndDate: adjustedEndDate };
    }
    
    adjustedStartDate = new Date(dataReferencia.getFullYear(), dataReferencia.getMonth(), 1);
    adjustedEndDate = new Date(adjustedStartDate.getFullYear(), adjustedStartDate.getMonth() + 12, 0);
    
    const originalEndDate = new Date(anoFim, 11, 31);
    if (adjustedEndDate.getTime() > originalEndDate.getTime()) {
      adjustedEndDate = originalEndDate;
    }
    
    const originalStartDate = new Date(anoInicio, 0, 1);
    if (adjustedStartDate.getTime() < originalStartDate.getTime()) {
      adjustedStartDate = originalStartDate;
    }
    
    if (adjustedEndDate.getTime() - adjustedStartDate.getTime() > originalEndDate.getTime() - originalStartDate.getTime() ||
        adjustedStartDate.getFullYear() > originalEndDate.getFullYear() ||
        adjustedEndDate.getFullYear() < originalStartDate.getFullYear()) {
      return { adjustedStartDate: originalStartDate, adjustedEndDate: originalEndDate };
    }
    
    return { adjustedStartDate: adjustedStartDate, adjustedEndDate: adjustedEndDate };
  }
  
  return { adjustedStartDate: null, adjustedEndDate: null };
}

/**
 * Sinaliza visualmente o mês de início e o mês de fim do período ajustado na linha especificada.
 * Pinta o mês de início de VERDE e o mês de fim de VERMELHO, e escreve o ano.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} abaConsulta A aba "Consulta Conferencia".
 * @param {Date} filtroStartDate Data de início real (primeiro mês com dados).
 * @param {Date} filtroEndDate Data de fim real (último mês com dados dentro de 12 meses).
 * @param {number} linhaSinalizacao A linha onde a sinalização deve ser aplicada (ex: 5).
 */
function sinalizarPeriodoVisualmente(abaConsulta, filtroStartDate, filtroEndDate, linhaSinalizacao) {
  const corParaInicio = "#CCFFCC"; // Verde claro
  const corParaFim = "#FFCCCC";    // Vermelho claro
  const corPadrao = null;
  
  // --- LIMPEZA COMPLETA DA LINHA (C5:Z5) ---
  const rangeTotalLimpeza = abaConsulta.getRange(linhaSinalizacao, 3, 1, 24); // C5 a Z5
  rangeTotalLimpeza.clearContent(); // Limpa todo o texto/valores
  rangeTotalLimpeza.setBackground(corPadrao); // Limpa todas as cores de fundo
  
  // Se não tiver filtros de data definidos, encerra a função após a limpeza
  if (!filtroStartDate || !filtroEndDate) {
    return;
  }
  
  const colInicio = 3; // C = 3
  const paresMeses = 12;
  
  for (let i = 0; i < paresMeses; i++) {
    // O cabeçalho da planilha começa em Janeiro do ano de início (ex: 2024)
    const anoBase = filtroStartDate.getFullYear();
    const mesColuna = i; // Janeiro = 0
    let anoColuna = anoBase;
    
    if (mesColuna < filtroStartDate.getMonth()) {
      // Janeiro (0) até mês anterior ao de início estão no ano seguinte
      anoColuna++;
    }
    
    const dataColuna = new Date(anoColuna, mesColuna, 1);
    const colBase = colInicio + i * 2;
    
    // Verifica se a data da coluna corresponde ao filtroStartDate
    const isMesInicio = filtroStartDate.getFullYear() === dataColuna.getFullYear() &&
                        filtroStartDate.getMonth() === dataColuna.getMonth();
    
    // Verifica se a data da coluna corresponde ao filtroEndDate
    const isMesFim = filtroEndDate.getFullYear() === dataColuna.getFullYear() &&
                     filtroEndDate.getMonth() === dataColuna.getMonth();
    
    const celObra = abaConsulta.getRange(linhaSinalizacao, colBase);
    const celDev = abaConsulta.getRange(linhaSinalizacao, colBase + 1);
    
    if (isMesInicio) {
      celObra.setBackground(corParaInicio);
      celDev.setBackground(corParaInicio);
      celObra.setValue(filtroStartDate.getFullYear());
    }
    
    // Sinaliza o mês de fim APENAS SE FOR DIFERENTE do mês de início,
    // para evitar sobrescrever a sinalização de início em períodos de 1 mês.
    if (isMesFim && !(isMesInicio && filtroStartDate.getTime() === filtroEndDate.getTime())) {
      celObra.setBackground(corParaFim);
      celDev.setBackground(corParaFim);
      celObra.setValue(filtroEndDate.getFullYear());
    }
  }
}