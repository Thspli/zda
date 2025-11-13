const ExcelJS = require('exceljs');

class CalendarioHandler {
  async lerCalendarioPCP(caminhoArquivo) {
    console.log('📅 Lendo calendário PCP...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(caminhoArquivo);
    const sheet = workbook.worksheets[0];
    
    const calendario = {
      linhas: [],
      datas: [],
      slots: []
    };
    
    let linhaHeader = null;
    let linhaDatas = null;
    let linhaTurnos = null;
    let colunaInicio = null;
    
    sheet.eachRow((row, rowNumber) => {
      const primeirasCelulas = [];
      for (let i = 1; i <= 30; i++) {
        const valor = row.getCell(i).value;
        if (valor) primeirasCelulas.push(String(valor).toUpperCase());
      }
      
      if (primeirasCelulas.some(v => v.includes('DOMINGO') || v.includes('SEGUNDA'))) {
        linhaHeader = rowNumber;
        console.log(`   📍 Header encontrado na linha ${rowNumber}`);
        console.log(`   📋 Conteúdo: ${primeirasCelulas.slice(0, 10).join(' | ')}`);
      }
      
      if (!linhaDatas) {
        for (let i = 1; i <= 25; i++) {
          const valor = row.getCell(i).value;
          if (valor && (valor instanceof Date || String(valor).includes('2025') || String(valor).includes('/'))) {
            linhaDatas = rowNumber;
            console.log(`   📅 Datas encontradas na linha ${rowNumber}`);
            break;
          }
        }
      }
      
      if (!linhaTurnos && primeirasCelulas.some(v => v === 'A' || v === 'B' || v === 'C')) {
        linhaTurnos = rowNumber;
        console.log(`   ⏰ Turnos encontrados na linha ${rowNumber}`);
        console.log(`   📋 Conteúdo: ${primeirasCelulas.slice(0, 15).join(' | ')}`);
      }
    });
    
    if (!linhaHeader || !linhaDatas || !linhaTurnos) {
      console.error(`   ❌ Header: ${linhaHeader || 'NÃO ENCONTRADO'}`);
      console.error(`   ❌ Datas: ${linhaDatas || 'NÃO ENCONTRADO'}`);
      console.error(`   ❌ Turnos: ${linhaTurnos || 'NÃO ENCONTRADO'}`);
      throw new Error(`Não foi possível identificar o header do calendário`);
    }
    
    const rowTurnos = sheet.getRow(linhaTurnos);
    for (let col = 1; col <= 30; col++) {
      const valor = String(rowTurnos.getCell(col).value || '').toUpperCase();
      if (valor === 'A') {
        colunaInicio = col;
        console.log(`   🎯 Primeira coluna de turnos: ${col}`);
        break;
      }
    }
    
    if (!colunaInicio) {
      throw new Error('Não foi possível encontrar a coluna inicial dos turnos (A)');
    }
    
    const rowDatas = sheet.getRow(linhaDatas);
    let colAtual = colunaInicio;
    let ultimaData = null;
    
    while (colAtual <= 100) {
      const valorData = rowDatas.getCell(colAtual).value;
      const turnoAtual = String(rowTurnos.getCell(colAtual).value || '').toUpperCase();
      
      if (valorData) {
        if (valorData instanceof Date) {
          ultimaData = valorData.toLocaleDateString('pt-BR');
        } else {
          ultimaData = String(valorData);
        }
      }
      
      if ((turnoAtual === 'A' || turnoAtual === 'B' || turnoAtual === 'C') && ultimaData) {
        calendario.datas.push({
          data: ultimaData,
          turno: turnoAtual,
          coluna: colAtual
        });
      }
      
      if (!turnoAtual && colAtual > colunaInicio + 20) break;
      
      colAtual++;
    }
    
    console.log(`   📊 ${calendario.datas.length} turnos identificados`);
    
    const primeiraLinhaEquipamento = linhaTurnos + 1;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber < primeiraLinhaEquipamento) return;
      
      const nomeLinha = row.getCell(1).value;
      if (!nomeLinha || String(nomeLinha).trim() === '') return;
      
      const linha = {
        nome: String(nomeLinha).trim(),
        equipamento: row.getCell(2).value ? String(row.getCell(2).value).trim() : '',
        rowNumber
      };
      
      calendario.linhas.push(linha);
      
      calendario.datas.forEach(dataInfo => {
        const celula = row.getCell(dataInfo.coluna);
        const valor = String(celula.value || '').toUpperCase().trim();
        
        const disponivel = !valor || valor.includes('MANUTENÇÃO') || !valor.includes('SIM');
        
        if (disponivel) {
          calendario.slots.push({
            linha: linha.nome,
            equipamento: linha.equipamento,
            data: dataInfo.data,
            turno: dataInfo.turno,
            coluna: dataInfo.coluna,
            rowNumber: linha.rowNumber,
            ocupado: false,
            os: null
          });
        }
      });
    });
    
    console.log(`✅ Calendário carregado:`);
    console.log(`   ${calendario.linhas.length} linhas de produção`);
    console.log(`   ${calendario.datas.length} turnos totais`);
    console.log(`   ${calendario.slots.length} slots disponíveis para manutenção\n`);
    
    return calendario;
  }

  async exportarCalendarioPreenchido(calendarioOriginal, slotsAlocados, caminhoSaida) {
    console.log('📝 Gerando calendário preenchido...');
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(calendarioOriginal);
    const sheet = workbook.worksheets[0];
    
    let osAlocadas = 0;
    slotsAlocados.forEach(slot => {
      if (slot.os) {
        const celula = sheet.getRow(slot.rowNumber).getCell(slot.coluna);
        
        celula.value = `OS ${slot.os.ordemServico}`;
        
        let corFundo;
        if (slot.os.scoreFinal >= 80) {
          corFundo = 'FFFF6B6B';
        } else if (slot.os.scoreFinal >= 50) {
          corFundo = 'FFFFD93D';
        } else {
          corFundo = 'FF6BCF7F';
        }
        
        celula.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: corFundo }
        };
        
        celula.font = { 
          bold: true, 
          size: 8,
          color: { argb: 'FFFFFFFF' }
        };
        
        celula.alignment = { 
          vertical: 'middle', 
          horizontal: 'center', 
          wrapText: true,
          textRotation: 0
        };
        
        const column = sheet.getColumn(slot.coluna);
        if (!column.width || column.width < 12) {
          column.width = 12;
        }
        
        const row = sheet.getRow(slot.rowNumber);
        if (!row.height || row.height < 20) {
          row.height = 20;
        }
        
        osAlocadas++;
      }
    });
    
    const ultimaLinha = sheet.rowCount + 2;
    sheet.getCell(`A${ultimaLinha}`).value = 'LEGENDA:';
    sheet.getCell(`A${ultimaLinha}`).font = { bold: true };
    
    sheet.getCell(`B${ultimaLinha}`).value = 'Crítica (>80)';
    sheet.getCell(`B${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF6B6B' }
    };
    sheet.getCell(`B${ultimaLinha}`).font = { color: { argb: 'FFFFFFFF' }, bold: true };
    
    sheet.getCell(`C${ultimaLinha}`).value = 'Média (50-80)';
    sheet.getCell(`C${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFD93D' }
    };
    sheet.getCell(`C${ultimaLinha}`).font = { bold: true };
    
    sheet.getCell(`D${ultimaLinha}`).value = 'Baixa (<50)';
    sheet.getCell(`D${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF6BCF7F' }
    };
    sheet.getCell(`D${ultimaLinha}`).font = { bold: true };
    
    await this.criarAbaEscalaTecnicos(workbook, slotsAlocados);
    
    await workbook.xlsx.writeFile(caminhoSaida);
    console.log(`✅ ${osAlocadas} OS alocadas no calendário`);
    console.log(`✅ Aba de escala de técnicos criada`);
  }

  async criarAbaEscalaTecnicos(workbook, slotsAlocados) {
    console.log('👷 Duplicando calendário para escala de técnicos...');
    
    const sheetOriginal = workbook.worksheets[0];
    const sheetTecnicos = workbook.addWorksheet('Escala de Técnicos');
    
    let ultimaLinhaComDados = 0;
    
    sheetOriginal.eachRow((rowOrig, rowNum) => {
      const rowNovo = sheetTecnicos.getRow(rowNum);
      
      rowOrig.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const cellNovo = rowNovo.getCell(colNum);
        
        cellNovo.value = cell.value;
        if (cell.font) cellNovo.font = { ...cell.font };
        if (cell.fill) cellNovo.fill = { ...cell.fill };
        if (cell.alignment) cellNovo.alignment = { ...cell.alignment };
        if (cell.border) cellNovo.border = { ...cell.border };
        if (cell.numFmt) cellNovo.numFmt = cell.numFmt;
      });
      
      rowNovo.height = rowOrig.height;
      
      const valor = String(rowOrig.getCell(1).value || '').toUpperCase();
      if (valor && !valor.includes('LEGENDA')) {
        ultimaLinhaComDados = rowNum;
      }
    });
    
    sheetOriginal.columns.forEach((col, idx) => {
      sheetTecnicos.getColumn(idx + 1).width = col.width;
    });
    
    if (sheetOriginal.model && sheetOriginal.model.merges) {
      sheetOriginal.model.merges.forEach(merge => {
        sheetTecnicos.mergeCells(merge);
      });
    }
    
    let count = 0;
    slotsAlocados.forEach(slot => {
      if (slot.os && slot.tecnico) {
        const celula = sheetTecnicos.getRow(slot.rowNumber).getCell(slot.coluna);
        
        const partes = slot.tecnico.nome.split(' ');
        const nomeAbrev = partes.length > 1 
          ? `${partes[0]} ${partes[partes.length - 1]}`
          : partes[0];
        
        celula.value = nomeAbrev;
        
        const area = String(slot.tecnico.area || '').toUpperCase();
        let cor = 'FF808080';
        if (area.includes('ELETROMECANICO')) cor = 'FF4A90E2';
        else if (area.includes('MECANICO GERAL')) cor = 'FF50C878';
        else if (area.includes('EMBALAGEM')) cor = 'FFFFA500';
        else if (area.includes('SERVICO')) cor = 'FF9370DB';
        
        celula.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: cor }
        };
        
        count++;
      }
    });
    
    const linhaLegenda = ultimaLinhaComDados + 2;
    
    sheetTecnicos.getCell(`A${linhaLegenda}`).value = 'LEGENDA - ÁREAS:';
    sheetTecnicos.getCell(`A${linhaLegenda}`).font = { bold: true, size: 11 };
    
    sheetTecnicos.getCell(`B${linhaLegenda}`).value = 'Eletromecanicos';
    sheetTecnicos.getCell(`B${linhaLegenda}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4A90E2' }
    };
    sheetTecnicos.getCell(`B${linhaLegenda}`).font = { color: { argb: 'FFFFFFFF' }, bold: true };
    sheetTecnicos.getCell(`B${linhaLegenda}`).alignment = { vertical: 'middle', horizontal: 'center' };
    
    sheetTecnicos.getCell(`C${linhaLegenda}`).value = 'Mecânico Geral';
    sheetTecnicos.getCell(`C${linhaLegenda}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF50C878' }
    };
    sheetTecnicos.getCell(`C${linhaLegenda}`).font = { color: { argb: 'FFFFFFFF' }, bold: true };
    sheetTecnicos.getCell(`C${linhaLegenda}`).alignment = { vertical: 'middle', horizontal: 'center' };
    
    sheetTecnicos.getCell(`D${linhaLegenda}`).value = 'Mecânico Embalagem';
    sheetTecnicos.getCell(`D${linhaLegenda}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFA500' }
    };
    sheetTecnicos.getCell(`D${linhaLegenda}`).font = { color: { argb: 'FFFFFFFF' }, bold: true };
    sheetTecnicos.getCell(`D${linhaLegenda}`).alignment = { vertical: 'middle', horizontal: 'center' };
    
    sheetTecnicos.getCell(`E${linhaLegenda}`).value = 'Serviços Gerais';
    sheetTecnicos.getCell(`E${linhaLegenda}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF9370DB' }
    };
    sheetTecnicos.getCell(`E${linhaLegenda}`).font = { color: { argb: 'FFFFFFFF' }, bold: true };
    sheetTecnicos.getCell(`E${linhaLegenda}`).alignment = { vertical: 'middle', horizontal: 'center' };
    
    console.log(`   ✅ ${count} técnicos inseridos`);
    console.log(`   ✅ Legenda de áreas adicionada na linha ${linhaLegenda}`);
  }

  /**
   * ✅ CORRIGIDO: Data Programada não mostra mais [object Object]
   * ✅ ATUALIZADO: Agora inclui a coluna TIPO
   */
  async exportarClassificacaoOS(ordensProcessadas, caminhoSaida) {
    console.log('📊 Gerando planilha de classificação...');
    
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Classificação de OS');
    
    // ✅ COLUNA "TIPO" ADICIONADA APÓS "OS"
    sheet.columns = [
      { header: 'Prioridade', key: 'prioridade', width: 12 },
      { header: 'Score', key: 'score', width: 10 },
      { header: 'OS', key: 'ordemServico', width: 12 },
      { header: 'Tipo', key: 'tipo', width: 15 }, // ← NOVA COLUNA
      { header: 'Tag Identificada', key: 'tagIdentificada', width: 18 },
      { header: 'Descrição', key: 'descricao', width: 50 },
      { header: 'Equipamento', key: 'equipamento', width: 25 },
      { header: 'Especialidade', key: 'especialidade', width: 20 },
      { header: 'Técnico Alocado', key: 'tecnico', width: 30 },
      { header: 'Área Técnico', key: 'areaTecnico', width: 20 },
      { header: 'Local', key: 'local', width: 25 },
      { header: 'Classe', key: 'classe', width: 10 },
      { header: 'Método Match', key: 'metodoMatch', width: 25 },
      { header: 'Status Alocação', key: 'statusAlocacao', width: 20 },
      { header: 'Data Programada', key: 'dataProgramada', width: 20 },
      { header: 'Motivo Pendência', key: 'motivoNaoAlocacao', width: 25 }
    ];
    
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F4E78' }
    };
    sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
    
    ordensProcessadas.forEach((ordem, index) => {
      const tecnicoNome = ordem.tecnicoAlocado ? ordem.tecnicoAlocado.nome : '-';
      const tecnicoArea = ordem.tecnicoAlocado ? ordem.tecnicoAlocado.area : '-';
      
      // ✅ CORREÇÃO CRÍTICA: Converter dataProgramada para string sempre
      let dataProgramadaStr = '-';
      if (ordem.dataProgramada) {
        if (typeof ordem.dataProgramada === 'string') {
          dataProgramadaStr = ordem.dataProgramada;
        } else if (ordem.dataProgramada instanceof Date) {
          dataProgramadaStr = this.formatarData(ordem.dataProgramada);
        } else if (typeof ordem.dataProgramada === 'object') {
          // Se for objeto, tentar extrair informações úteis
          if (ordem.slotAlocado) {
            dataProgramadaStr = `${ordem.slotAlocado.data} - Turno ${ordem.slotAlocado.turno}`;
          } else {
            dataProgramadaStr = JSON.stringify(ordem.dataProgramada);
          }
        }
      }
      
      // ✅ VALOR DA COLUNA TIPO INCLUÍDO
      const row = sheet.addRow({
        prioridade: index + 1,
        score: ordem.scoreFinal.toFixed(1),
        ordemServico: ordem.ordemServico,
        tipo: ordem.tipo || 'N/A', // ← NOVA LINHA
        tagIdentificada: ordem.tagIdentificada || 'N/A',
        descricao: ordem.descricao || '',
        equipamento: ordem.equipamento || '',
        especialidade: ordem.especialidadeNecessaria || '-',
        tecnico: tecnicoNome,
        areaTecnico: tecnicoArea,
        local: ordem.local || '',
        classe: ordem.classe,
        metodoMatch: ordem.metodoMatch || 'N/A',
        statusAlocacao: ordem.alocada ? 'Programada' : 'Pendente',
        dataProgramada: dataProgramadaStr, // ✅ USANDO STRING SEMPRE
        motivoNaoAlocacao: ordem.motivoNaoAlocacao || '-'
      });
      
      let corFundo;
      if (ordem.scoreFinal >= 80) {
        corFundo = 'FFFF6B6B';
      } else if (ordem.scoreFinal >= 50) {
        corFundo = 'FFFFD93D';
      } else {
        corFundo = 'FF6BCF7F';
      }
      
      row.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: corFundo }
      };
      
      row.alignment = { vertical: 'middle' };
      
      if (ordem.tagIdentificada) {
        row.getCell('tagIdentificada').font = { bold: true };
      }
    });
    
    const statsSheet = workbook.addWorksheet('Estatísticas');
    statsSheet.columns = [
      { header: 'Métrica', key: 'metrica', width: 35 },
      { header: 'Valor', key: 'valor', width: 20 }
    ];
    
    const total = ordensProcessadas.length;
    const alocadas = ordensProcessadas.filter(o => o.alocada).length;
    const pendentes = total - alocadas;
    const criticas = ordensProcessadas.filter(o => o.scoreFinal >= 80).length;
    const medias = ordensProcessadas.filter(o => o.scoreFinal >= 50 && o.scoreFinal < 80).length;
    const baixas = ordensProcessadas.filter(o => o.scoreFinal < 50).length;
    
    const comTagIdentificada = ordensProcessadas.filter(o => o.tagIdentificada).length;
    const matchCampoBem = ordensProcessadas.filter(o => o.metodoMatch === 'Tag campo BEM').length;
    const matchExtraido = ordensProcessadas.filter(o => o.metodoMatch === 'Tag extraída da descrição').length;
    const matchNome = ordensProcessadas.filter(o => o.metodoMatch === 'Match por nome').length;
    
    const comTecnicoAlocado = ordensProcessadas.filter(o => o.tecnicoAlocado).length;
    const semTecnicoDisponivel = ordensProcessadas.filter(o => 
      o.motivoNaoAlocacao === 'Sem técnico disponível no turno'
    ).length;
    
    // ✅ ESTATÍSTICAS DE TIPO ADICIONADAS
    const preventivas = ordensProcessadas.filter(o => 
      String(o.tipo || '').toUpperCase().includes('PREVENTIVA')
    ).length;
    const corretivas = ordensProcessadas.filter(o => 
      String(o.tipo || '').toUpperCase().includes('CORRETIVA')
    ).length;
    
    statsSheet.addRow({ metrica: '=== ALOCAÇÃO ===', valor: '' });
    statsSheet.addRow({ metrica: 'Total de OS', valor: total });
    statsSheet.addRow({ metrica: 'OS Programadas', valor: alocadas });
    statsSheet.addRow({ metrica: 'OS Pendentes', valor: pendentes });
    statsSheet.addRow({ metrica: '', valor: '' });
    
    // ✅ SEÇÃO DE TIPOS
    statsSheet.addRow({ metrica: '=== TIPOS DE MANUTENÇÃO ===', valor: '' });
    statsSheet.addRow({ metrica: 'Preventivas', valor: preventivas });
    statsSheet.addRow({ metrica: 'Corretivas', valor: corretivas });
    statsSheet.addRow({ metrica: '', valor: '' });
    
    statsSheet.addRow({ metrica: '=== PRIORIDADE ===', valor: '' });
    statsSheet.addRow({ metrica: 'Prioridade Crítica (>80)', valor: criticas });
    statsSheet.addRow({ metrica: 'Prioridade Média (50-80)', valor: medias });
    statsSheet.addRow({ metrica: 'Prioridade Baixa (<50)', valor: baixas });
    statsSheet.addRow({ metrica: '', valor: '' });
    
    statsSheet.addRow({ metrica: '=== IDENTIFICAÇÃO DE EQUIPAMENTOS ===', valor: '' });
    statsSheet.addRow({ metrica: 'OS com Tag identificada', valor: comTagIdentificada });
    statsSheet.addRow({ metrica: 'Taxa de identificação', valor: `${(comTagIdentificada/total*100).toFixed(1)}%` });
    statsSheet.addRow({ metrica: 'Match por campo BEM', valor: matchCampoBem });
    statsSheet.addRow({ metrica: 'Match por extração da descrição', valor: matchExtraido });
    statsSheet.addRow({ metrica: 'Match por nome/similaridade', valor: matchNome });
    statsSheet.addRow({ metrica: '', valor: '' });
    
    statsSheet.addRow({ metrica: '=== ALOCAÇÃO DE TÉCNICOS ===', valor: '' });
    statsSheet.addRow({ metrica: 'OS com técnico alocado', valor: comTecnicoAlocado });
    statsSheet.addRow({ metrica: 'Taxa de alocação de técnicos', valor: `${(comTecnicoAlocado/total*100).toFixed(1)}%` });
    statsSheet.addRow({ metrica: 'OS sem técnico disponível', valor: semTecnicoDisponivel });
    
    statsSheet.getRow(1).font = { bold: true };
    statsSheet.getRow(6).font = { bold: true };
    statsSheet.getRow(10).font = { bold: true };
    statsSheet.getRow(15).font = { bold: true };
    statsSheet.getRow(22).font = { bold: true };
    
    await workbook.xlsx.writeFile(caminhoSaida);
    console.log(`✅ Classificação gerada:`);
    console.log(`   ${alocadas}/${total} OS programadas`);
    console.log(`   ${preventivas} Preventivas | ${corretivas} Corretivas`);
    console.log(`   ${comTagIdentificada}/${total} TAGs identificadas (${(comTagIdentificada/total*100).toFixed(1)}%)`);
    
    if (comTecnicoAlocado > 0) {
      console.log(`   ${comTecnicoAlocado}/${total} Técnicos alocados (${(comTecnicoAlocado/total*100).toFixed(1)}%)`);
    }
  }

  /**
   * ✅ Função auxiliar para formatar datas
   */
  formatarData(data) {
    if (!data) return '';
    try {
      const dia = String(data.getDate()).padStart(2, '0');
      const mes = String(data.getMonth() + 1).padStart(2, '0');
      const ano = data.getFullYear();
      return `${dia}/${mes}/${ano}`;
    } catch (error) {
      return String(data);
    }
  }
}

module.exports = new CalendarioHandler();