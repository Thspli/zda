const ExcelJS = require('exceljs');
const path = require('path');

class ExcelHandler {
  /**
   * Lê planilha de ordens de serviço do TOTVS
   */
  async lerOrdensServico(caminhoArquivo) {
    console.log('📖 Lendo ordens de serviço...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(caminhoArquivo);
    const sheet = workbook.worksheets[0];
    
    const ordens = [];
    const headers = {};
    
    // Encontrar linha de headers
    let headerRowNumber = 1;
    for (let i = 1; i <= 10; i++) {
      const row = sheet.getRow(i);
      let foundOrderColumn = false;
      
      row.eachCell((cell) => {
        const value = String(cell.value || '').trim();
        if (value.includes('Ordem') || value === 'Ordem Serv.') {
          foundOrderColumn = true;
        }
      });
      
      if (foundOrderColumn) {
        headerRowNumber = i;
        break;
      }
    }
    
    console.log(`📋 Linha de headers: ${headerRowNumber}`);
    
    // Mapear headers
    const headerRow = sheet.getRow(headerRowNumber);
    headerRow.eachCell((cell, colNumber) => {
      if (cell.value) {
        const headerName = String(cell.value).trim();
        headers[headerName] = colNumber;
      }
    });
    
    console.log('📋 Headers encontrados:', Object.keys(headers));
    
    // Processar linhas de dados
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber <= headerRowNumber) return;
      if (!row.hasValues) return;
      
      const ordemServicoValue = this.getCellValue(row, headers['Ordem Serv.']);
      if (!ordemServicoValue) return;
      
      const ordem = {
        ordemServico: ordemServicoValue,
        bem: this.getCellValue(row, headers['Bem']),
        nomeBem: this.getCellValue(row, headers['Nome do Bem']),
        servico: this.getCellValue(row, headers['Servico']),
        nomeServico: this.getCellValue(row, headers['Nome Servico']),
        prevInicio: this.parseData(this.getCellValue(row, headers['Prev. Inicio'])),
        pInicioMan: this.getCellValue(row, headers['P. In. Man.']),
        rInicioMan: this.getCellValue(row, headers['R. In. Man.']),
        status: this.getCellValue(row, headers['Status da OS']),
        descricao: this.getCellValue(row, headers['Descricao'])
      };
      
      if (ordem.status === 'PENDEN' || ordem.status === 'LIBERA') {
        ordens.push(ordem);
      }
    });
    
    console.log(`✅ ${ordens.length} ordens de serviço carregadas`);
    return ordens;
  }

  /**
   * Versão simplificada para formato de 3 colunas (ID, Descrição, Tipo)
   * ✅ ATUALIZADO: Agora lê a coluna TIPO (Preventiva/Corretiva)
   */
  async lerOrdensServicoSimplificada(caminhoArquivo) {
    console.log('📖 Lendo ordens de serviço (formato simplificado)...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(caminhoArquivo);
    const sheet = workbook.worksheets[0];
    
    const ordens = [];
    
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Pular header
      
      const os = this.getCellValue(row, 1);
      const descricao = this.getCellValue(row, 2);
      const tipo = this.getCellValue(row, 3); // ← NOVA LINHA: Coluna 3 = Tipo
      
      if (os && descricao) {
        ordens.push({
          ordemServico: String(os).trim(),
          descricao: String(descricao).trim(),
          tipo: tipo ? String(tipo).trim() : 'N/A', // ← NOVA LINHA: Captura Preventiva/Corretiva
          bem: null,
          nomeBem: null,
          servico: null,
          nomeServico: null,
          prevInicio: new Date('2025-09-15'),
          status: 'PENDEN'
        });
      }
    });
    
    console.log(`✅ ${ordens.length} ordens de serviço carregadas (formato simplificado)`);
    
    // Estatística de tipos
    const preventivas = ordens.filter(o => String(o.tipo).toUpperCase().includes('PREVENTIVA')).length;
    const corretivas = ordens.filter(o => String(o.tipo).toUpperCase().includes('CORRETIVA')).length;
    console.log(`   📊 ${preventivas} Preventivas | ${corretivas} Corretivas\n`);
    
    return ordens;
  }

  /**
   * NOVO: Extrai TAG da descrição e vincula com equipamentos
   * Prioridade: bem → extração de TAG da descrição → match por nome
   */
  enriquecerOSComEquipamentos(ordens, ativos) {
    console.log('🔗 Vinculando OS com equipamentos (por Tag)...');
    
    let matchExato = 0;
    let matchExtraidoDescricao = 0;
    let matchPorNome = 0;
    let semMatch = 0;
    
    // Criar índice reverso para busca rápida
    const indiceDescricoes = this.criarIndiceDescricoes(ativos);
    
    ordens.forEach(ordem => {
      // Inicializar valores padrão
      ordem.equipamento = null;
      ordem.local = null;
      ordem.classe = 'C';
      ordem.metodoMatch = 'Nenhum';
      ordem.tagIdentificada = null;
      
      // PRIORIDADE 1: Tag exata no campo 'bem'
      if (ordem.bem) {
        const tagBuscada = String(ordem.bem).trim();
        
        if (ativos[tagBuscada]) {
          const ativo = ativos[tagBuscada];
          ordem.equipamento = ativo.descricao;
          ordem.local = ativo.local;
          ordem.classe = ativo.classe;
          ordem.tagIdentificada = ativo.tag;
          ordem.metodoMatch = 'Tag campo BEM';
          matchExato++;
          return;
        }
      }
      
      // PRIORIDADE 2: Extrair TAG da descrição
      if (ordem.descricao) {
        const tagExtraida = this.extrairTagDaDescricao(ordem.descricao, ativos);
        
        if (tagExtraida && ativos[tagExtraida]) {
          const ativo = ativos[tagExtraida];
          ordem.equipamento = ativo.descricao;
          ordem.local = ativo.local;
          ordem.classe = ativo.classe;
          ordem.tagIdentificada = ativo.tag;
          ordem.metodoMatch = 'Tag extraída da descrição';
          matchExtraidoDescricao++;
          return;
        }
      }
      
      // PRIORIDADE 3: Match por nome/descrição
      if (ordem.descricao) {
        const resultadoBusca = this.buscarPorSimilaridade(ordem.descricao, indiceDescricoes, ativos);
        
        if (resultadoBusca) {
          ordem.equipamento = resultadoBusca.descricao;
          ordem.local = resultadoBusca.local;
          ordem.classe = resultadoBusca.classe;
          ordem.tagIdentificada = resultadoBusca.tag;
          ordem.metodoMatch = 'Match por nome';
          matchPorNome++;
          return;
        }
      }
      
      // PRIORIDADE 4: Usar nomeBem se existir
      if (!ordem.equipamento && ordem.nomeBem) {
        ordem.equipamento = ordem.nomeBem;
        ordem.metodoMatch = 'Nome do Bem (sem classe)';
      }
      
      semMatch++;
    });
    
    console.log(`✅ Vinculação concluída:`);
    console.log(`   ${matchExato} matches por Tag no campo BEM`);
    console.log(`   ${matchExtraidoDescricao} matches por Tag extraída da descrição`);
    console.log(`   ${matchPorNome} matches por nome/similaridade`);
    console.log(`   ${semMatch} sem match (usando classe C padrão)\n`);
    
    // Debug: Mostrar exemplos
    const exemploExato = ordens.find(o => o.metodoMatch === 'Tag campo BEM');
    const exemploExtraido = ordens.find(o => o.metodoMatch === 'Tag extraída da descrição');
    const exemploNome = ordens.find(o => o.metodoMatch === 'Match por nome');
    const exemploSemMatch = ordens.find(o => o.metodoMatch === 'Nenhum');
    
    if (exemploExato) {
      console.log(`   ✅ Exemplo Tag campo BEM:`);
      console.log(`      OS ${exemploExato.ordemServico} → Tag: ${exemploExato.tagIdentificada} → Classe: ${exemploExato.classe}`);
    }
    if (exemploExtraido) {
      console.log(`   ✅ Exemplo Tag extraída:`);
      console.log(`      OS ${exemploExtraido.ordemServico} → "${exemploExtraido.descricao}"`);
      console.log(`      Tag encontrada: ${exemploExtraido.tagIdentificada} → Classe: ${exemploExtraido.classe}`);
    }
    if (exemploNome) {
      console.log(`   ✅ Exemplo Match por nome:`);
      console.log(`      OS ${exemploNome.ordemServico} → Tag: ${exemploNome.tagIdentificada} → Classe: ${exemploNome.classe}`);
    }
    if (exemploSemMatch) {
      console.log(`   ⚠️ Exemplo sem match:`);
      console.log(`      OS ${exemploSemMatch.ordemServico} → Descrição: "${exemploSemMatch.descricao}"`);
      console.log(`      Classe padrão: C\n`);
    }
    
    return ordens;
  }

  /**
   * NOVO: Extrai padrões de TAG da descrição
   */
  extrairTagDaDescricao(descricao, ativos) {
    if (!descricao) return null;
    
    const texto = String(descricao).toUpperCase();
    
    // Padrões comuns de TAG
    const padroes = [
      /\b([A-Z]{2,3}[-\s]?\d{4,6})\b/g,
      /\b([A-Z]{2,4}\d{2,6})\b/g,
      /\b([A-Z]+[-\s]\d+)\b/g
    ];
    
    for (const padrao of padroes) {
      const matches = texto.matchAll(padrao);
      
      for (const match of matches) {
        let tagCandidato = match[1];
        const tagNormalizada = tagCandidato.replace(/[-\s]/g, '');
        
        if (ativos[tagNormalizada]) return tagNormalizada;
        
        const tagComHifen = tagCandidato.replace(/\s/g, '-');
        if (ativos[tagComHifen]) return tagComHifen;
        
        if (ativos[tagCandidato]) return tagCandidato;
      }
    }
    
    // Último recurso: buscar tags conhecidas no texto
    for (const tag of Object.keys(ativos)) {
      const tagUpper = tag.toUpperCase();
      if (texto.includes(tagUpper)) return tag;
      
      const tagSemSeparadores = tagUpper.replace(/[-\s]/g, '');
      if (texto.includes(tagSemSeparadores)) return tag;
    }
    
    return null;
  }

  /**
   * NOVO: Cria índice de descrições para busca rápida
   */
  criarIndiceDescricoes(ativos) {
    const indice = {};
    
    for (const [tag, ativo] of Object.entries(ativos)) {
      const descNormalizada = this.normalizarTexto(ativo.descricao);
      if (!indice[descNormalizada]) indice[descNormalizada] = [];
      indice[descNormalizada].push(tag);
      
      if (ativo.sigla) {
        const siglaNormalizada = this.normalizarTexto(ativo.sigla);
        if (!indice[siglaNormalizada]) indice[siglaNormalizada] = [];
        indice[siglaNormalizada].push(tag);
      }
      
      const palavras = descNormalizada.split(/\s+/);
      palavras.forEach(palavra => {
        if (palavra.length >= 4) {
          if (!indice[palavra]) indice[palavra] = [];
          if (!indice[palavra].includes(tag)) {
            indice[palavra].push(tag);
          }
        }
      });
    }
    
    return indice;
  }

  /**
   * NOVO: Busca por similaridade de texto
   */
  buscarPorSimilaridade(descricaoOS, indice, ativos) {
    const descNormalizada = this.normalizarTexto(descricaoOS);
    const palavras = descNormalizada.split(/\s+/).filter(p => p.length >= 4);
    
    const candidatos = new Map();
    
    palavras.forEach(palavra => {
      if (indice[palavra]) {
        indice[palavra].forEach(tag => {
          const scoreAtual = candidatos.get(tag) || 0;
          candidatos.set(tag, scoreAtual + 1);
        });
      }
      
      Object.keys(indice).forEach(chave => {
        if (chave.includes(palavra) || palavra.includes(chave)) {
          indice[chave].forEach(tag => {
            const scoreAtual = candidatos.get(tag) || 0;
            candidatos.set(tag, scoreAtual + 0.5);
          });
        }
      });
    });
    
    const candidatosOrdenados = Array.from(candidatos.entries())
      .sort((a, b) => b[1] - a[1]);
    
    if (candidatosOrdenados.length > 0 && candidatosOrdenados[0][1] > 1) {
      const melhorTag = candidatosOrdenados[0][0];
      return ativos[melhorTag];
    }
    
    return null;
  }

  /**
   * NOVO: Normaliza texto para comparação
   */
  normalizarTexto(texto) {
    if (!texto) return '';
    return String(texto)
      .toUpperCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^\w\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /**
   * Lê cadastro de ativos (bens)
   */
  async lerAtivos(caminhoArquivo) {
    console.log('📖 Lendo cadastro de ativos...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(caminhoArquivo);
    const sheet = workbook.worksheets[0];
    
    const ativos = {};
    let headerRow = null;
    
    for (let i = 1; i <= 20; i++) {
      const row = sheet.getRow(i);
      let temTag = false;
      let temDescricao = false;
      let temClasse = false;
      
      row.eachCell((cell) => {
        const valor = String(cell.value || '').trim().toUpperCase();
        if (valor === 'TAG') temTag = true;
        if (valor.includes('DESCRIÇÃO') || valor.includes('DESCRICAO')) temDescricao = true;
        if (valor === 'CLASSE') temClasse = true;
      });
      
      if (temTag && temDescricao && temClasse) {
        headerRow = i;
        console.log(`   📍 Header de ativos encontrado na linha ${i}`);
        break;
      }
    }
    
    if (!headerRow) {
      console.log('⚠️ Header não encontrado automaticamente. Tentando linha 8...');
      headerRow = 8;
    }
    
    const headers = {};
    const row = sheet.getRow(headerRow);
    row.eachCell((cell, colNumber) => {
      if (cell.value) {
        const headerName = String(cell.value).trim();
        headers[headerName] = colNumber;
      }
    });
    
    console.log('📋 Colunas encontradas:', Object.keys(headers));
    
    let count = 0;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber <= headerRow) return;
      if (!row.hasValues) return;
      
      const tag = this.getCellValue(row, headers['Tag']);
      const descricao = this.getCellValue(row, headers['Descrição'] || headers['Descricao']);
      const classe = this.getCellValue(row, headers['Classe']);
      
      if (tag && descricao && classe) {
        const tagStr = String(tag).trim();
        const classeStr = String(classe).trim().toUpperCase();
        
        if (classeStr === 'A' || classeStr === 'B' || classeStr === 'C') {
          ativos[tagStr] = {
            tag: tagStr,
            sigla: this.getCellValue(row, headers['Sigla']) || '',
            descricao: String(descricao).trim(),
            local: this.getCellValue(row, headers['Local']) || '',
            classe: classeStr
          };
          count++;
        }
      }
    });
    
    console.log(`✅ ${count} ativos carregados (indexados por Tag)`);
    
    const primeirosTres = Object.entries(ativos).slice(0, 3);
    console.log(`   Exemplos de Tags cadastradas:`);
    primeirosTres.forEach(([tag, ativo]) => {
      console.log(`   - ${tag}: ${ativo.descricao} (Classe ${ativo.classe})`);
    });
    
    const classesCount = { A: 0, B: 0, C: 0 };
    Object.values(ativos).forEach(ativo => {
      classesCount[ativo.classe]++;
    });
    console.log(`   📊 Distribuição: A=${classesCount.A}, B=${classesCount.B}, C=${classesCount.C}\n`);
    
    return ativos;
  }

  /**
   * Exporta resultado para Excel
   */
  async exportarResultado(ordensProcessadas, caminhoSaida) {
    console.log('📝 Gerando planilha de resultado...');
    
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Cronograma Priorizado');
    
    sheet.columns = [
      { header: 'Prioridade', key: 'prioridade', width: 12 },
      { header: 'Score', key: 'score', width: 10 },
      { header: 'Ordem Serv.', key: 'ordemServico', width: 15 },
      { header: 'Tag Identificada', key: 'tagIdentificada', width: 18 },
      { header: 'Bem (Tag)', key: 'bem', width: 15 },
      { header: 'Equipamento', key: 'equipamento', width: 30 },
      { header: 'Classe', key: 'classe', width: 10 },
      { header: 'Método Match', key: 'metodoMatch', width: 25 },
      { header: 'Prev. Início', key: 'prevInicio', width: 15 },
      { header: 'Dias Atraso', key: 'diasAtraso', width: 12 },
      { header: 'Score Crit.', key: 'scoreCriticidade', width: 12 },
      { header: 'Score Atraso', key: 'scoreAtraso', width: 12 },
      { header: 'Descrição', key: 'descricao', width: 40 }
    ];
    
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F4E78' }
    };
    sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
    
    ordensProcessadas.forEach((ordem, index) => {
      const row = sheet.addRow({
        prioridade: index + 1,
        score: ordem.scoreFinal.toFixed(1),
        ordemServico: ordem.ordemServico,
        tagIdentificada: ordem.tagIdentificada || 'N/A',
        bem: ordem.bem || 'N/A',
        equipamento: ordem.equipamento || ordem.nomeBem || 'N/A',
        classe: ordem.classe,
        metodoMatch: ordem.metodoMatch || 'N/A',
        prevInicio: ordem.prevInicio ? this.formatarData(ordem.prevInicio) : 'N/A',
        diasAtraso: ordem.diasAtraso || 0,
        scoreCriticidade: ordem.scoreCriticidade.toFixed(1),
        scoreAtraso: ordem.scoreAtraso.toFixed(1),
        descricao: ordem.descricao || ''
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
    });
    
    await workbook.xlsx.writeFile(caminhoSaida);
    console.log(`✅ Planilha gerada: ${caminhoSaida}`);
  }

  /**
   * Utilitários
   */
  getCellValue(row, colNumber) {
    if (!colNumber) return null;
    try {
      const cell = row.getCell(colNumber);
      return cell.value || null;
    } catch (error) {
      return null;
    }
  }

  parseData(valor) {
    if (!valor) return null;
    if (valor instanceof Date) return valor;
    
    if (typeof valor === 'string') {
      const partes = valor.split('/');
      if (partes.length === 3) {
        return new Date(partes[2], partes[1] - 1, partes[0]);
      }
    }
    
    return null;
  }

  formatarData(data) {
    if (!data) return '';
    const dia = String(data.getDate()).padStart(2, '0');
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const ano = data.getFullYear();
    return `${dia}/${mes}/${ano}`;
  }
}

module.exports = new ExcelHandler();