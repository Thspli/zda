const ExcelJS = require('exceljs');

/**
 * Gerenciador de Funcionários/Manutentores
 * Lê disponibilidade de técnicos e aloca OS baseado em especialidade e turno
 */
class FuncionariosHandler {
  /**
   * Lê planilha de funcionários (manutentores)
   */
  async lerFuncionarios(caminhoArquivo) {
    console.log('👷 Lendo cadastro de funcionários...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(caminhoArquivo);
    const sheet = workbook.worksheets[0];
    
    const funcionarios = [];
    const headers = {};
    
    // Encontrar linha de headers
    let headerRowNumber = 1;
    for (let i = 1; i <= 10; i++) {
      const row = sheet.getRow(i);
      let foundColumn = false;
      
      row.eachCell((cell) => {
        const value = String(cell.value || '').trim().toUpperCase();
        if (value.includes('AREA') || value.includes('FUNÇÃO') || value.includes('FUNCAO')) {
          foundColumn = true;
        }
      });
      
      if (foundColumn) {
        headerRowNumber = i;
        break;
      }
    }
    
    console.log(`   📋 Header encontrado na linha ${headerRowNumber}`);
    
    // Mapear headers
    const headerRow = sheet.getRow(headerRowNumber);
    headerRow.eachCell((cell, colNumber) => {
      if (cell.value) {
        const headerName = String(cell.value).trim();
        headers[headerName] = colNumber;
      }
    });
    
    console.log('   📋 Colunas encontradas:', Object.keys(headers));
    
    // Processar funcionários
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber <= headerRowNumber) return;
      if (!row.hasValues) return;
      
      const area = this.getCellValue(row, headers['AREA']);
      const numero = this.getCellValue(row, headers['NUMERO']);
      const nome = this.getCellValue(row, headers['FUNCIONARIO']);
      const funcao = this.getCellValue(row, headers['FUNÇÃO'] || headers['FUNCAO']);
      const turno = this.getCellValue(row, headers['TURNO'] || headers['FUNÇÃO']); // Pode estar em FUNÇÃO
      const horario = this.getCellValue(row, headers['HORARIO'] || headers['horário']);
      
      if (area && numero && nome) {
        // Extrair turno se estiver em FUNÇÃO (ex: "TURNO A")
        let turnoIdentificado = null;
        if (turno) {
          const turnoStr = String(turno).toUpperCase();
          if (turnoStr.includes('TURNO A')) turnoIdentificado = 'A';
          else if (turnoStr.includes('TURNO B')) turnoIdentificado = 'B';
          else if (turnoStr.includes('TURNO C')) turnoIdentificado = 'C';
          else if (turnoStr.includes('ADMINISTRATIVO')) turnoIdentificado = 'ADM';
        }
        
        // Extrair horário se estiver na última coluna
        let horarioIdentificado = horario || this.getCellValue(row, row.cellCount);
        
        const funcionario = {
          numero: String(numero),
          nome: String(nome),
          area: String(area).trim(),
          funcao: funcao ? String(funcao).trim() : '',
          turno: turnoIdentificado,
          horario: horarioIdentificado ? String(horarioIdentificado) : '',
          osAlocadas: 0, // Contador de OS alocadas
          disponivel: true
        };
        
        funcionarios.push(funcionario);
      }
    });
    
    console.log(`✅ ${funcionarios.length} funcionários carregados`);
    
    // Estatísticas por área
    const porArea = this.agruparPorArea(funcionarios);
    console.log('\n   📊 Distribuição por área:');
    Object.keys(porArea).forEach(area => {
      console.log(`      ${area}: ${porArea[area].length} técnicos`);
    });
    
    // Estatísticas por turno
    const porTurno = this.agruparPorTurno(funcionarios);
    console.log('\n   📊 Distribuição por turno:');
    Object.keys(porTurno).forEach(turno => {
      console.log(`      Turno ${turno}: ${porTurno[turno].length} técnicos`);
    });
    
    console.log('');
    
    return funcionarios;
  }

  /**
   * Identifica especialidade necessária da OS
   */
  identificarEspecialidade(ordem) {
    const descricao = String(ordem.descricao || '').toUpperCase();
    const equipamento = String(ordem.equipamento || '').toUpperCase();
    const texto = `${descricao} ${equipamento}`;
    
    // Palavras-chave por área
    const palavrasChave = {
      'Eletromecanicos': [
        'ELETRI', 'ELETRO', 'MOTOR', 'AUTOMAÇÃO', 'INVERSOR', 'PAINEL',
        'SENSOR', 'CLP', 'VARIADOR', 'COMANDO', 'ELÉTRICO', 'ELETRICO'
      ],
      'Mecânico Geral': [
        'MECANICO', 'MECÂNICO', 'ROLAMENTO', 'CORREIA', 'CORRENTE',
        'REDUTOR', 'ACOPLAMENTO', 'EIXO', 'MANCAL', 'LUBRIFICAÇÃO', 'LUBRIFICACAO'
      ],
      'Mecânico Embalagem': [
        'EMBALAGEM', 'EMBALADORA', 'SELADORA', 'ENVASE', 'EMPACOTADORA',
        'FLOWPACK', 'SACHET', 'DOSADORA'
      ],
      'Serviços gerais': [
        'LIMPEZA', 'PINTURA', 'SOLDA', 'ESTRUTURA', 'GERAL'
      ]
    };
    
    // Verificar cada área
    for (const [area, palavras] of Object.entries(palavrasChave)) {
      for (const palavra of palavras) {
        if (texto.includes(palavra)) {
          return area;
        }
      }
    }
    
    // Padrão: Mecânico Geral
    return 'Mecânico Geral';
  }

  /**
   * Busca técnicos disponíveis para uma OS
   */
  buscarTecnicosDisponiveis(ordem, funcionarios, turnoSlot) {
    const especialidadeNecessaria = this.identificarEspecialidade(ordem);
    
    // Filtrar por especialidade
    let candidatos = funcionarios.filter(f => 
      this.normalizarArea(f.area) === this.normalizarArea(especialidadeNecessaria)
    );
    
    // Filtrar por turno se especificado
    if (turnoSlot) {
      candidatos = candidatos.filter(f => {
        if (!f.turno) return true; // Se não tem turno definido, considera disponível
        if (f.turno === 'ADM') return true; // Administrativo trabalha em todos os turnos
        return f.turno === turnoSlot;
      });
    }
    
    // Ordenar por quantidade de OS já alocadas (balanceamento de carga)
    candidatos.sort((a, b) => a.osAlocadas - b.osAlocadas);
    
    return candidatos;
  }

  /**
   * Aloca técnico para uma OS
   */
  alocarTecnico(ordem, funcionarios, turnoSlot) {
    const tecnicosDisponiveis = this.buscarTecnicosDisponiveis(ordem, funcionarios, turnoSlot);
    
    if (tecnicosDisponiveis.length === 0) {
      return null;
    }
    
    // Pegar técnico com menos OS alocadas
    const tecnicoSelecionado = tecnicosDisponiveis[0];
    tecnicoSelecionado.osAlocadas++;
    
    return {
      numero: tecnicoSelecionado.numero,
      nome: tecnicoSelecionado.nome,
      area: tecnicoSelecionado.area,
      funcao: tecnicoSelecionado.funcao,
      turno: tecnicoSelecionado.turno
    };
  }

  /**
   * Gera relatório de alocação de técnicos
   */
  gerarRelatorioTecnicos(funcionarios) {
    console.log('\n📊 RELATÓRIO DE ALOCAÇÃO DE TÉCNICOS:');
    console.log('═══════════════════════════════════════════════════════════');
    
    const porArea = this.agruparPorArea(funcionarios);
    
    Object.keys(porArea).forEach(area => {
      const tecnicos = porArea[area];
      const totalOS = tecnicos.reduce((sum, t) => sum + t.osAlocadas, 0);
      const mediaOS = totalOS / tecnicos.length;
      
      console.log(`\n${area}:`);
      console.log(`   Total de técnicos: ${tecnicos.length}`);
      console.log(`   Total de OS alocadas: ${totalOS}`);
      console.log(`   Média por técnico: ${mediaOS.toFixed(1)}`);
      
      // Top 3 mais alocados
      const top3 = tecnicos
        .filter(t => t.osAlocadas > 0)
        .sort((a, b) => b.osAlocadas - a.osAlocadas)
        .slice(0, 3);
      
      if (top3.length > 0) {
        console.log(`   Top 3 mais alocados:`);
        top3.forEach((t, i) => {
          console.log(`      ${i + 1}. ${t.nome} - ${t.osAlocadas} OS`);
        });
      }
    });
    
    console.log('═══════════════════════════════════════════════════════════\n');
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

  normalizarArea(area) {
    if (!area) return '';
    const normalizado = String(area)
      .toUpperCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, ''); // Remove acentos
    
    // Mapeamento de variações
    if (normalizado.includes('ELETROMECANICO')) return 'ELETROMECANICOS';
    if (normalizado.includes('MECANICO GERAL')) return 'MECANICO GERAL';
    if (normalizado.includes('MECANICO EMBALAGEM')) return 'MECANICO EMBALAGEM';
    if (normalizado.includes('SERVICOS GERAIS') || normalizado.includes('SERVICO GERAL')) return 'SERVICOS GERAIS';
    
    return normalizado;
  }

  agruparPorArea(funcionarios) {
    const grupos = {};
    funcionarios.forEach(f => {
      const area = f.area;
      if (!grupos[area]) grupos[area] = [];
      grupos[area].push(f);
    });
    return grupos;
  }

  agruparPorTurno(funcionarios) {
    const grupos = {};
    funcionarios.forEach(f => {
      const turno = f.turno || 'Não definido';
      if (!grupos[turno]) grupos[turno] = [];
      grupos[turno].push(f);
    });
    return grupos;
  }
}

module.exports = new FuncionariosHandler();