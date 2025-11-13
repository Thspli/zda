/**
 * CORRIGIDO: Módulo de Cálculo PCM
 * Cálculo correto de criticidade e backlog (atraso)
 */

class PCMScoring {
  constructor() {
    // Pesos dos critérios
    this.pesos = {
      criticidade: 0.4,
      atraso: 0.4,
      perfilTecnico: 0.2,
      disponibilidade: 0.0
    };
  }

  /**
   * CORRIGIDO: Processa ordens com cálculo correto
   */
  processarOrdens(ordens, ativos, dataReferencia = new Date()) {
    console.log('🔄 Iniciando processamento PCM...');
    console.log(`   Data de referência: ${this.formatarData(dataReferencia)}`);
    
    const ordensProcessadas = ordens.map(ordem => {
      // CORREÇÃO: Classe já vem do enriquecimento (Tag → Classe)
      const classe = ordem.classe || 'C';
      
      // Calcular scores das 4 etapas
      const scoreCriticidade = this.calcularCriticidade(classe);
      const scoreAtraso = this.calcularAtraso(ordem.prevInicio, dataReferencia);
      const scorePerfilTecnico = this.calcularPerfilTecnico(ordem);
      const scoreDisponibilidade = this.calcularDisponibilidade(ordem);
      
      // Score final ponderado
      const scoreFinal = 
        (scoreCriticidade * this.pesos.criticidade) +
        (scoreAtraso * this.pesos.atraso) +
        (scorePerfilTecnico * this.pesos.perfilTecnico) +
        (scoreDisponibilidade * this.pesos.disponibilidade);
      
      // Calcular dias de atraso
      const diasAtraso = this.calcularDiasAtraso(ordem.prevInicio, dataReferencia);
      
      return {
        ...ordem,
        classe,
        scoreCriticidade,
        scoreAtraso,
        scorePerfilTecnico,
        scoreDisponibilidade,
        scoreFinal,
        diasAtraso
      };
    });
    
    // Ordenar por score (maior = mais prioritário)
    ordensProcessadas.sort((a, b) => b.scoreFinal - a.scoreFinal);
    
    console.log(`✅ ${ordensProcessadas.length} ordens processadas e ordenadas`);
    this.exibirTop10(ordensProcessadas);
    this.exibirEstatisticasDetalhadas(ordensProcessadas);
    
    return ordensProcessadas;
  }

  /**
   * ETAPA 1: Criticidade do Ativo
   * A=100, B=50, C=25
   */
  calcularCriticidade(classe) {
    const scores = {
      'A': 100,
      'B': 50,
      'C': 25
    };
    
    return scores[classe] || scores['C'];
  }

  /**
   * ETAPA 2: Perfil Técnico (simplificado)
   */
  calcularPerfilTecnico(ordem) {
    return 75;
  }

  /**
   * CORRIGIDO: ETAPA 3 - BackLog (Atraso)
   * Cálculo baseado em dias corridos de atraso
   */
  calcularAtraso(prevInicio, dataReferencia) {
    if (!prevInicio) {
      console.log('   ⚠️ OS sem data prevista - score atraso = 0');
      return 0;
    }
    
    const diasAtraso = this.calcularDiasAtraso(prevInicio, dataReferencia);
    
    // Escala progressiva de urgência
    if (diasAtraso <= 0) return 0;           // No prazo ou futuro
    if (diasAtraso <= 7) return 30;          // Até 1 semana atrasado
    if (diasAtraso <= 15) return 60;         // Até 2 semanas
    if (diasAtraso <= 30) return 85;         // Até 1 mês
    return 100;                               // Mais de 1 mês atrasado
  }

  /**
   * ETAPA 4: Disponibilidade (simplificado)
   */
  calcularDisponibilidade(ordem) {
    return 50;
  }

  /**
   * CORRIGIDO: Calcula dias de atraso corretamente
   */
  calcularDiasAtraso(prevInicio, dataReferencia) {
    if (!prevInicio) return 0;
    
    // Garantir que ambas são Date objects
    const dataInicio = prevInicio instanceof Date ? prevInicio : new Date(prevInicio);
    const dataRef = dataReferencia instanceof Date ? dataReferencia : new Date(dataReferencia);
    
    // Calcular diferença em milissegundos
    const diffMs = dataRef.getTime() - dataInicio.getTime();
    
    // Converter para dias (arredondar para baixo)
    const diffDias = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    
    // Retornar apenas se positivo (atrasado)
    return Math.max(0, diffDias);
  }

  /**
   * Exibe top 10 ordens priorizadas
   */
  exibirTop10(ordens) {
    console.log('\n🏆 TOP 10 ORDENS MAIS PRIORITÁRIAS:');
    console.log('═══════════════════════════════════════════════════════════');
    console.log('Rank | Score | OS      | Bem     | Classe | Crit | Atraso | Dias');
    console.log('─────────────────────────────────────────────────────────────');
    
    ordens.slice(0, 10).forEach((ordem, index) => {
      const indicador = ordem.scoreFinal >= 80 ? '🔴' : 
                       ordem.scoreFinal >= 50 ? '🟡' : '🟢';
      
      const rank = String(index + 1).padStart(4);
      const score = ordem.scoreFinal.toFixed(1).padStart(5);
      const os = String(ordem.ordemServico).padEnd(8);
      const bem = String(ordem.bem || 'N/A').padEnd(8);
      const classe = ordem.classe.padEnd(6);
      const crit = ordem.scoreCriticidade.toFixed(0).padStart(4);
      const atraso = ordem.scoreAtraso.toFixed(0).padStart(6);
      const dias = String(ordem.diasAtraso).padStart(4);
      
      console.log(`${indicador} ${rank} | ${score} | ${os} | ${bem} | ${classe} | ${crit} | ${atraso} | ${dias}`);
    });
    
    console.log('═══════════════════════════════════════════════════════════\n');
  }

  /**
   * Exibe estatísticas detalhadas
   */
  exibirEstatisticasDetalhadas(ordens) {
    console.log('📊 ESTATÍSTICAS DO PROCESSAMENTO:');
    console.log('─────────────────────────────────');
    
    const total = ordens.length;
    
    // Por prioridade
    const criticas = ordens.filter(o => o.scoreFinal >= 80).length;
    const medias = ordens.filter(o => o.scoreFinal >= 50 && o.scoreFinal < 80).length;
    const baixas = ordens.filter(o => o.scoreFinal < 50).length;
    
    console.log(`Prioridade:`);
    console.log(`  🔴 Críticas (≥80): ${criticas} (${(criticas/total*100).toFixed(1)}%)`);
    console.log(`  🟡 Médias (50-79): ${medias} (${(medias/total*100).toFixed(1)}%)`);
    console.log(`  🟢 Baixas (<50): ${baixas} (${(baixas/total*100).toFixed(1)}%)`);
    
    // Por classe
    const classeA = ordens.filter(o => o.classe === 'A').length;
    const classeB = ordens.filter(o => o.classe === 'B').length;
    const classeC = ordens.filter(o => o.classe === 'C').length;
    
    console.log(`\nClasse do Equipamento:`);
    console.log(`  Classe A: ${classeA} (${(classeA/total*100).toFixed(1)}%)`);
    console.log(`  Classe B: ${classeB} (${(classeB/total*100).toFixed(1)}%)`);
    console.log(`  Classe C: ${classeC} (${(classeC/total*100).toFixed(1)}%)`);
    
    // Por atraso
    const atrasadas = ordens.filter(o => o.diasAtraso > 0).length;
    const noPrazo = total - atrasadas;
    const mediaAtraso = ordens.reduce((acc, o) => acc + o.diasAtraso, 0) / total;
    
    console.log(`\nAtraso (BackLog):`);
    console.log(`  Atrasadas: ${atrasadas} (${(atrasadas/total*100).toFixed(1)}%)`);
    console.log(`  No prazo: ${noPrazo} (${(noPrazo/total*100).toFixed(1)}%)`);
    console.log(`  Média de atraso: ${mediaAtraso.toFixed(1)} dias`);
    
    // Alertas
    const muitoAtrasadas = ordens.filter(o => o.diasAtraso > 30).length;
    if (muitoAtrasadas > 0) {
      console.log(`\n⚠️  ALERTA: ${muitoAtrasadas} OS com mais de 30 dias de atraso!`);
    }
    
    const criticasAtrasadas = ordens.filter(o => o.classe === 'A' && o.diasAtraso > 0).length;
    if (criticasAtrasadas > 0) {
      console.log(`⚠️  ALERTA: ${criticasAtrasadas} equipamentos Classe A com atraso!`);
    }
    
    console.log('');
  }

  /**
   * Utilitários
   */
  formatarData(data) {
    if (!data) return 'N/A';
    const dia = String(data.getDate()).padStart(2, '0');
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const ano = data.getFullYear();
    return `${dia}/${mes}/${ano}`;
  }

  /**
   * Gera estatísticas estruturadas
   */
  gerarEstatisticas(ordens) {
    const total = ordens.length;
    const criticas = ordens.filter(o => o.scoreFinal >= 80).length;
    const medias = ordens.filter(o => o.scoreFinal >= 50 && o.scoreFinal < 80).length;
    const baixas = ordens.filter(o => o.scoreFinal < 50).length;
    const atrasadas = ordens.filter(o => o.diasAtraso > 0).length;
    
    const classeA = ordens.filter(o => o.classe === 'A').length;
    const classeB = ordens.filter(o => o.classe === 'B').length;
    const classeC = ordens.filter(o => o.classe === 'C').length;
    
    return {
      total,
      prioridade: { criticas, medias, baixas },
      atraso: atrasadas,
      classes: { A: classeA, B: classeB, C: classeC }
    };
  }
}

module.exports = new PCMScoring();