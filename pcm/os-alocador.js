const funcionariosHandler = require('./funcionarios-handler');

/**
 * Alocador Inteligente de Ordens de Serviço
 * ATUALIZADO: Considera disponibilidade de técnicos
 */
class OSAlocador {
  /**
   * Aloca ordens de serviço priorizadas nos slots do calendário
   * NOVO: Considera disponibilidade de técnicos por área e turno
   */
  alocarOrdens(ordensProcessadas, calendario, funcionarios = null) {
    console.log('🎯 Iniciando alocação de OS no calendário...');
    
    if (funcionarios) {
      console.log(`   👷 Modo: Alocação com ${funcionarios.length} técnicos`);
    } else {
      console.log('   ⚠️ Modo: Alocação sem verificação de técnicos');
    }
    
    const slots = [...calendario.slots];
    let osAlocadas = 0;
    let osPendentes = 0;
    let osSemTecnico = 0;
    
    // Processar cada OS por ordem de prioridade
    ordensProcessadas.forEach(ordem => {
      // Encontrar slots compatíveis
      const slotsCompativeis = this.encontrarSlotsCompativeis(ordem, slots);
      
      if (slotsCompativeis.length === 0) {
        ordem.alocada = false;
        ordem.dataProgramada = null;
        ordem.motivoNaoAlocacao = 'Sem slot disponível';
        osPendentes++;
        return;
      }
      
      // Se tem funcionários, verificar disponibilidade de técnico
      if (funcionarios) {
        let slotComTecnico = null;
        
        // Tentar cada slot até encontrar um com técnico disponível
        for (const slot of slotsCompativeis) {
          const tecnicoDisponivel = funcionariosHandler.alocarTecnico(
            ordem, 
            funcionarios, 
            slot.turno
          );
          
          if (tecnicoDisponivel) {
            slotComTecnico = slot;
            slot.tecnico = tecnicoDisponivel;
            ordem.tecnicoAlocado = tecnicoDisponivel;
            break;
          }
        }
        
        if (!slotComTecnico) {
          ordem.alocada = false;
          ordem.dataProgramada = null;
          ordem.motivoNaoAlocacao = 'Sem técnico disponível no turno';
          osSemTecnico++;
          return;
        }
        
        // Usar slot com técnico
        slotComTecnico.ocupado = true;
        slotComTecnico.os = ordem;
        ordem.alocada = true;
        ordem.dataProgramada = `${slotComTecnico.data} - Turno ${slotComTecnico.turno}`;
        ordem.slotAlocado = slotComTecnico;
        osAlocadas++;
        
      } else {
        // Modo sem verificação de técnicos (comportamento antigo)
        const melhorSlot = slotsCompativeis[0];
        melhorSlot.ocupado = true;
        melhorSlot.os = ordem;
        ordem.alocada = true;
        ordem.dataProgramada = `${melhorSlot.data} - Turno ${melhorSlot.turno}`;
        ordem.slotAlocado = melhorSlot;
        osAlocadas++;
      }
    });
    
    console.log(`✅ Alocação concluída:`);
    console.log(`   ${osAlocadas} OS programadas`);
    
    if (funcionarios) {
      console.log(`   ${osSemTecnico} OS sem técnico disponível`);
      console.log(`   ${osPendentes} OS sem slot disponível`);
    } else {
      console.log(`   ${osPendentes} OS pendentes (sem slot disponível)`);
    }
    
    return {
      ordensProcessadas,
      slots,
      funcionarios,
      estatisticas: {
        alocadas: osAlocadas,
        semTecnico: osSemTecnico,
        pendentes: osPendentes,
        total: ordensProcessadas.length
      }
    };
  }

  /**
   * Encontra slots compatíveis com uma OS
   */
  encontrarSlotsCompativeis(ordem, slots) {
    const slotsDisponiveis = slots.filter(s => !s.ocupado);
    
    // Tentar encontrar slots da mesma linha
    const equipamentoOS = ordem.equipamento || ordem.bem;
    const slotsExatos = slotsDisponiveis.filter(slot => 
      this.equipamentosCompativeis(slot.linha, equipamentoOS) ||
      this.equipamentosCompativeis(slot.equipamento, equipamentoOS)
    );
    
    if (slotsExatos.length > 0) {
      return this.ordenarSlotsPorData(slotsExatos);
    }
    
    // Se não encontrou slot exato, tentar por área
    const localOS = ordem.local || '';
    const slotsPorArea = slotsDisponiveis.filter(slot =>
      this.locaisCompativeis(slot.linha, localOS)
    );
    
    if (slotsPorArea.length > 0) {
      return this.ordenarSlotsPorData(slotsPorArea);
    }
    
    // Último caso: qualquer slot disponível
    return this.ordenarSlotsPorData(slotsDisponiveis);
  }

  /**
   * Verifica se equipamentos são compatíveis
   */
  equipamentosCompativeis(equipamentoSlot, equipamentoOS) {
    if (!equipamentoSlot || !equipamentoOS) return false;
    
    const slot = String(equipamentoSlot).toUpperCase().trim();
    const os = String(equipamentoOS).toUpperCase().trim();
    
    if (slot === os) return true;
    if (slot.includes(os) || os.includes(slot)) return true;
    
    return false;
  }

  /**
   * Verifica se locais são compatíveis
   */
  locaisCompativeis(linhaSlot, localOS) {
    if (!linhaSlot || !localOS) return false;
    
    const slot = String(linhaSlot).toUpperCase();
    const local = String(localOS).toUpperCase();
    
    const areas = ['CANDY', 'MARSHMALLOW', 'MOLDADOS', 'CHIPS', 'EMBALAGEM'];
    
    for (const area of areas) {
      if (slot.includes(area) && local.includes(area)) {
        return true;
      }
    }
    
    return false;
  }

  /**
   * Ordena slots por data
   */
  ordenarSlotsPorData(slots) {
    return slots.sort((a, b) => {
      const dataA = this.parseDataSlot(a.data);
      const dataB = this.parseDataSlot(b.data);
      return dataA - dataB;
    });
  }

  /**
   * Converte string de data em objeto Date
   */
  parseDataSlot(dataStr) {
    try {
      const partes = dataStr.split('/');
      if (partes.length === 3) {
        return new Date(partes[2], partes[0] - 1, partes[1]);
      }
    } catch (e) {
      // Se falhar, retornar data futura
    }
    return new Date(2099, 0, 1);
  }

  /**
   * Gera resumo da alocação
   */
  gerarResumoAlocacao(resultado) {
    console.log('\n📋 RESUMO DA ALOCAÇÃO:');
    console.log('═══════════════════════════════════════════════════════════');
    
    const { ordensProcessadas, estatisticas, funcionarios } = resultado;
    
    console.log(`Total de OS: ${estatisticas.total}`);
    console.log(`OS Programadas: ${estatisticas.alocadas} (${(estatisticas.alocadas/estatisticas.total*100).toFixed(1)}%)`);
    
    if (funcionarios) {
      console.log(`OS sem técnico: ${estatisticas.semTecnico}`);
    }
    
    console.log(`OS Pendentes: ${estatisticas.pendentes}`);
    
    // Distribuição por prioridade
    const criticasAlocadas = ordensProcessadas.filter(o => o.alocada && o.scoreFinal >= 80).length;
    const mediasAlocadas = ordensProcessadas.filter(o => o.alocada && o.scoreFinal >= 50 && o.scoreFinal < 80).length;
    const baixasAlocadas = ordensProcessadas.filter(o => o.alocada && o.scoreFinal < 50).length;
    
    console.log(`\n📊 Distribuição por Prioridade (Alocadas):`);
    console.log(`   🔴 Críticas: ${criticasAlocadas}`);
    console.log(`   🟡 Médias: ${mediasAlocadas}`);
    console.log(`   🟢 Baixas: ${baixasAlocadas}`);
    
    // OS críticas pendentes
    const criticasPendentes = ordensProcessadas
      .filter(o => !o.alocada && o.scoreFinal >= 80)
      .slice(0, 5);
    
    if (criticasPendentes.length > 0) {
      console.log(`\n⚠️  OS CRÍTICAS PENDENTES:`);
      criticasPendentes.forEach(ordem => {
        const motivo = ordem.motivoNaoAlocacao || 'Sem slot';
        console.log(`   OS ${ordem.ordemServico} | Score: ${ordem.scoreFinal.toFixed(1)} | Motivo: ${motivo}`);
      });
    }
    
    // Relatório de técnicos
    if (funcionarios) {
      funcionariosHandler.gerarRelatorioTecnicos(funcionarios);
    }
    
    console.log('═══════════════════════════════════════════════════════════\n');
  }
}

module.exports = new OSAlocador();