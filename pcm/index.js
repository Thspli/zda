const path = require('path');
const fs = require('fs');
const excelHandler = require('./excel-handler');
const calendarioHandler = require('./calendario-handler');
const pcmScoring = require('./pcm-scoring');
const osAlocador = require('./os-alocador');
const funcionariosHandler = require('./funcionarios-handler');

async function main() {
  try {
    const pastaAtual = process.argv[2] || process.cwd();
    const pastaOutput = path.join(pastaAtual, 'output');

    if (!fs.existsSync(pastaOutput)) {
      fs.mkdirSync(pastaOutput);
      console.log('📁 Pasta output criada\n');
    }

    console.log('PASSO 1: Carregando dados...');
    console.log('─────────────────────────────────');

    const caminhoCalendario = path.join(pastaAtual, 'calendario-pcp.xlsx');
    const caminhoOS = path.join(pastaAtual, 'ordens-servico.xlsx');
    const caminhoAtivos = path.join(pastaAtual, 'Controle-Bens-SENAI-SPRINT-1.xlsx');
    const caminhoFuncionarios = path.join(pastaAtual, 'funcionarios.xlsx');

    const calendario = await calendarioHandler.lerCalendarioPCP(caminhoCalendario);
    const ordensOS = await excelHandler.lerOrdensServicoSimplificada(caminhoOS);
    const ativos = await excelHandler.lerAtivos(caminhoAtivos);

    let funcionarios = null;
    if (fs.existsSync(caminhoFuncionarios)) {
      funcionarios = await funcionariosHandler.lerFuncionarios(caminhoFuncionarios);
    } else {
      console.log('⚠️ Arquivo funcionarios.xlsx não encontrado');
      console.log('   Sistema rodará sem verificação de técnicos\n');
    }

    if (ordensOS.length === 0) {
      console.error('❌ Nenhuma ordem de serviço encontrada!');
      return;
    }

    console.log('\nPASSO 2: Vinculando OS com equipamentos...');
    console.log('─────────────────────────────────');

    const ordensEnriquecidas = excelHandler.enriquecerOSComEquipamentos(ordensOS, ativos);

    if (funcionarios) {
      console.log('🔧 Identificando especialidades necessárias...');
      ordensEnriquecidas.forEach(ordem => {
        ordem.especialidadeNecessaria = funcionariosHandler.identificarEspecialidade(ordem);
      });
      
      const porEspecialidade = {};
      ordensEnriquecidas.forEach(ordem => {
        const esp = ordem.especialidadeNecessaria;
        porEspecialidade[esp] = (porEspecialidade[esp] || 0) + 1;
      });
      
      console.log('   📊 OS por especialidade:');
      Object.keys(porEspecialidade).forEach(esp => {
        console.log(`      ${esp}: ${porEspecialidade[esp]} OS`);
      });
      console.log('');
    }

    console.log('\nPASSO 3: Aplicando algoritmo PCM...');
    console.log('─────────────────────────────────');
    console.log('Critérios de Priorização:');
    console.log('  - Criticidade (40%): Classe A=100, B=50, C=25');
    console.log('  - Atraso (40%): Score baseado em urgência');
    console.log('  - Perfil Técnico (20%): Disponibilidade de recursos');
    console.log('');

    const dataReferencia = new Date();
    const ordensProcessadas = pcmScoring.processarOrdens(ordensEnriquecidas, ativos, dataReferencia);

    console.log('PASSO 4: Alocando OS no calendário...');
    console.log('─────────────────────────────────');

    const resultadoAlocacao = osAlocador.alocarOrdens(
      ordensProcessadas, 
      calendario,
      funcionarios
    );

    osAlocador.gerarResumoAlocacao(resultadoAlocacao);

    console.log('PASSO 5: Gerando arquivos de saída...');
    console.log('─────────────────────────────────');

    const dataHora = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);

    const caminhoCalendarioSaida = path.join(pastaOutput, `CALENDARIO-PREENCHIDO-${dataHora}.xlsx`);
    await calendarioHandler.exportarCalendarioPreenchido(
      caminhoCalendario,
      resultadoAlocacao.slots,
      caminhoCalendarioSaida
    );

    console.log('📊 Adicionando tipo de OS do TOTVS...');
    resultadoAlocacao.ordensProcessadas.forEach(ordem => {
      const ordemOriginal = ordensOS.find(o => String(o.OS || o.Ordem).trim() === String(ordem.OS || ordem.Ordem).trim());
      
      if (ordemOriginal) {
        ordem.TIPO_TOTVS = (ordemOriginal['Tipo O. S.'] || 'N/A').toUpperCase().trim();
        ordem.CONFIANCA_TOTVS = '100%';
      } else {
        ordem.TIPO_TOTVS = 'N/A';
        ordem.CONFIANCA_TOTVS = 'N/A';
      }
    });

    const caminhoClassificacao = path.join(pastaOutput, `CLASSIFICACAO-OS-${dataHora}.xlsx`);
    await calendarioHandler.exportarClassificacaoOS(
      resultadoAlocacao.ordensProcessadas,
      caminhoClassificacao
    );

    console.log('\n═════════════════════════════════════════════════════════════');
    console.log('✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!');
    console.log('═════════════════════════════════════════════════════════════');

    console.log('\n📁 Arquivos gerados:');
    console.log(`   1. ${path.basename(caminhoCalendarioSaida)}`);
    console.log(`      └─ Calendário PCP com OS alocadas`);
    console.log(`   2. ${path.basename(caminhoClassificacao)}`);
    console.log(`      └─ Classificação e priorização das OS${funcionarios ? ' + técnicos alocados' : ''}`);

    console.log(`\n📍 Localização: ${pastaOutput}`);

    console.log('\n💡 Próximos passos:');
    console.log('   1. Abrir o calendário preenchido');
    console.log('   2. Revisar as OS alocadas (cores indicam prioridade)');
    if (funcionarios) {
      console.log('   3. Verificar técnicos alocados para cada OS');
      console.log('   4. Conferir balanceamento de carga entre técnicos');
    }
    console.log('   5. Verificar OS pendentes na planilha de classificação');
    console.log('   6. Ajustar manualmente se necessário');
    console.log('   7. Comunicar programação para as equipes\n');

    const stats = resultadoAlocacao.estatisticas;
    const taxaAlocacao = (stats.alocadas / stats.total * 100).toFixed(1);

    console.log('📊 RESUMO EXECUTIVO:');
    console.log(`   - ${stats.total} OS processadas`);
    console.log(`   - ${stats.alocadas} OS programadas (${taxaAlocacao}%)`);

    if (funcionarios) {
      console.log(`   - ${stats.semTecnico} OS sem técnico disponível`);
      console.log(`   - ${funcionarios.length} técnicos no sistema`);
      
      const tecnicosComOS = funcionarios.filter(f => f.osAlocadas > 0).length;
      console.log(`   - ${tecnicosComOS} técnicos com OS alocadas`);
    }

    console.log(`   - ${stats.pendentes} OS aguardando slot`);
    console.log(`   - ${calendario.slots.length} slots disponíveis no calendário\n`);
  } catch (erro) {
    console.error('\n❌ ERRO NO PROCESSAMENTO:');
    console.error('═════════════════════════════════════════════════════════════');
    console.error(erro.message);
    console.error('\n📋 Stack trace:');
    console.error(erro.stack);
    console.error('\n💡 Possíveis soluções:');
    console.error('  - Verificar se os arquivos existem');
    console.error('  - Confirmar formato das planilhas');
    console.error('  - Verificar permissões de escrita na pasta output');
    console.error('  - Executar: npm install\n');
    
    process.exit(1);
  }
}

main();