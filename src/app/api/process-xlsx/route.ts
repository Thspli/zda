// app/api/process-xlsx/route.ts
import { NextRequest, NextResponse } from 'next/server';
import { spawn } from 'child_process';
import path from 'path';
import fs from 'fs';
import { randomUUID } from 'crypto';
import * as XLSX from 'xlsx';

// ==================== UTILITÁRIOS ====================

function normalizarNomeColuna(nome: string): string {
  return nome
    .toString()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '_')
    .replace(/[^A-Z0-9_]/g, '')
    .trim();
}

function normalizarObjeto(obj: any): any {
  const normalizado: any = {};
  for (const [chave, valor] of Object.entries(obj)) {
    const chaveNormalizada = normalizarNomeColuna(chave);
    normalizado[chaveNormalizada] = valor;
  }
  return normalizado;
}

function lerExcelComHeaderCorreto(caminhoArquivo: string): any[] {
  const workbook = XLSX.readFile(caminhoArquivo);
  const primeiraSheet = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[primeiraSheet];

  let dados = XLSX.utils.sheet_to_json(worksheet);

  if (dados.length === 0) {
    return [];
  }

  const primeiraLinha = dados[0];
  const colunas = Object.keys(primeiraLinha);

  const temColunasInvalidas = colunas.some(col =>
    /^\d+$/.test(col) || col.includes('__EMPTY')
  );

  if (temColunasInvalidas) {
    console.log('⚠️ Detectado header na primeira linha de dados');

    const headerReal = Object.values(primeiraLinha).map(v => String(v));
    dados = dados.slice(1);

    dados = dados.map(linha => {
      const novaLinha: any = {};
      Object.values(linha).forEach((valor, index) => {
        const nomeColuna = headerReal[index] || `COLUNA_${index}`;
        novaLinha[nomeColuna] = valor;
      });
      return novaLinha;
    });
  }

  return dados.map(linha => normalizarObjeto(linha));
}

function salvarExcelComDados(caminho: string, dados: any[], nomeSheet: string = 'Dados'): void {
  const worksheet = XLSX.utils.json_to_sheet(dados);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, nomeSheet);
  XLSX.writeFile(workbook, caminho);
}

function buscarValor(obj: any, variacoes: string[]): any {
  for (const variacao of variacoes) {
    const chaveNormalizada = normalizarNomeColuna(variacao);
    if (obj[chaveNormalizada] !== undefined) {
      return obj[chaveNormalizada];
    }
  }
  return '';
}

// ==================== INTEGRAÇÃO COM IA ====================

interface ClassificacaoIA {
  tipo: string;
  confianca: number;
}

interface ResultadoIA {
  [osNum: string]: ClassificacaoIA;
}

async function chamarIAParaClassificar(caminhoArquivo: string): Promise<ResultadoIA> {
  const classificacoes: ResultadoIA = {};

  try {
    console.log('🤖 Chamando IA para classificar OS...');

    if (!fs.existsSync(caminhoArquivo)) {
      console.error('❌ Arquivo não encontrado');
      return classificacoes;
    }

    const stats = fs.statSync(caminhoArquivo);
    console.log(`📊 Enviando ${(stats.size / 1024).toFixed(2)} KB...`);

    // ✅ FORMA QUE VAI FUNCIONAR: ler e enviar como buffer com FormData
    const fileBuffer = fs.readFileSync(caminhoArquivo);
    const fileName = path.basename(caminhoArquivo);

    const FormData = (await import('form-data')).default;
    const formData = new FormData();

    // ✅ Importante: enviar como Buffer direto com options
    formData.append('arquivo', fileBuffer, {
      filename: fileName,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    console.log('📤 Enviando para IA...');

    // ======================================================
    //     INÍCIO DA ALTERAÇÃO SOLICITADA (LOGS + HEADERS)
    // ======================================================

    const response = await fetch('http://localhost:3000/processar', {
      method: 'POST',
      body: formData as any,
      headers: {
        ...formData.getHeaders(),
        'Connection': 'close'
      },
    });

    console.log(`📡 Status: ${response.status}`);
    if (!response.ok) {
      console.warn('❌ Erro:', response.status);
      const errorText = await response.text();
      console.log('Resposta de erro:', errorText);
      return classificacoes;
    }

    // ✅ LOG CRÍTICO AQUI
    const dados = await response.json();
    console.log('📊 DADOS COMPLETOS DA IA:', JSON.stringify(dados, null, 2).substring(0, 1000));
    console.log('📋 Campo linhas existe?', !!dados.linhas);
    console.log('📋 Quantidade de linhas:', dados.linhas?.length);

    if (!dados.linhas?.length) {
      console.warn('⚠️ Sem linhas na resposta');
      return classificacoes;
    }

    // ======================================================
    //     FIM DA ALTERAÇÃO SOLICITADA
    // ======================================================

    console.log(`📊 ${dados.linhas.length} linhas processadas`);

    dados.linhas.forEach((linha: any) => {
      const linhaNormalizada = normalizarObjeto(linha);
      const osNum = buscarValor(linhaNormalizada, ['OS', 'Ordem', 'Numero_OS']);

      if (osNum) {
        const tipo = buscarValor(linhaNormalizada, ['Classificacao', 'CLASSIFICACAO']) || 'N/A';
        const confianca = buscarValor(linhaNormalizada, ['Confianca', 'CONFIANCA']) || 0;

        classificacoes[String(osNum)] = {
          tipo: String(tipo),
          confianca: Number(confianca)
        };
      }
    });

    console.log(`✅ ${Object.keys(classificacoes).length} OS mapeadas com sucesso!\n`);
    return classificacoes;
  } catch (error: any) {
    console.error('❌ Erro:', error.message);
    return classificacoes;
  }
}

function adicionarClassificacaoIANoExcel(
  caminhoArquivo: string,
  classificacoes: ResultadoIA
): void {
  console.log('\n📝 Iniciando adição de classificações IA no Excel...');
  console.log(`📊 Total de classificações recebidas: ${Object.keys(classificacoes).length}`);
  
  if (Object.keys(classificacoes).length > 0) {
    const primeirasChaves = Object.keys(classificacoes).slice(0, 3);
    console.log('🔍 Exemplos de classificações:');
    primeirasChaves.forEach(chave => {
      console.log(`   OS ${chave}: ${classificacoes[chave].tipo} (${classificacoes[chave].confianca}%)`);
    });
  }

  if (Object.keys(classificacoes).length === 0) {
    console.log('⏭️ Nenhuma classificação IA para adicionar');
    return;
  }

  const dados = lerExcelComHeaderCorreto(caminhoArquivo);
  console.log(`📄 Total de linhas no arquivo: ${dados.length}`);

  if (dados.length === 0) {
    console.warn('⚠️ Arquivo Excel está vazio');
    return;
  }

  if (dados.length > 0) {
    console.log('📋 Colunas disponíveis na primeira linha:', Object.keys(dados[0]).slice(0, 10).join(', '));
  }

  let linhasComClassificacao = 0;
  const dadosAtualizados = dados.map((linha, index) => {
    const osNum = buscarValor(linha, ['OS', 'Ordem', 'ordem', 'Numero_OS', 'NumeroOS', 'NUMERO_OS']);
    const osNumString = String(osNum).trim();

    if (osNum && classificacoes[osNumString]) {
      linhasComClassificacao++;
      return {
        ...linha,
        CLASSIFICACAO_IA: classificacoes[osNumString].tipo,
        CONFIANCA_IA: classificacoes[osNumString].confianca + '%'
      };
    }

    return {
      ...linha,
      CLASSIFICACAO_IA: 'N/A',
      CONFIANCA_IA: 'N/A'
    };
  });

  salvarExcelComDados(caminhoArquivo, dadosAtualizados);
  console.log(`✅ ${linhasComClassificacao} linhas receberam classificação IA`);
  console.log(`✅ Colunas CLASSIFICACAO_IA e CONFIANCA_IA adicionadas com sucesso!\n`);
}

// ==================== PROCESSAMENTO PCM ====================

interface ResultadoPCM {
  resumo: {
    total: number;
    alocadas: number;
    pendentes: number;
    taxaSucesso: string;
  };
  estatisticas: {
    slotsDisponiveis: number;
    tempoProcessamento: string;
  };
}

function extrairNumero(texto: string, pattern: RegExp): number {
  const match = texto.match(pattern);
  return match ? parseInt(match[1], 10) : 0;
}

function parseResultadoPCM(stdout: string): ResultadoPCM {
  const linhas = stdout.split('\n');

  let osTotal = 0;
  let osAlocadas = 0;
  let osPendentes = 0;
  let slotsDisponiveis = 0;

  linhas.forEach(linha => {
    if (linha.includes('OS processadas')) {
      osTotal = extrairNumero(linha, /(\d+)\s+OS processadas/);
    }
    if (linha.includes('OS programadas') || linha.includes('OS alocadas')) {
      osAlocadas = extrairNumero(linha, /(\d+)\s+OS/);
    }
    if (linha.includes('OS aguardando slot') || linha.includes('OS pendentes')) {
      osPendentes = extrairNumero(linha, /(\d+)\s+OS/);
    }
    if (linha.includes('slots disponíveis') || linha.includes('slots disponiveis')) {
      slotsDisponiveis = extrairNumero(linha, /(\d+)\s+slots/);
    }
  });

  return {
    resumo: {
      total: osTotal,
      alocadas: osAlocadas,
      pendentes: osPendentes,
      taxaSucesso: osTotal > 0 ? ((osAlocadas / osTotal) * 100).toFixed(1) : '0'
    },
    estatisticas: {
      slotsDisponiveis,
      tempoProcessamento: 'Concluído'
    }
  };
}

function executarPCM(tempDir: string): Promise<ResultadoPCM> {
  return new Promise((resolve, reject) => {
    const scriptPath = path.join(process.cwd(), 'pcm', 'index.js');

    const processo = spawn('node', [scriptPath, tempDir], {
      cwd: process.cwd(),
      env: { ...process.env }
    });

    let stdout = '';
    let stderr = '';

    processo.stdout.on('data', (data) => {
      const texto = data.toString();
      stdout += texto;
      console.log(texto);
    });

    processo.stderr.on('data', (data) => {
      const texto = data.toString();
      stderr += texto;
      console.error(texto);
    });

    processo.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(`Processo PCM falhou com código ${code}\n${stderr}`));
        return;
      }
      
      const resultado = parseResultadoPCM(stdout);
      resolve(resultado);
    });

    processo.on('error', (error) => {
      reject(new Error(`Erro ao executar processo PCM: ${error.message}`));
    });
  });
}

// ==================== HANDLER PRINCIPAL ====================

export async function POST(request: NextRequest) {
  let tempDir: string | null = null;

  try {
    console.log('\n🚀 Iniciando processamento de planilhas...\n');

    const formData = await request.formData();
    const calendarioPcp = formData.get('file1') as File;
    const ordensServico = formData.get('file2') as File;

    if (!calendarioPcp || !ordensServico) {
      return NextResponse.json(
        { error: 'Ambos os arquivos são obrigatórios' },
        { status: 400 }
      );
    }

    const tempId = randomUUID();
    tempDir = path.join(process.cwd(), 'temp', tempId);
    const outputDir = path.join(tempDir, 'output');

    fs.mkdirSync(outputDir, { recursive: true });
    console.log(`📁 Diretório temporário criado: ${tempDir}`);

    const buffer1 = Buffer.from(await calendarioPcp.arrayBuffer());
    const buffer2 = Buffer.from(await ordensServico.arrayBuffer());

    const caminhoCalendario = path.join(tempDir, 'calendario-pcp.xlsx');
    const caminhoOS = path.join(tempDir, 'ordens-servico.xlsx');

    fs.writeFileSync(caminhoCalendario, buffer1);
    fs.writeFileSync(caminhoOS, buffer2);
    console.log('✅ Arquivos salvos no disco com sucesso');

    // Chamar IA passando o CAMINHO do arquivo salvo no disco (string)
    console.log('\n📤 Enviando arquivo para classificação IA...');
    const classificacoes = await chamarIAParaClassificar(caminhoOS);
    console.log(`✅ Recebidas ${Object.keys(classificacoes).length} classificações da IA\n`);

    const rootPath = process.cwd();
    const arquivosFixos = [
      {
        origem: path.join(rootPath, 'Controle-Bens-SENAI-SPRINT-1.xlsx'),
        destino: path.join(tempDir, 'Controle-Bens-SENAI-SPRINT-1.xlsx'),
        obrigatorio: true
      },
      {
        origem: path.join(rootPath, 'funcionarios.xlsx'),
        destino: path.join(tempDir, 'funcionarios.xlsx'),
        obrigatorio: false
      }
    ];

    for (const arquivo of arquivosFixos) {
      if (fs.existsSync(arquivo.origem)) {
        fs.copyFileSync(arquivo.origem, arquivo.destino);
        console.log(`✅ Copiado: ${path.basename(arquivo.origem)}`);
      } else if (arquivo.obrigatorio) {
        return NextResponse.json(
          { error: `Arquivo ${path.basename(arquivo.origem)} não encontrado na raiz do projeto` },
          { status: 400 }
        );
      }
    }

    console.log('\n⚙️ Executando processo PCM...\n');
    const resultado = await executarPCM(tempDir);
    console.log('\n✅ Processo PCM concluído com sucesso\n');

    const arquivosSaida = fs.readdirSync(outputDir);
    const calendarioPreenchido = arquivosSaida.find(f => 
      f.startsWith('CALENDARIO-PREENCHIDO') || f.includes('calendario')
    );
    const classificacaoOS = arquivosSaida.find(f => 
      f.startsWith('CLASSIFICACAO-OS') || f.includes('classificacao')
    );

    if (!calendarioPreenchido || !classificacaoOS) {
      throw new Error('Arquivos de saída não foram gerados pelo processo PCM');
    }

    console.log(`📄 Arquivo de classificação gerado: ${classificacaoOS}`);

    // Adicionar classificações IA no arquivo de saída
    const caminhoClassificacaoFull = path.join(outputDir, classificacaoOS);
    adicionarClassificacaoIANoExcel(caminhoClassificacaoFull, classificacoes);

    const calendarioBuffer = fs.readFileSync(path.join(outputDir, calendarioPreenchido));
    const classificacaoBuffer = fs.readFileSync(caminhoClassificacaoFull);

    console.log('✅ Arquivos de saída preparados para download\n');

    const response = {
      success: true,
      message: 'Planilhas processadas com sucesso!',
      data: {
        resumo: resultado.resumo,
        estatisticas: resultado.estatisticas,
        arquivos: {
          calendarioPreenchido: {
            nome: calendarioPreenchido,
            dados: calendarioBuffer.toString('base64')
          },
          classificacaoOS: {
            nome: classificacaoOS,
            dados: classificacaoBuffer.toString('base64')
          }
        }
      }
    };

    // Limpar arquivos temporários após 5 segundos
    setTimeout(() => {
      if (tempDir && fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
        console.log('🧹 Arquivos temporários removidos');
      }
    }, 5000);

    return NextResponse.json(response);
  } catch (error: any) {
    console.error('\n❌ Erro ao processar planilhas:', error.message);
    console.error('❌ Stack:', error.stack);

    if (tempDir && fs.existsSync(tempDir)) {
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch (cleanupError) {
        console.error('Erro ao limpar arquivos temporários:', cleanupError);
      }
    }

    return NextResponse.json(
      {
        error: 'Erro ao processar as planilhas',
        detalhes: error.message
      },
      { status: 500 }
    );
  }
}
