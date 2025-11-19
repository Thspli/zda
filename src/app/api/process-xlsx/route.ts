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
const primeiraSheet = workbook.SheetNames[0]; // ✅ Pega o primeiro nome (string)
const worksheet = workbook.Sheets[primeiraSheet]; // ✅ Agora funciona

  let dados = XLSX.utils.sheet_to_json(worksheet);

  if (dados.length === 0) {
    return [];
  }

  const primeiraLinha = dados;
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
  return match ? parseInt(match, 10) : 0;
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

    const calendarioBuffer = fs.readFileSync(path.join(outputDir, calendarioPreenchido));
    const classificacaoBuffer = fs.readFileSync(path.join(outputDir, classificacaoOS));

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
