// app/api/process-xlsx/route.ts
import { NextRequest, NextResponse } from 'next/server';
import { spawn } from 'child_process';
import path from 'path';
import fs from 'fs';
import { randomUUID } from 'crypto';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const calendarioPcp = formData.get('file1') as File;
    const ordensServico = formData.get('file2') as File;

    if (!calendarioPcp || !ordensServico) {
      return NextResponse.json(
        { error: 'Ambos os arquivos são obrigatórios' },
        { status: 400 }
      );
    }

    // Criar pasta temporária única para este processamento
    const tempId = randomUUID();
    const tempDir = path.join(process.cwd(), 'temp', tempId);
    fs.mkdirSync(tempDir, { recursive: true });

    // Salvar arquivos temporariamente
    const buffer1 = Buffer.from(await calendarioPcp.arrayBuffer());
    const buffer2 = Buffer.from(await ordensServico.arrayBuffer());

    const caminhoCalendario = path.join(tempDir, 'calendario-pcp.xlsx');
    const caminhoOS = path.join(tempDir, 'ordens-servico.xlsx');

    fs.writeFileSync(caminhoCalendario, buffer1);
    fs.writeFileSync(caminhoOS, buffer2);

    // COPIAR arquivos fixos do sistema (se existirem)
    const rootPath = process.cwd();
    const controleBemsOriginal = path.join(rootPath, 'Controle-Bens-SENAI-SPRINT-1.xlsx');
    const funcionariosOriginal = path.join(rootPath, 'funcionarios.xlsx');

    if (fs.existsSync(controleBemsOriginal)) {
      const controleBemsTemp = path.join(tempDir, 'Controle-Bens-SENAI-SPRINT-1.xlsx');
      fs.copyFileSync(controleBemsOriginal, controleBemsTemp);
    } else {
      return NextResponse.json(
        { error: 'Arquivo Controle-Bens-SENAI-SPRINT-1.xlsx não encontrado na raiz do projeto' },
        { status: 400 }
      );
    }

    if (fs.existsSync(funcionariosOriginal)) {
      const funcionariosTemp = path.join(tempDir, 'funcionarios.xlsx');
      fs.copyFileSync(funcionariosOriginal, funcionariosTemp);
    }

    // Criar pasta output dentro do temp
    const outputDir = path.join(tempDir, 'output');
    fs.mkdirSync(outputDir, { recursive: true });

    // Executar o algoritmo PCM
    const resultado = await executarPCM(tempDir);

    // Ler arquivos de saída
    const arquivosSaida = fs.readdirSync(outputDir);
    const calendarioPreenchido = arquivosSaida.find(f => f.startsWith('CALENDARIO-PREENCHIDO'));
    const classificacaoOS = arquivosSaida.find(f => f.startsWith('CLASSIFICACAO-OS'));

    if (!calendarioPreenchido || !classificacaoOS) {
      throw new Error('Arquivos de saída não foram gerados');
    }

    // Converter para base64 para enviar ao cliente
    const calendarioBuffer = fs.readFileSync(path.join(outputDir, calendarioPreenchido));
    const classificacaoBuffer = fs.readFileSync(path.join(outputDir, classificacaoOS));

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

    // Limpar arquivos temporários
    setTimeout(() => {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }, 5000);

    return NextResponse.json(response);

  } catch (error: any) {
    console.error('Erro ao processar planilhas:', error);
    return NextResponse.json(
      { 
        error: 'Erro ao processar as planilhas',
        detalhes: error.message 
      },
      { status: 500 }
    );
  }
}

/**
 * Executa o algoritmo PCM em Node.js
 * ✅ CORRIGIDO: Roda da raiz e passa tempDir como argumento
 */
function executarPCM(tempDir: string): Promise<any> {
  return new Promise((resolve, reject) => {
    const scriptPath = path.join(process.cwd(), 'pcm', 'index.js');

    // ✅ CORRIGIDO: Passa tempDir como argumento e roda da raiz
    const processo = spawn('node', [scriptPath, tempDir], {
      cwd: process.cwd(),  // ← Roda da raiz (onde está node_modules)
      env: { ...process.env }
    });

    let stdout = '';
    let stderr = '';

    processo.stdout.on('data', (data) => {
      stdout += data.toString();
      console.log(data.toString());
    });

    processo.stderr.on('data', (data) => {
      stderr += data.toString();
      console.error(data.toString());
    });

    processo.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(`Processo PCM falhou com código ${code}\n${stderr}`));
        return;
      }

      // Extrair resumo e estatísticas do stdout
      const resultado = parseResultado(stdout);
      resolve(resultado);
    });

    processo.on('error', (error) => {
      reject(new Error(`Erro ao executar processo PCM: ${error.message}`));
    });
  });
}

/**
 * Parse do output do console para extrair informações
 */
function parseResultado(stdout: string) {
  const linhas = stdout.split('\n');
  
  let osAlocadas = 0;
  let osTotal = 0;
  let osPendentes = 0;
  let slotsDisponiveis = 0;

  linhas.forEach(linha => {
    if (linha.includes('OS processadas')) {
      const match = linha.match(/(\d+) OS processadas/);
      if (match) osTotal = parseInt(match[1]);
    }
    if (linha.includes('OS programadas')) {
      const match = linha.match(/(\d+) OS programadas/);
      if (match) osAlocadas = parseInt(match[1]);
    }
    if (linha.includes('OS aguardando slot') || linha.includes('OS pendentes')) {
      const match = linha.match(/(\d+) OS/);
      if (match) osPendentes = parseInt(match[1]);
    }
    if (linha.includes('slots disponíveis')) {
      const match = linha.match(/(\d+) slots/);
      if (match) slotsDisponiveis = parseInt(match[1]);
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
