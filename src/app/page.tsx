"use client";
import { useState } from 'react';
import { Upload, FileSpreadsheet, X, Loader2, CheckCircle2, Sparkles, Download, Calendar, ClipboardList } from 'lucide-react';
import styles from './XlsxUploader.module.css';

export default function XlsxUploader() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [loading1, setLoading1] = useState(false);
  const [loading2, setLoading2] = useState(false);
  const [processed1, setProcessed1] = useState(false);
  const [processed2, setProcessed2] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [resultado, setResultado] = useState(null);

  const handleFileChange = (e, fileNumber) => {
    const file = e.target.files?.[0];
    if (file && file.name.endsWith('.xlsx')) {
      if (fileNumber === 1) {
        setFile1(file);
        setProcessed1(false);
        simulateProcessing(1);
      } else {
        setFile2(file);
        setProcessed2(false);
        simulateProcessing(2);
      }
    }
  };

  const simulateProcessing = (fileNumber) => {
    if (fileNumber === 1) {
      setLoading1(true);
      setTimeout(() => {
        setLoading1(false);
        setProcessed1(true);
      }, 1000);
    } else {
      setLoading2(true);
      setTimeout(() => {
        setLoading2(false);
        setProcessed2(true);
      }, 1000);
    }
  };

  const handleProcessFiles = async () => {
    if (!file1 || !file2) return;

    setProcessing(true);
    setResultado(null);

    try {
      const formData = new FormData();
      formData.append('file1', file1);
      formData.append('file2', file2);

      const response = await fetch('/api/process-xlsx', {
        method: 'POST',
        body: formData,
      });

      const result = await response.json();

      if (response.ok) {
        setResultado(result.data);
      } else {
        alert(`Erro: ${result.error}\n${result.detalhes || ''}`);
      }
    } catch (error) {
      console.error('Erro ao processar:', error);
      alert('Erro ao processar as planilhas');
    } finally {
      setProcessing(false);
    }
  };

  const downloadArquivo = (nome, dadosBase64) => {
    const byteCharacters = atob(dadosBase64);
    const byteNumbers = new Array(byteCharacters.length);
    
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });

    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = nome;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  };

  const removeFile = (fileNumber) => {
    if (fileNumber === 1) {
      setFile1(null);
      setLoading1(false);
      setProcessed1(false);
    } else {
      setFile2(null);
      setLoading2(false);
      setProcessed2(false);
    }
  };

  const UploadZone = ({ fileNumber, file, loading, processed, title, subtitle, icon: Icon, color }) => (
    <div className={styles.uploadCard}>
      <div className={styles.cardHeader} style={{ background: `linear-gradient(135deg, ${color}15 0%, ${color}25 100%)` }}>
        <div className={styles.cardHeaderIcon} style={{ background: color }}>
          <Icon className={styles.cardHeaderIconSvg} />
        </div>
        <div className={styles.cardHeaderText}>
          <h3 className={styles.cardTitle}>{title}</h3>
          <p className={styles.cardSubtitle}>{subtitle}</p>
        </div>
      </div>

      <div className={styles.uploadZoneWrapper}>
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => handleFileChange(e, fileNumber)}
          className={styles.fileInput}
          id={`file-input-${fileNumber}`}
        />
        
        <div className={styles.glowEffect} style={{ background: color }} />
        
        <label
          htmlFor={`file-input-${fileNumber}`}
          className={`${styles.uploadLabel} ${file ? styles.uploadLabelActive : ''}`}
        >
          <div className={styles.uploadContent}>
            {loading ? (
              <>
                <div className={styles.loaderWrapper}>
                  <Loader2 className={styles.loaderIcon} style={{ color }} />
                  <div className={styles.loaderGlow} style={{ background: color }} />
                </div>
                <div className={styles.loadingText}>
                  <p className={styles.loadingTitle}>Processando planilha...</p>
                  <div className={styles.loadingDots}>
                    <div className={styles.dot} style={{ background: color, animationDelay: '0ms' }} />
                    <div className={styles.dot} style={{ background: color, animationDelay: '150ms' }} />
                    <div className={styles.dot} style={{ background: color, animationDelay: '300ms' }} />
                  </div>
                </div>
              </>
            ) : processed ? (
              <>
                <div className={styles.successIcon}>
                  <CheckCircle2 className={styles.checkIcon} />
                  <Sparkles className={styles.sparkleIcon} style={{ color }} />
                </div>
                <p className={styles.successText}>Planilha processada com sucesso!</p>
              </>
            ) : file ? (
              <>
                <div className={styles.fileIconWrapper}>
                  <FileSpreadsheet className={styles.fileIcon} style={{ color }} />
                  <div className={styles.fileIconGlow} style={{ background: color }} />
                </div>
                <div className={styles.fileInfo}>
                  <p className={styles.fileName}>{file.name}</p>
                  <p className={styles.fileSize}>{(file.size / 1024).toFixed(2)} KB</p>
                </div>
              </>
            ) : (
              <>
                <div className={styles.uploadIconWrapper}>
                  <Upload className={styles.uploadIcon} style={{ color }} />
                  <div className={styles.uploadIconGlow} style={{ background: color }} />
                </div>
                <div className={styles.uploadTextWrapper}>
                  <p className={styles.uploadTitle}>
                    Clique para selecionar ou arraste aqui
                  </p>
                  <p className={styles.uploadSubtitle}>
                    Formato .xlsx • Máximo 50MB
                  </p>
                </div>
              </>
            )}
          </div>
        </label>

        {file && !loading && (
          <button
            onClick={() => removeFile(fileNumber)}
            className={styles.removeButton}
          >
            <X className={styles.removeIcon} />
          </button>
        )}
      </div>
    </div>
  );

  const canProcess = file1 && file2 && !loading1 && !loading2 && !processing;

  return (
    <div className={styles.container}>
      <div className={styles.bgCircle1} />
      <div className={styles.bgCircle2} />
      <div className={styles.bgCircle3} />
      
      <div className={styles.content}>
        <div className={styles.header}>
          <div className={styles.headerIcon}>
            <FileSpreadsheet className={styles.headerIconSvg} />
          </div>
          <h1 className={styles.title}>
            Sistema de Gestão PCP
          </h1>
          <p className={styles.subtitle}>
            Planejamento e Controle de Produção Inteligente
          </p>
          <div className={styles.badge}>
            <Sparkles className={styles.badgeIcon} />
            <span className={styles.badgeText}>Rápido • Seguro • Profissional</span>
            <Sparkles className={styles.badgeIcon} />
          </div>
        </div>

        <div className={styles.uploadGrid}>
          <div className={styles.uploadColumn1}>
            <UploadZone
              fileNumber={1}
              file={file1}
              loading={loading1}
              processed={processed1}
              title="Calendário PCP"
              subtitle="Planilha de Planejamento e Controle"
              icon={Calendar}
              color="#3b82f6"
            />
          </div>
          <div className={styles.uploadColumn2}>
            <UploadZone
              fileNumber={2}
              file={file2}
              loading={loading2}
              processed={processed2}
              title="Solicitações de Serviço"
              subtitle="Ordens de Serviço e Demandas"
              icon={ClipboardList}
              color="#8b5cf6"
            />
          </div>
        </div>

        <div className={styles.buttonWrapper}>
          <button
            disabled={!canProcess}
            onClick={handleProcessFiles}
            className={`${styles.processButton} ${canProcess ? styles.processButtonActive : styles.processButtonDisabled}`}
          >
            <span className={styles.buttonOverlay} />
            <span className={styles.buttonContent}>
              {processing ? (
                <>
                  <Loader2 className={`${styles.buttonIcon} ${styles.spinning}`} />
                  Processando...
                </>
              ) : canProcess ? (
                <>
                  <Sparkles className={styles.buttonIcon} />
                  Processar Planilhas
                  <Sparkles className={styles.buttonIcon} />
                </>
              ) : (
                'Aguardando uploads...'
              )}
            </span>
          </button>
        </div>

        {resultado && (
          <div className={styles.resultCard}>
            <div className={styles.resultHeader}>
              <CheckCircle2 className={styles.resultHeaderIcon} />
              <h2 className={styles.resultTitle}>Processamento Concluído!</h2>
            </div>

            <div className={styles.statsGrid}>
              <div className={styles.statBox}>
                <span className={styles.statLabel}>Total de OS</span>
                <span className={styles.statValue}>{resultado.resumo?.total || 0}</span>
              </div>
              <div className={styles.statBox}>
                <span className={styles.statLabel}>OS Programadas</span>
                <span className={styles.statValue}>{resultado.resumo?.alocadas || 0}</span>
              </div>
              <div className={styles.statBox}>
                <span className={styles.statLabel}>Taxa de Sucesso</span>
                <span className={styles.statValue}>{resultado.resumo?.taxaSucesso || 0}%</span>
              </div>
              <div className={styles.statBox}>
                <span className={styles.statLabel}>OS Pendentes</span>
                <span className={styles.statValue}>{resultado.resumo?.pendentes || 0}</span>
              </div>
            </div>

            <div className={styles.downloadSection}>
              <h3 className={styles.downloadTitle}>
                <Download className={styles.downloadTitleIcon} />
                Baixar Resultados
              </h3>
              
              <button
                onClick={() => downloadArquivo(
                  resultado.arquivos.calendarioPreenchido.nome,
                  resultado.arquivos.calendarioPreenchido.dados
                )}
                className={styles.downloadButton}
              >
                <FileSpreadsheet className={styles.downloadButtonIcon} />
                <span>Calendário Preenchido</span>
                <Download className={styles.downloadButtonIconRight} />
              </button>

              <button
                onClick={() => downloadArquivo(
                  resultado.arquivos.classificacaoOS.nome,
                  resultado.arquivos.classificacaoOS.dados
                )}
                className={styles.downloadButton}
              >
                <FileSpreadsheet className={styles.downloadButtonIcon} />
                <span>Classificação de OS</span>
                <Download className={styles.downloadButtonIconRight} />
              </button>
            </div>
          </div>
        )}

        {(processed1 || processed2) && !resultado && (
          <div className={styles.statusCard}>
            <div className={styles.statusHeader}>
              <div className={styles.statusHeaderIcon}>
                <CheckCircle2 className={styles.statusHeaderIconSvg} />
              </div>
              <h3 className={styles.statusTitle}>Status do Processamento</h3>
            </div>
            <div className={styles.statusList}>
              <div className={styles.statusItem}>
                <span className={styles.statusLabel}>
                  <Calendar style={{ width: '20px', height: '20px', marginRight: '8px' }} />
                  Calendário PCP:
                </span>
                <span className={`${styles.statusValue} ${processed1 ? styles.statusValueSuccess : styles.statusValuePending}`}>
                  {processed1 && <CheckCircle2 className={styles.statusCheckIcon} />}
                  {processed1 ? 'Processado' : 'Aguardando'}
                </span>
              </div>
              <div className={styles.statusItem}>
                <span className={styles.statusLabel}>
                  <ClipboardList style={{ width: '20px', height: '20px', marginRight: '8px' }} />
                  Solicitações de Serviço:
                </span>
                <span className={`${styles.statusValue} ${processed2 ? styles.statusValueSuccess : styles.statusValuePending}`}>
                  {processed2 && <CheckCircle2 className={styles.statusCheckIcon} />}
                  {processed2 ? 'Processado' : 'Aguardando'}
                </span>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}