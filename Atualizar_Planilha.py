import time
import logging
import win32com.client as win32
from pathlib import Path
from datetime import datetime

# --- Configuração de Observabilidade (Logs) ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger("ExcelOrchestrator")

def executar_pipeline_atualizacao(caminho_arquivo: str):
    """
    Executa o ciclo de vida de atualização do Excel com buffers de tempo
    para garantir estabilidade em ambientes de alta latência.
    """
    
    # 1. Validação de entrada (Defensive Programming)
    path_obj = Path(caminho_arquivo)
    if not path_obj.exists():
        logger.error(f"❌ Arquivo não encontrado: {path_obj}")
        return

    excel = None
    workbook = None

    logger.info(f"🚀 Iniciando pipeline para: {path_obj.name}")

    try:
        # 2. Inicialização da Instância COM
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # 3. Abertura do Arquivo
        logger.info("📂 Abrindo planilha...")
        workbook = excel.Workbooks.Open(str(path_obj.resolve()))
        
        # [BUFFER DE ESTABILIDADE 1]
        # Garante que o arquivo foi completamente carregado na memória/rede
        logger.info("⏳ Aguardando carregamento completo (5s)...")
        time.sleep(5)

        # 4. Atualização de Dados (ETL)
        logger.info("🔄 Executando RefreshAll...")
        workbook.RefreshAll()
        
        # Sincronização Híbrida: Método Nativo + Buffer
        excel.CalculateUntilAsyncQueriesDone()
        
        # 5. Persistência
        logger.info("💾 Salvando alterações...")
        workbook.Save()

        # [BUFFER DE ESTABILIDADE 2]
        # Garante que o I/O do disco finalizou a gravação antes de fechar
        logger.info("⏳ Aguardando commit no disco (5s)...")
        time.sleep(5)

        workbook.Close(SaveChanges=False) # Já salvamos antes
        workbook = None # Marca como fechado para o bloco finally
        
        logger.info("✅ Planilha salva e fechada com sucesso!")

    except Exception as e:
        logger.error(f"💥 Falha crítica no processo: {e}")

    finally:
        # 6. Limpeza de Recursos (Garbage Collection Manual)
        logger.info("🧹 Iniciando limpeza de processos...")
        
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
        
        if excel:
            try:
                excel.Quit()
            except:
                pass
            
        # Libera os objetos COM da memória do Windows
        del workbook
        del excel
        
        logger.info(f"🏁 Processo para {path_obj.name} finalizado.")

    # [BUFFER DE ESTABILIDADE 3]
    # Pausa final para garantir que o processo do Excel sumiu do Task Manager
    # antes de uma próxima execução ou fim do script.
    logger.info("⏳ Cooldown final do sistema (5s)...")
    time.sleep(5)

if __name__ == "__main__":
    # Utilize r-strings para caminhos Windows
    ARQUIVO_ALVO = r"C:\Caminho\Para\Sua\Planilha.xlsx"
    
    executar_pipeline_atualizacao(ARQUIVO_ALVO)
