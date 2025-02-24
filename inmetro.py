import os
import sys
import time
import datetime
import logging
import ftplib
import shutil
import win32file
import win32con
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Configuração de logging
logging.basicConfig(
    filename='ftp_upload.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Configurações FTP
FTP_SERVER = 'SERVER'
FTP_USERNAME = 'USER'
FTP_PASSWORD = 'PW'
FTP_DIRECTORY = 'PATH'  # Diretório no servidor FTP

# Diretório onde a Intelbras salva os vídeos
VIDEO_DIRECTORY = r'PATH'  # Ajuste para o caminho correto

# Diretório para arquivos já enviados (opcional)
SENT_DIRECTORY = r'PATH'

# Tempo mínimo (em horas) que um arquivo deve ter desde a última modificação
MIN_AGE_HOURS = 3

class FTPUploader:
    def __init__(self):
        self.ensure_directories_exist()
    
    def ensure_directories_exist(self):
        """Garante que os diretórios necessários existam"""
        if not os.path.exists(SENT_DIRECTORY):
            os.makedirs(SENT_DIRECTORY)
    
    def connect_ftp(self):
        """Estabelece conexão com o servidor FTP"""
        try:
            ftp = ftplib.FTP(FTP_SERVER)
            ftp.login(FTP_USERNAME, FTP_PASSWORD)
            return ftp
        except Exception as e:
            logging.error(f"Erro ao conectar ao servidor FTP: {e}")
            return None
    
    def create_remote_directory(self, ftp, directory):
        """Cria diretório remoto se não existir"""
        try:
            ftp.cwd('/')  # Volta para o diretório raiz
            for folder in directory.strip('/').split('/'):
                if folder:
                    try:
                        ftp.cwd(folder)
                    except:
                        ftp.mkd(folder)
                        ftp.cwd(folder)
        except Exception as e:
            logging.error(f"Erro ao criar diretório remoto: {e}")
    
    def upload_file(self, file_path):
        """Faz upload de um arquivo para o servidor FTP"""
        # Verificar se o arquivo existe
        if not os.path.exists(file_path):
            logging.warning(f"Arquivo não encontrado: {file_path}")
            return False
        
        # Verificar a idade do arquivo
        file_mod_time = os.path.getmtime(file_path)
        file_age = datetime.datetime.now() - datetime.datetime.fromtimestamp(file_mod_time)
        
        # Se o arquivo foi modificado nas últimas 3 horas, não fazer upload ainda
        if file_age.total_seconds() < MIN_AGE_HOURS * 3600:
            logging.info(f"Arquivo muito recente, ignorando por enquanto: {file_path}")
            return False
        
        # Verificar se o arquivo está em uso
        if self.is_file_in_use(file_path):
            logging.warning(f"Arquivo está em uso, ignorando por enquanto: {file_path}")
            return False
        
        # Conectar ao FTP
        ftp = self.connect_ftp()
        if not ftp:
            return False
        
        try:
            # Criar estrutura de diretórios no servidor FTP
            today = datetime.datetime.now().strftime('%Y-%m-%d')
            remote_dir = f"{FTP_DIRECTORY}{today}/"
            self.create_remote_directory(ftp, remote_dir)
            ftp.cwd(remote_dir)
            
            # Fazer upload do arquivo
            file_name = os.path.basename(file_path)
            with open(file_path, 'rb') as file:
                ftp.storbinary(f'STOR {file_name}', file)
            
            logging.info(f"Arquivo enviado com sucesso: {file_name}")
            
            # Mover arquivo para pasta de enviados
            self.move_to_sent(file_path)
            
            return True
        except Exception as e:
            logging.error(f"Erro ao enviar arquivo {file_path}: {e}")
            return False
        finally:
            try:
                ftp.quit()
            except:
                pass
    
    def is_file_in_use(self, file_path):
        """Verifica se o arquivo está em uso"""
        try:
            # Tenta abrir o arquivo em modo exclusivo
            handle = win32file.CreateFile(
                file_path, 
                win32con.GENERIC_READ, 
                0,  # Não permite compartilhamento
                None, 
                win32con.OPEN_EXISTING, 
                win32con.FILE_ATTRIBUTE_NORMAL, 
                None
            )
            win32file.CloseHandle(handle)
            return False  # Arquivo não está em uso
        except:
            return True  # Arquivo está em uso
    
    def move_to_sent(self, file_path):
        """Move o arquivo para a pasta de enviados"""
        try:
            if os.path.exists(SENT_DIRECTORY):
                today = datetime.datetime.now().strftime('%Y-%m-%d')
                sent_subdir = os.path.join(SENT_DIRECTORY, today)
                
                if not os.path.exists(sent_subdir):
                    os.makedirs(sent_subdir)
                    
                file_name = os.path.basename(file_path)
                destination = os.path.join(sent_subdir, file_name)
                shutil.move(file_path, destination)
                logging.info(f"Arquivo movido para: {destination}")
        except Exception as e:
            logging.error(f"Erro ao mover arquivo para pasta de enviados: {e}")

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self):
        self.uploader = FTPUploader()
        # Conjunto para armazenar arquivos que ainda precisam ser processados
        self.pending_files = set()
    
    def on_created(self, event):
        """Chamado quando um arquivo é criado"""
        if not event.is_directory:
            file_path = event.src_path
            logging.info(f"Novo arquivo detectado: {file_path}")
            self.process_file(file_path)
    
    def on_modified(self, event):
        """Chamado quando um arquivo é modificado"""
        if not event.is_directory:
            file_path = event.src_path
            logging.info(f"Arquivo modificado: {file_path}")
            self.process_file(file_path)
    
    def process_file(self, file_path):
        """Processa um arquivo para upload"""
        # Verificar se é um arquivo de vídeo
        if file_path.lower().endswith(('.mp4', '.avi', '.mkv')):
            # Adicionar à lista de arquivos pendentes
            self.pending_files.add(file_path)
    
    def process_pending_files(self):
        """Processa todos os arquivos pendentes"""
        files_to_remove = set()
        
        for file_path in self.pending_files:
            # Tentar fazer upload do arquivo
            success = self.uploader.upload_file(file_path)
            if success:
                files_to_remove.add(file_path)
        
        # Remover os arquivos processados com sucesso da lista de pendentes
        self.pending_files -= files_to_remove

class WatchDogService(win32serviceutil.ServiceFramework):
    _svc_name_ = "FTPWatchDogService"
    _svc_display_name_ = "Serviço de Monitoramento e Upload FTP"
    _svc_description_ = "Monitora a pasta de vídeos e faz upload automático via FTP"
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.is_running = False
    
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_running = False
    
    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        self.is_running = True
        self.main()
    
    def main(self):
        # Inicializar o handler e o observer
        event_handler = FileChangeHandler()
        observer = Observer()
        observer.schedule(event_handler, path=VIDEO_DIRECTORY, recursive=False)
        observer.start()
        
        try:
            # Processar arquivos existentes na inicialização
            for file_name in os.listdir(VIDEO_DIRECTORY):
                file_path = os.path.join(VIDEO_DIRECTORY, file_name)
                if os.path.isfile(file_path) and file_path.lower().endswith(('.mp4', '.avi', '.mkv')):
                    event_handler.process_file(file_path)
            
            # Loop principal do serviço
            while self.is_running:
                # Verificar a cada minuto se há arquivos pendentes para processar
                event_handler.process_pending_files()
                
                # Verificar se o serviço deve parar
                if win32event.WaitForSingleObject(self.hWaitStop, 60000) == win32event.WAIT_OBJECT_0:
                    break
        
        finally:
            observer.stop()
            observer.join()

def run_as_console():
    """Executar o script como aplicativo de console (para teste)"""
    # Configurar logging no console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    logging.getLogger().addHandler(console_handler)
    
    print("Iniciando monitoramento de arquivos...")
    
    # Inicializar o handler e o observer
    event_handler = FileChangeHandler()
    observer = Observer()
    observer.schedule(event_handler, path=VIDEO_DIRECTORY, recursive=False)
    observer.start()
    
    try:
        # Processar arquivos existentes na inicialização
        for file_name in os.listdir(VIDEO_DIRECTORY):
            file_path = os.path.join(VIDEO_DIRECTORY, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith(('.mp4', '.avi', '.mkv')):
                event_handler.process_file(file_path)
        
        # Loop principal
        while True:
            # Verificar a cada minuto se há arquivos pendentes para processar
            event_handler.process_pending_files()
            time.sleep(60)
    
    except KeyboardInterrupt:
        observer.stop()
    
    observer.join()

if __name__ == "__main__":
    if len(sys.argv) == 1:
        # Sem argumentos - determinar se executa como serviço ou console
        if sys.stdout.isatty():
            # Rodando em um terminal - modo console
            run_as_console()
        else:
            # Não tem terminal - provavelmente rodando como serviço
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(WatchDogService)
            servicemanager.StartServiceCtrlDispatcher()
    else:
        # Com argumentos - processar comandos do serviço
        win32serviceutil.HandleCommandLine(WatchDogService)