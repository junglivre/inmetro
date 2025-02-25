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
FTP_PORT = 21
FTP_USERNAME = 'USER'
FTP_PASSWORD = 'PW'
FTP_DIRECTORY = 'PATH'

# Diretório onde a Intelbras salva os vídeos
VIDEO_DIRECTORY = r'PATH'

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
    
    def get_relative_path(self, file_path):
        """Obtém o caminho relativo do arquivo em relação ao VIDEO_DIRECTORY"""
        return os.path.relpath(file_path, VIDEO_DIRECTORY)

    def file_exists_on_ftp(self, ftp, filename):
        """Verifica se o arquivo existe no servidor FTP"""
        try:
            ftp.size(filename)  # Tenta obter o tamanho do arquivo
            return True
        except ftplib.error_perm:  # Arquivo não existe
            return False
        except Exception as e:
            logging.error(f"Erro ao verificar existência do arquivo no FTP: {str(e)}")
            return False

    def connect_ftp(self):
        """Estabelece conexão com o servidor FTP"""
        try:
            ftp = ftplib.FTP()
            ftp.connect(FTP_SERVER, FTP_PORT, timeout=30)
            ftp.login(FTP_USERNAME, FTP_PASSWORD)
            return ftp
        except ftplib.error_perm as e:
            logging.error(f"Erro de permissão FTP: {str(e)}")
            if "530" in str(e):
                logging.error("Credenciais FTP inválidas")
            return None
        except socket.gaierror as e:
            logging.error(f"Erro de resolução do servidor FTP: {str(e)}")
            return None
        except socket.timeout as e:
            logging.error(f"Timeout na conexão FTP: {str(e)}")
            return None
        except ConnectionRefusedError as e:
            logging.error(f"Conexão recusada pelo servidor FTP: {str(e)}")
            return None
        except Exception as e:
            logging.error(f"Erro inesperado ao conectar ao FTP: {str(e)}")
            return None
    
    def create_remote_directory(self, ftp, directory):
        """Cria diretório remoto se não existir"""
        try:
            path_parts = directory.strip('/').split('/')
            
            try:
                ftp.cwd('/')
            except ftplib.error_perm as e:
                logging.error(f"Erro ao acessar diretório raiz: {str(e)}")
                return False
            
            current_path = ''
            for folder in path_parts:
                if folder:
                    current_path += '/' + folder
                    try:
                        ftp.cwd(current_path)
                    except ftplib.error_perm:
                        try:
                            ftp.mkd(current_path)
                            ftp.cwd(current_path)
                            logging.info(f"Criado diretório remoto: {current_path}")
                        except ftplib.error_perm as e:
                            logging.error(f"Erro ao criar diretório {current_path}: {str(e)}")
                            return False
            return True
        except Exception as e:
            logging.error(f"Erro inesperado ao criar diretório remoto: {str(e)}")
            return False
    
    def upload_file(self, file_path):
        """Faz upload de um arquivo para o servidor FTP"""
        if not os.path.exists(file_path):
            logging.warning(f"Arquivo não encontrado: {file_path}")
            return False
        
        if self.is_file_in_use(file_path):
            logging.warning(f"Arquivo está em uso: {file_path}")
            return False
        
        ftp = self.connect_ftp()
        if not ftp:
            return False
        
        try:
            # Obtém o caminho relativo e combina com o diretório FTP base
            relative_path = self.get_relative_path(file_path)
            remote_path = os.path.join(FTP_DIRECTORY, os.path.dirname(relative_path)).replace('\\', '/')
            file_name = os.path.basename(file_path)
            remote_file = os.path.join(remote_path, file_name).replace('\\', '/')
            
            # Cria a estrutura de diretórios no FTP
            if not self.create_remote_directory(ftp, remote_path):
                logging.error(f"Falha ao criar estrutura de diretórios: {remote_path}")
                return False
            
            # Navega até o diretório correto antes do upload
            try:
                ftp.cwd('/')  # Volta para a raiz
                ftp.cwd(remote_path)  # Navega até o diretório de destino
                logging.info(f"Navegado para diretório: {remote_path}")
            except ftplib.error_perm as e:
                logging.error(f"Erro ao navegar para diretório de destino: {str(e)}")
                return False
            
            # Verifica se o arquivo já existe no diretório atual
            file_exists = self.file_exists_on_ftp(ftp, file_name)
            
            if file_exists:
                # Se existe, aplica a regra de 3 horas
                file_mod_time = os.path.getmtime(file_path)
                file_age = datetime.datetime.now() - datetime.datetime.fromtimestamp(file_mod_time)
                
                if file_age.total_seconds() < MIN_AGE_HOURS * 3600:
                    return False
            
            # Se não existe ou já passou das 3 horas, faz o upload
            try:
                with open(file_path, 'rb') as file:
                    ftp.voidcmd('TYPE I')
                    # Upload do arquivo no diretório atual
                    ftp.storbinary(f'STOR {file_name}', file)
                
                logging.info(f"Upload concluído: {remote_file}")
                return True
                
            except ftplib.error_perm as e:
                logging.error(f"Erro de permissão durante upload: {str(e)}")
                if "553" in str(e):
                    logging.error("Permissão negada para escrita do arquivo")
                return False
            except IOError as e:
                logging.error(f"Erro de I/O durante upload: {str(e)}")
                return False
            
        except Exception as e:
            logging.error(f"Erro inesperado durante upload: {str(e)}")
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
        # Cria uma cópia da lista para evitar o erro de modificação durante iteração
        pending_files_copy = self.pending_files.copy()
        files_to_remove = set()
        
        for file_path in pending_files_copy:
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