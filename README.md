# Instruções de Instalação e Configuração

## Requisitos

1. Python 3.6 ou superior
2. Bibliotecas Python necessárias:
   - pywin32 (para serviço Windows)
   - watchdog (para monitoramento de arquivos)

## Passo a Passo

### 1. Instalar as bibliotecas necessárias

```bash
pip install pywin32 watchdog
```

### 2. Editar o script

Edite as seguintes variáveis no script:

- `FTP_SERVER`: endereço do seu servidor FTP
- `FTP_USERNAME`: nome do usuário FTP
- `FTP_PASSWORD`: senha do usuário FTP
- `FTP_DIRECTORY`: diretório remoto onde os arquivos serão salvos
- `VIDEO_DIRECTORY`: diretório local onde os arquivos da câmera Intelbras são salvos
- `SENT_DIRECTORY`: diretório local para mover os arquivos após o upload
- `MIN_AGE_HOURS`: tempo mínimo (em horas) desde a última modificação do arquivo para permitir o upload

### 3. Teste o script no modo console

Execute o script normalmente para verificar se funciona:

```bash
python script_ftp_watchdog.py
```

Isso deve iniciar o monitoramento no modo console.

### 4. Instalar como serviço

Depois de confirmar que o script funciona corretamente, instale-o como um serviço:

```bash
python script_ftp_watchdog.py install
```

### 5. Configurar o serviço para iniciar automaticamente

1. Abra o "Serviços" do Windows (services.msc)
2. Localize o serviço "Serviço de Monitoramento e Upload FTP"
3. Dê duplo clique para abrir as propriedades
4. Mude o "Tipo de inicialização" para "Automático"
5. Clique em "OK"

### 6. Iniciar o serviço

```bash
python script_ftp_watchdog.py start
```

Ou inicie pelo gerenciador de serviços do Windows.

## Verificação

1. Verifique o arquivo de log (`ftp_upload.log`) para confirmar que o serviço está funcionando corretamente
2. Você pode monitorar os diretórios para verificar se os arquivos estão sendo movidos após o upload

## Comandos do Serviço

- Iniciar: `python script_ftp_watchdog.py start`
- Parar: `python script_ftp_watchdog.py stop`
- Reiniciar: `python script_ftp_watchdog.py restart`
- Desinstalar: `python script_ftp_watchdog.py remove`
- Ver status: `python script_ftp_watchdog.py status`