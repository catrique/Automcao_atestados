import requests
import os
import ctypes
import socket
from datetime import datetime
import logging
from config_network import PROXIES_OFF

class OperationResult:
    def __init__(self, success, message, data=None):
        self.success = success
        self.message = message
        self.data = data

    @staticmethod
    def ok(message, data=None):
        return OperationResult(True, message, data)

    @staticmethod
    def fail(message):
        return OperationResult(False, message)

class ErrorTranslator:
    @staticmethod
    def traduzir(e):
        err_str = str(e).lower()
        if isinstance(e, requests.exceptions.HTTPError):
            if e.response.status_code == 401:
                return "🔑 Erro de Autenticação no Betha. Verifique as credenciais."
            if e.response.status_code >= 500:
                return "🌐 O servidor da Betha está instável ou fora do ar."
        if isinstance(e, requests.exceptions.ConnectionError):
            return "🌐 Sem conexão com a internet ou servidor da Betha inacessível."
        if isinstance(e, requests.exceptions.Timeout):
            return "⏳ A Betha demorou muito para responder (Timeout)."
        return f"⚠️ Erro inesperado: {str(e)[:100]}"
    

def obter_identificacao_usuario():
    """Retorna um dicionário com Nome/Login, IP e Horário do sistema."""
    dados = {
        "usuario": "USUARIO",
        "ip": "0.0.0.0",
        "horario": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    }

    try:
        buffer = ctypes.create_unicode_buffer(100)
        tamanho = ctypes.pointer(ctypes.c_uint32(100))
        
        if ctypes.windll.secur32.GetUserNameExW(3, buffer, tamanho) and buffer.value:
            dados["usuario"] = buffer.value
        else:
            dados["usuario"] = os.getlogin()
    except:
        dados["usuario"] = os.environ.get('USERNAME', 'USUARIO')

    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        dados["ip"] = s.getsockname()[0]
        s.close()
    except:
        try:
            dados["ip"] = socket.gethostbyname(socket.gethostname())
        except:
            pass 

    return dados

logger = logging.getLogger("AutomacaoRH")
logger.setLevel(logging.INFO)

class GuiHandler(logging.Handler):
    """Handler que redireciona logs para a função da GUI."""
    def __init__(self, log_func):
        super().__init__()
        self.log_func = log_func

    def emit(self, record):
        msg = self.format(record)
        self.log_func(msg)

def configurar_log_gui(funcao_log_gui):
    """Ativa o envio de logs para a GUI."""
    gui_handler = GuiHandler(funcao_log_gui)
    gui_handler.setFormatter(logging.Formatter('%(message)s'))
    logger.addHandler(gui_handler)

console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
logger.addHandler(console_handler)