import requests
import sys
import os
import time
from config_global import sheets
from config_network import PROXIES_OFF
from services.utils_service import OperationResult, ErrorTranslator 
from services.utils_service import logger

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config.loaders import reload_settings, get_config


class DataUpdater:
    def __init__(self):
        self.base_url = get_config('betha','api','base_url')
        # self.auth_token = get_config('betha','api','authorization')
        # self.user_access = get_config('betha','api','user_access')
        self.endpoints = get_config('betha','api','endpoints')

        # self.headers = {
        #     "Authorization": self.auth_token,
        #     "user-access": self.user_access,
        #     "Content-Type": "application/json"
        # }

        self.FILTROS = {
            "medico": 'filter=(nome+like+"%2525%2525"+and+profissao+=+"MEDICO")',
            "cid": 'filter=(codigo+like+"%2525%2525"+or+descricao+like+"%2525%2525")',
            "tipo_afastamento": 'filter=(descricao+like+"%2525%2525")',
            "tipo_atestado": 'filter=(descricao+like+"%2525%2525")',
            "motivo_consulta": 'filter=(descricao+like+"%2525%2525")',
            "pessoa_juridica": 'filter=(razaoSocial+like+"%2525%2525"+and+tipo+in+("GERAL","OPERADORA_PLANO_SAUDE"))'
        }

    @property
    def headers(self):
        """Toda vez que alguém acessar 'self.headers', ele lerá o token novo."""
        reload_settings() 
        
        return {
            "Authorization": get_config('betha', 'api', 'authorization'),
            "user-access": get_config('betha', 'api', 'user_access'),
            "Content-Type": "application/json"
        }

    def _executar_requisicao(self, url):
        """Faz a chamada GET e lança exceções para o tradutor capturar."""
        response = requests.get(url, headers=self.headers, timeout=3, proxies=PROXIES_OFF)
        response.raise_for_status() # Isso dispara o erro para o try/except pai
        return response.json()

    def buscar_dados(self, chave_endpoint):
        path = self.endpoints.get(chave_endpoint)
        filtro = self.FILTROS.get(chave_endpoint, 'limit=100') 
        
        if not path:
            return []

        todos_registros = []
        offset = 0
        limit = 1000
        has_next = True

        while has_next:
            conector = "&" if "?" in path or "?" in filtro else "?"
            url = f"{self.base_url}{path}{filtro}{conector}limit={limit}&offset={offset}"
            
            dados = self._executar_requisicao(url)
            
            if not dados or 'content' not in dados:
                break
                
            todos_registros.extend(dados.get('content', []))
            has_next = dados.get('hasNext', False)
            
            offset += limit
            if has_next: time.sleep(0.3) 

        return todos_registros

    def medicos(self): return self.buscar_dados('medico')
    def cids(self): return self.buscar_dados('cid')
    def tipos_afastamento(self): return self.buscar_dados('tipo_afastamento')
    def tipos_atestado(self): return self.buscar_dados('tipo_atestado')
    def motivos_consulta(self): return self.buscar_dados('motivo_consulta')
    def pessoas_juridicas(self): return self.buscar_dados('pessoa_juridica')

updater = DataUpdater()

def sincronizar_bases_betha():
    """
    Agora retorna um OperationResult para o gui.py tratar.
    """
    NOME_PLANILHA = "Atestados"
    tarefas = {
        "CID": updater.cids,
        "MEDICOS": updater.medicos,
        "TIPOS_AFASTAMENTO": updater.tipos_afastamento,
        "TIPOS_ATESTADO": updater.tipos_atestado,
        "MOTIVO_CONSULTA": updater.motivos_consulta,
        "EMPRESAS": updater.pessoas_juridicas
    }

    try:
        for aba, metodo in tarefas.items():
            logger.info(f"🔄 Sincronizando aba: {aba}...")
            dados = metodo()
            
            if dados:
                sheets.atualizar_aba_com_json(aba, dados)
            else:
                logger.info(f"⚠️ {aba} sem dados.")
        
        return OperationResult.ok("✅ Todas as bases foram sincronizadas com sucesso!")

    except Exception as e:
        mensagem_amigavel = ErrorTranslator.traduzir(e)
        return OperationResult.fail(mensagem_amigavel)