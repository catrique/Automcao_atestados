import re

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
        self.endpoints = get_config('betha','api','endpoints')

        self.FILTROS = {
            "medico": 'filter=(nome+like+"%2525%2525"+and+profissao+=+"MEDICO")',
            "cid": 'filter=(codigo+like+"%2525%2525"+or+descricao+like+"%2525%2525")',
            "tipo_afastamento": 'filter=(descricao+like+"%2525%2525")',
            "tipo_atestado": 'filter=(descricao+like+"%2525%2525")',
            "motivo_consulta": 'filter=(descricao+like+"%2525%2525")',
            "pessoa_juridica": 'filter=(razaoSocial+like+"%2525%2525"+and+tipo+in+("GERAL","OPERADORA_PLANO_SAUDE"))',
            "listagem_matricula": 'filtroSituacao=ATIVOS'
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
        response = requests.get(url, headers=self.headers, timeout=60, proxies=PROXIES_OFF)
        response.raise_for_status()
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
    def servidores(self):
        raw_data = self.buscar_dados('listagem_matricula')
        
        dados_formatados = []
        for item in raw_data:
            pessoa = item.get('pessoa') or {}
            cargo = item.get('cargo') or {}
            matricula_info = item.get('matriculaLotacaoFisica') or {}
            lotacao_fisica = matricula_info.get('lotacaoFisica') or {}
            vinculo = item.get('vinculoEmpregaticio') or {}
            raw_vinculo = str(vinculo.get('descricao') or "NÃO INFORMADO")
            vinculo_limpo = re.sub(r'^\d+\s*-\s*', '', raw_vinculo).strip()
            linha = {
                "Id": pessoa.get('id'),
                "Matricula": item.get('numeroCartaoPonto') or item.get('descricao'),
                "Nome": pessoa.get('nome'),
                "Vínculo": vinculo_limpo,
                "Situação": item.get('situacao') or "NÃO INFORMADO",
                "CPF": pessoa.get('cpf'),
                "Cargo": cargo.get('descricao') or "NÃO INFORMADO",
                "Organograma": lotacao_fisica.get('descricao') or "SEM LOTAÇÃO",
                "Data_inicio": matricula_info.get('dataInicio')
            }
            dados_formatados.append(linha)
        
        return dados_formatados

updater = DataUpdater()

def sincronizar_bases_betha():
    """
    Agora retorna um OperationResult para o gui.py tratar.
    """
    tarefas = {
        "CID": updater.cids,
        "MEDICOS": updater.medicos,
        "TIPOS_AFASTAMENTO": updater.tipos_afastamento,
        "TIPOS_ATESTADO": updater.tipos_atestado,
        "MOTIVO_CONSULTA": updater.motivos_consulta,
        "EMPRESAS": updater.pessoas_juridicas,
        "SERVIDORES": updater.servidores
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
        erro= ErrorTranslator.traduzir(e)
        return OperationResult.fail(erro)