import os
import sys
import time
import gspread
import pandas as pd
from google.auth.transport.requests import Request
import unicodedata
from oauth2client.service_account import ServiceAccountCredentials
from config.loaders import get_config
from services.utils_service import OperationResult, ErrorTranslator
from services.utils_service import obter_identificacao_usuario
from services.utils_service import logger
from config_network import PROXIES_OFF
NOME_ABA = get_config("google_sheets", "aba")

class SheetsService:
    def __init__(self, nome_planilha):
        if getattr(sys, 'frozen', False):
            self.RAIZ_PROJETO = sys._MEIPASS
        else:
            diretorio_atual = os.path.dirname(os.path.abspath(__file__))
            self.RAIZ_PROJETO = os.path.dirname(diretorio_atual)

        self.PATH_TO_JSON = os.path.join(self.RAIZ_PROJETO, "config", "credentials.json")
        
        logger.info(f"🔍 Conectando à planilha via: {self.PATH_TO_JSON}")
        self.planilha = self._conectar(nome_planilha)
        self._cache_abas = {}


    def _conectar(self, nome_planilha):
        import gspread
        import os
        from google.oauth2.service_account import Credentials 
        from google.auth.transport.requests import Request
        from config_network import sessao_limpa

        creds = Credentials.from_service_account_file(
            self.PATH_TO_JSON, 
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )

        autenticador_google = Request(session=sessao_limpa)
        
        client = gspread.authorize(creds)
        client.session = sessao_limpa
        
        return client.open(nome_planilha)

    def _limpar_texto(self, texto):
        if not texto: return ""
        nfkd = unicodedata.normalize('NFKD', str(texto))
        sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
        return sem_acento.replace(".", "").strip().upper()

    def achatar_json(self, dados_json):
        df = pd.json_normalize(dados_json)
        df.columns = [c.replace('.', '_') for c in df.columns]
        return df

    def obter_aba(self, aba = NOME_ABA):
        try:
            return self.planilha.worksheet(aba)
        except:
            return None

    def ler_planilha_para_automacao(self, linha_inicio):
        """Lê a aba a partir de uma linha específica para uso no Selenium."""
        try:
            aba = self.obter_aba()
            cabecalho = aba.row_values(1)
            dados = aba.get_all_values()[linha_inicio-1:] 

            lista_final = []
            for valores_linha in dados:
                linha_dict = dict(zip(cabecalho, valores_linha))
                lista_final.append(linha_dict)
            return lista_final
        except Exception as e:
            logger.info(f"❌ Erro ao ler planilha para automação: {e}")
            return []
        
    def atualizar_aba_com_json(self, aba_de_controle, dados_content) -> OperationResult:
        """Sobrescreve uma aba com dados da API."""
        try:
            if not dados_content: return OperationResult.fail("Nenhum dado recebido para atualizar.")
            
            df = self.achatar_json(dados_content).fillna('')
            dados_com_cabecalho = [df.columns.values.tolist()] + df.values.tolist()
            
            aba = self.obter_aba(aba=aba_de_controle)
            if not aba:
                aba = self.planilha.add_worksheet(title=aba_de_controle, rows="100", cols="20")
            
            aba.clear()
            aba.update('A1', dados_com_cabecalho, value_input_option='RAW')
            
            if aba_de_controle in self._cache_abas: del self._cache_abas[aba_de_controle]
            return OperationResult.ok(f"✅ Aba '{aba_de_controle}' atualizada com sucesso.")
        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def importar_excel_para_aba(self, diretorio_excel) -> OperationResult:
        """Importa o Excel para a planilha a partir da primeira linha onde as colunas chave estão vazias."""
        
        try:
            if not os.path.exists(diretorio_excel):
                return OperationResult.fail(f"📂 Pasta não encontrada: {diretorio_excel}")

            arquivos = sorted([f for f in os.listdir(diretorio_excel) if f.endswith(('.xlsx', '.xls'))], reverse=True)
            if not arquivos: 
                return OperationResult.fail("🔍 Nenhum arquivo Excel encontrado na pasta.")

            caminho_excel = os.path.join(diretorio_excel, arquivos[0])
            pular = 4 if caminho_excel.endswith('.xls') else 0
            df_novo = pd.read_excel(caminho_excel, skiprows=pular)
            df_novo = df_novo.loc[:, ~df_novo.columns.astype(str).str.contains('^Unnamed')].dropna(axis=1, how='all').fillna('')

            for col in df_novo.select_dtypes(include=['datetime']).columns:
                df_novo[col] = df_novo[col].dt.strftime('%d/%m/%Y')

            aba = self.obter_aba()
            if not aba:
                return OperationResult.fail(f"❌ Aba '{NOME_ABA}' não existe.")

            dados_sheets = aba.get_all_values()
            
            if not dados_sheets:
                proxima_linha = 1
                corpo_dados = [df_novo.columns.values.tolist()] + df_novo.values.tolist()
            else:
                df_atual = pd.DataFrame(dados_sheets[1:], columns=dados_sheets[0])
                colunas_ref = ['Código Funcionário', 'Nome Funcionário', 'Código Ficha Clínica']
                
                colunas_existentes = [c for c in colunas_ref if c in df_atual.columns]
                
                if not colunas_existentes:
                    proxima_linha = len(dados_sheets) + 1
                else:
                    preenchidos = df_atual[colunas_existentes].apply(lambda x: x.str.strip().ne('')).any(axis=1)
                    
                    if not preenchidos.any():
                        proxima_linha = 2 
                    else:
                        ultima_preenchida_idx = preenchidos[preenchidos].index[-1]
                        proxima_linha = ultima_preenchida_idx + 3 # +1 (index 0), +1 (header), +1 (próxima)

                corpo_dados = df_novo.values.tolist()

            if corpo_dados:
                range_inicio = f"A{proxima_linha}"
                aba.update(range_inicio, corpo_dados, value_input_option='RAW')

            if NOME_ABA in self._cache_abas: del self._cache_abas[NOME_ABA]
            return OperationResult.ok(f"✅ Excel importado na linha {proxima_linha}!")

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def obter_mapa_validacao(self, aba, coluna_chave, coluna_valor):
        """Gera um dicionário de DE-PARA de uma aba."""
        try:
            aba = self.obter_aba(aba=aba)
            dados = aba.get_all_records()
            return {str(item.get(coluna_chave, "")).strip().split('.')[0]: 
                    "".join(c for c in str(item.get(coluna_valor, "")) if c.isprintable()) 
                    for item in dados if item.get(coluna_chave)}
        except Exception as e:
            logger.info(f"❌ Erro ao gerar mapa: {e}")
            return {}

    def buscar_id(self, aba, codigo_busca, col_codigo, nome_busca=None, col_nome=None):
        """
        Busca o ID (Coluna A) na aba informada. 
        Usa cache para não baixar a aba a cada consulta.
        """
        if aba not in self._cache_abas:
            aba = self.obter_aba(aba=aba)
            self._cache_abas[NOME_ABA] = aba.get_all_values() if aba else []

        dados = self._cache_abas[NOME_ABA]
        if not dados: return None

        cabecalho = dados[0]
        try:
            idx_cod = cabecalho.index(col_codigo)
            idx_nome = cabecalho.index(col_nome) if col_nome else None
        except: return None

        alvo_cod = str(codigo_busca).replace(".", "").strip().lower()
        alvo_nome = self._limpar_texto(nome_busca) if nome_busca else None

        for linha in dados[1:]:
            if str(linha[idx_cod]).replace(".", "").strip().lower() == alvo_cod:
                if alvo_nome and idx_nome is not None:
                    if self._limpar_texto(linha[idx_nome]) == alvo_nome:
                        return linha[0]
                else:
                    return linha[0]
        return None
      

    def marcar_status_na_planilha(self, id_busca, col_referencia="Código Ficha Clínica", erro=False) -> OperationResult:
        """Marca status, responsável, IP e horário na linha correta do Google Sheets."""
        NOME_ABA = get_config("google_sheets", "aba")
        
        try:
            aba_instancia = self.obter_aba()
            if not aba_instancia:
                return OperationResult.fail(f"⚠️ Aba '{NOME_ABA}' não encontrada.")
            dados_brutos = aba_instancia.get_all_values()
            if not dados_brutos:
                return OperationResult.fail("A planilha está vazia.")
                
            cabecalhos = dados_brutos[0]
            df = pd.DataFrame(dados_brutos[1:], columns=cabecalhos)

            id_busca_str = str(id_busca).strip()
            filtro = df[col_referencia].astype(str).str.strip() == id_busca_str
            indices = df.index[filtro].tolist()

            if not indices:
                return OperationResult.fail(f"ID '{id_busca_str}' não localizado.")

            linha_sheets = int(indices[0]) + 2 # +1 do header, +1 porque Sheets começa em 1

            info_usuario = obter_identificacao_usuario()
            
            status_texto = "❌ ERRO NO ENVIO" if erro else "✅ ENVIADO"
            
            atualizacoes = {
                "Status": f"{status_texto}",
                "Nome do responsável pelo envio": info_usuario['usuario'].upper(),
                "Ip do responsavel pelo envio": info_usuario['ip'],
                "Horário do envio": info_usuario['horario']
            }

            for nome_coluna, valor in atualizacoes.items():
                if nome_coluna in cabecalhos:
                    col_idx = cabecalhos.index(nome_coluna) + 1
                    aba_instancia.update_cell(linha_sheets, col_idx, valor)
                else:
                    logger.info(f"⚠️ Coluna '{nome_coluna}' não encontrada na aba.")

            if hasattr(self, '_cache_abas') and NOME_ABA in self._cache_abas:
                del self._cache_abas[NOME_ABA]
                
            return OperationResult.ok(f"Status e rastreio atualizados para a ficha {id_busca_str}.")

        except Exception as e:
            try:
                return ErrorTranslator.traduzir(e)
            except:
                return OperationResult.fail(f"Erro ao marcar status: {str(e)}")