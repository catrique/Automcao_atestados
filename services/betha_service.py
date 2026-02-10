import base64
from datetime import datetime
import mimetypes
import pandas as pd
import requests
import json
import os
from config_global import sheets
from services.utils_service import ErrorTranslator, OperationResult
from services.utils_service import logger
from config_network import PROXIES_OFF, sessao_limpa
from config.loaders import get_config


class BethaService:

    def __init__(self):
        # Não precisamos mais de self.path aqui, o loader cuida disso
        self.headers = {}
        self.inicializado = False
        self.erro_inicializacao = ""

        try:
            # Buscamos os valores diretamente usando nossa função dinâmica
            authorization = get_config("betha", "api", "authorization")
            user_access = get_config("betha", "api", "user_access")

            if not authorization or not user_access:
                self.erro_inicializacao = "🔑 Token ou User-Access não encontrados no arquivo de configuração."
            else:
                self.headers = {
                    "Authorization": authorization, 
                    "user-access": user_access,
                    "Content-Type": "application/json"
                }
                self.inicializado = True

        except Exception as e:
            self.inicializado = False
            self.erro_inicializacao = f"❌ Erro ao carregar configurações do Betha: {str(e)}"

    def buscar_matricula(self, termo) -> OperationResult:
        """
        Busca o ID da matrícula na Betha.
        Retorna OperationResult.ok(data=id_matricula) ou OperationResult.fail().
        """
        if not self.inicializado:
            return OperationResult.fail(self.erro_inicializacao)

        endpoint = get_config("betha", "api", "endpoints", "listagem_matricula")
        url = f"{get_config('betha', 'api', 'base_url').rstrip('/')}/{endpoint.lstrip('/')}"
        
        if "/" in termo:
            try:
                numero, contrato = termo.split("/", 1) 
            except ValueError:
                return OperationResult.fail(f"⚠️ Formato de matrícula inválido: {termo}")
                
            termo_sem_barra = termo.replace("/", "")
            filtro = (
                f'(pessoa.nome elike "%25{termo}%25") '
                f'or (codigo.numero = "{numero}" and codigo.contrato = "{contrato}") '
                f'or pessoa.cpf = "{termo_sem_barra}" '
                f'or pessoa.identidade = "{termo_sem_barra}" '
                f'or pessoa.pis = "{termo_sem_barra}"'
            )
        else:
            filtro = (
                f'(pessoa.nome elike "%25{termo}%25") '
                f'or (codigo.numero = "{termo}") '
                f'or pessoa.cpf = "{termo}" '
                f'or pessoa.identidade = "{termo}" '
                f'or pessoa.pis = "{termo}"'
            )

        params = {
            "filter": filtro,
            "filtroSituacao": "TODOS",
            "limit": 20
        }

        try:
            res = sessao_limpa.get(url, headers=self.headers, params=params, timeout=20, proxies=PROXIES_OFF)
            
            res.raise_for_status()
            
            dados = res.json()
            content = dados.get('content', [])

            if not content:
                return OperationResult.fail(f"🔍 Matrícula/Funcionário '{termo}' não localizado na Betha.")

            id_matricula = content[0].get('id')
            return OperationResult.ok(f"✅ Matrícula {termo} encontrada.", data=id_matricula)

        except requests.exceptions.HTTPError as e:
            if e.response.status_code in [401, 403]:
                return OperationResult.fail("🔑 Token da Betha expirado ou inválido. Por favor, atualize o token.")
            return ErrorTranslator.traduzir(e)
            
        except Exception as e:
            return ErrorTranslator.traduzir(e)
        

    def processar_lote_planilha(self, dados_planilha) -> OperationResult:
        """
        Filtra as linhas prontas e retorna um OperationResult com a lista filtrada.
        """
        if not dados_planilha:
            return OperationResult.fail("📊 A planilha parece estar vazia ou não foi lida corretamente.")

        linhas_filtradas = [
            linha for linha in dados_planilha 
            if str(linha.get('Pronto para importação', '')).strip().lower() == 'sim'
        ]

        if not linhas_filtradas:
            return OperationResult.fail(
                "🔍 Nenhuma linha encontrada com 'Sim' na coluna 'Pronto para importação'.\n"
                "Verifique se você marcou os itens corretamente na planilha."
            )

        logger.info(f"✅ Linhas prontas para o service: {len(linhas_filtradas)}.")
        return OperationResult.ok(f"{len(linhas_filtradas)} registros prontos.", data=linhas_filtradas)

    def formatar_data(self, data_planilha) -> str:
        """
        Garante formato YYYY-MM-DD para a API Betha.
        Se falhar, retorna string vazia para o validador de payload capturar.
        """
        if not data_planilha or pd.isna(data_planilha): 
            return ""
            
        data_str = str(data_planilha).strip()
        try:
            if "/" in data_str:
                return datetime.strptime(data_str, "%d/%m/%Y").strftime("%Y-%m-%d")
            
            return data_str[:10] 
        except Exception:
            logger.info(f"⚠️ Erro ao formatar data: {data_str}")
            return ""
        

    def gerar_payloads_lote(self, linhas_filtradas) -> OperationResult:
        """
        Gera a lista de payloads para a API Betha, validando matrículas e IDs do Sheets.
        """
        lista_payloads = []
        erros = 0

        if not linhas_filtradas:
            return OperationResult.fail("Nenhuma linha fornecida para geração de payloads.")

        for linha in linhas_filtradas:
            cod_ficha = str(linha.get('Código Ficha Clínica', '')).strip()
            try:
                matricula_input = str(linha.get('Matrícula Funcionário', '')).strip()
                res_matricula = self.buscar_matricula(matricula_input)
                
                if not res_matricula or not res_matricula.success:
                    logger.info(f"⚠️ Pulando {matricula_input}: {res_matricula.message if res_matricula else 'Erro na busca de matrícula'}")
                    erros += 1
                    continue
                
                id_betha = res_matricula.data
                logger.info(f"✅ Matrícula encontrada para: {matricula_input}: {id_betha}")

                def buscar_id_seguro(aba, codigo, col, nome=None, col_n=None):
                    resultado = sheets.buscar_id(nome_aba=aba, codigo_busca=codigo, col_codigo=col, nome_busca=nome, col_nome=col_n)
                    val = resultado.data if hasattr(resultado, 'success') and resultado.success else resultado
                    return int(val) if val and str(val).isdigit() else None

                id_profissional = buscar_id_seguro("MEDICOS", linha.get("CRM Médico assistente"), "numeroConselho", linha.get("Médico assistente"), "nome")
                id_cid = buscar_id_seguro("CID", linha.get("CID"), "codigo")
                id_tipo_atestado = buscar_id_seguro("TIPOS_ATESTADO", linha.get("Tipo de atestado"), "descricao")
                id_motivo = buscar_id_seguro("MOTIVO_CONSULTA", linha.get("Motivo da Consulta"), "descricao")
                id_tipo_afastamento = buscar_id_seguro("TIPOS_AFASTAMENTO", linha.get("Tipo de afastamento"), "descricao")

                if not all([id_motivo, id_betha]):
                    logger.info(f"⚠️ Dados obrigatórios faltando para Ficha {cod_ficha} (Tipo:{id_tipo_atestado}/Motivo:{id_motivo}/Matrícula:{id_betha})")
                    erros += 1
                    continue

                data_ini = self.formatar_data(linha.get("Data de Afastamento (de)"))
                data_fim = self.formatar_data(linha.get("Nova data final do afastamento"))
                
                caminho_anexos = str(linha.get("Pasta de anexos", "")).strip()
                lista_anexos = self.preparar_anexos(caminho_anexos, data_ini)

                payload = {
                    "matricula": { "id": id_betha },
                    "duracao": int(float(linha.get("Nº de Dias Abonados", 1))),
                    "unidade": "DIAS",
                    "inicioAtestado": data_ini,
                    "fimAtestado": data_fim,
                    "numeroAtestado": cod_ficha,
                    "tipo": { "id": id_tipo_atestado }, 
                    "motivoConsultaMedica": { "id": id_motivo },
                    "tipoAfastamento": { "id": id_tipo_afastamento } if id_tipo_afastamento else None,
                    "profissional": { "id": id_profissional } if id_profissional else None, 
                    "cidPrincipal": { "id": id_cid } if id_cid else None,   
                    "cids": [ { "id": id_cid } ] if id_cid else [],
                    "inseridoPeloRh": True,
                    "localAtendimento": "AMBULATORIO",
                    "geradoApartirDeAfastamento": bool(id_tipo_afastamento),
                    "dataEntrega": f"{data_ini} 08:00:00" if data_ini else None,
                    "anexos": lista_anexos
                }
                
                if payload["inicioAtestado"] and payload["fimAtestado"]:
                    lista_payloads.append(payload)
                else:
                    logger.info(f"⚠️ Ficha {cod_ficha} ignorada por falta de datas.")
                    erros += 1

            except Exception as e:
                logger.info(f"❌ Erro Crítico na linha da ficha {cod_ficha}: {e}")
                erros += 1
                continue
                
        if not lista_payloads:
            return OperationResult.fail(f"Falha ao gerar payloads. {erros} erros encontrados.")

        return OperationResult.ok(
            f"Processamento concluído: {len(lista_payloads)} payloads gerados, {erros} falhas.",
            data=lista_payloads
        )

    def preparar_anexos(self, caminho_pasta, data_afastamento):
        """
        Varre a pasta, faz os uploads e retorna o objeto esperado pela Betha.
        Retorna uma lista vazia se nada for processado.
        """
        if not caminho_pasta:
            return []

        caminho_pasta = os.path.normpath(caminho_pasta.strip())

        if not os.path.exists(caminho_pasta):
            logger.info(f"⚠️ Pasta de anexos não encontrada: {caminho_pasta}")
            return []

        arquivos_servidor = []

        try:
            lista_arquivos = os.listdir(caminho_pasta)
        except Exception as e:
            logger.info(f"❌ Erro ao acessar pasta {caminho_pasta}: {e}")
            return []

        for nome_arquivo in lista_arquivos:
            caminho_completo = os.path.join(caminho_pasta, nome_arquivo)
            
            if os.path.isfile(caminho_completo) and not nome_arquivo.startswith("~$"):
                logger.info(f"⬆️ Enviando anexo: {nome_arquivo}")
                
                try:
                    res_upload = self.upload_arquivo_betha(caminho_completo)
                    
                    dados_retorno = res_upload.data if hasattr(res_upload, 'data') else res_upload
                    
                    if dados_retorno:
                        arquivos_servidor.append(dados_retorno)
                    else:
                        logger.info(f"⚠️ Falha no upload do arquivo: {nome_arquivo}")
                except Exception as e:
                    logger.info(f"❌ Erro inesperado no upload de {nome_arquivo}: {e}")

        if not arquivos_servidor:
            return []

        return [
            {
                "data": data_afastamento,
                "tipoDocumento": {
                    "id": 1781,
                    "descricao": "ATESTADO",
                    "version": 0
                },
                "arquivos": arquivos_servidor
            }
        ]

    def upload_arquivo_betha(self, caminho_arquivo) -> OperationResult:
        """
        Faz o upload do arquivo binário e retorna OperationResult com os dados do servidor.
        """
        if not self.inicializado:
            return OperationResult.fail(self.erro_inicializacao)

         
        base_url = get_config("betha", "api", "base_url").strip()
        endpoint_anexo = get_config("betha", "api", "endpoints", "anexo").lstrip('/')
        
        url = f"{base_url.rstrip('/')}/{endpoint_anexo}"
        headers_upload = self.headers.copy()
        if "Content-Type" in headers_upload:
            del headers_upload["Content-Type"]

        try:
            nome_arquivo = os.path.basename(caminho_arquivo)
            tipo_mime = mimetypes.guess_type(caminho_arquivo)[0] or "application/octet-stream"
            tamanho_mb = os.path.getsize(caminho_arquivo) / (1024 * 1024)
            if tamanho_mb > 10: 
                return OperationResult.fail(f"📁 Arquivo muito grande ({tamanho_mb:.2f}MB). Limite sugerido: 10MB.")

            with open(caminho_arquivo, 'rb') as f:
                files = {
                    'file': (nome_arquivo, f, tipo_mime)
                }
                
                response = requests.post(url, headers=headers_upload, files=files, timeout=60, proxies=PROXIES_OFF)
                
                if response.status_code in [200, 201]:
                    dados = response.json()
                    return OperationResult.ok(f"✅ Upload concluído: {nome_arquivo}", data=dados)
                
                if response.status_code in [401, 403]:
                    return OperationResult.fail("🔑 Token expirado ao tentar fazer upload. Atualize o acesso.")

                return OperationResult.fail(f"❌ Erro Betha ({response.status_code}): {response.text}")

        except FileNotFoundError:
            return OperationResult.fail(f"🔍 Arquivo não encontrado para upload: {caminho_arquivo}")
        except requests.exceptions.Timeout:
            return OperationResult.fail(f"⏳ Tempo esgotado ao enviar '{nome_arquivo}'. Internet muito lenta ou arquivo pesado.")
        except Exception as e:
            return ErrorTranslator.traduzir(e)


    def enviar_atestado(self, payload) -> OperationResult:
        """
        Envia o payload do atestado para a Betha e retorna um OperationResult.
        """
        if not self.inicializado:
            return OperationResult.fail(self.erro_inicializacao)

        endpoint = get_config("betha", "api", "endpoints", "atestado")
        url = f"{get_config('betha', 'api', 'base_url').rstrip('/')}/{endpoint.lstrip('/')}"
        
        cod_ficha = payload.get("numeroAtestado", "S/N")

        try:
            response = requests.post(
                url, 
                headers=self.headers, 
                json=payload, 
                timeout=40,
                proxies=PROXIES_OFF
            )
            
            if response.status_code in [200, 201, 202, 204]:
                return OperationResult.ok(f"Enviado com sucesso!", data=response.json() if response.text else {})

            if response.status_code == 400:
                detalhes = response.json().get('message', response.text)
                return OperationResult.fail(f"❌ Erro de Validação Betha: {detalhes}")

            if response.status_code in [401, 403]:
                return OperationResult.fail("🔑 Token Betha expirado ou sem permissão para esta matrícula.")

            response.raise_for_status()

        except requests.exceptions.HTTPError as e:
            return OperationResult.fail(f"❌ Erro no servidor Betha (HTTP {e.response.status_code})")
        except requests.exceptions.ConnectionError:
            return OperationResult.fail("🌐 Falha de conexão. Verifique sua internet.")
        except Exception as e:
            return ErrorTranslator.traduzir(e)
        
        return OperationResult.fail("Erro desconhecido ao enviar atestado.")