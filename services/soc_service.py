import re
import os
import sys
import time
import unicodedata
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import zipfile
from datetime import datetime, timedelta
import pandas as pd
from config_global import sheets
from selenium.common.exceptions import TimeoutException
from services.auth_service import configurar_e_autenticar_proxy, descriptografar
from services.utils_service import logger
from services.utils_service import ErrorTranslator, OperationResult
from config.loaders import get_config

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class SessionExpired(Exception):
    pass


class SOCService:
    def __init__(self, url_soc):
        self.url_soc = url_soc
        self.driver = None
        self.wait = None

    def _inicializar_driver(self, output_dir) -> OperationResult:
        """Inicializa o driver com tratamento de erro para o executável."""
        try:
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            # options.add_argument("--headless") # Descomente para rodar escondido
            
            prefs = {
                "download.default_directory": os.path.abspath(output_dir),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.automatic_downloads": 1
            }
            options.add_experimental_option("prefs", prefs)
            
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 40) # Timeout de 40s é seguro
            return OperationResult.ok("Driver inicializado.")
        except Exception as e:
            return OperationResult.fail(f"❌ Falha ao iniciar Chrome: {str(e)}")

    def login(self, usuario, senha_texto, senha_virtual_clicks) -> OperationResult:
        """Realiza login com captura de erros de credenciais ou sistema."""
        
        try:
            logger.info(f"🔐 Acessando SOC: {self.url_soc}")
            self.driver.get(self.url_soc)
            configurar_e_autenticar_proxy()
            
            try:
                self.wait.until(EC.presence_of_element_located((By.ID, "bt_entrar")))
            except:
                return OperationResult.fail("⏳ O site do SOC demorou muito para responder.")

            self.driver.find_element(By.ID, "usu").send_keys(usuario)
            self.driver.find_element(By.ID, "senha").send_keys(senha_texto)
            self.driver.find_element(By.ID, "empsoc").click()
            
            self.wait.until(EC.visibility_of_element_located((By.ID, "teclado")))
            for val in senha_virtual_clicks:
                botao = self.driver.find_element(By.XPATH, f"//div[@id='teclado']//input[@value='{val}']")
                botao.click()
                time.sleep(0.3)

            self.driver.find_element(By.ID, "bt_entrar").click()
            self.wait.until(EC.url_changes(self.url_soc))
            
            time.sleep(5)
            if self.wait.until(EC.url_changes(self.url_soc)) is False:
                 return OperationResult.fail("❌ Falha no Login: Usuário, Senha ou Teclado Virtual incorretos.")

            logger.info("✅ Login realizado com sucesso.")
            return OperationResult.ok("Login realizado.")
            
        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def navegar_para_tela(self, cod_tela) -> OperationResult:
        """Navega entre frames com segurança."""
        try:
            time.sleep(2)
            self.driver.switch_to.default_content()
            time.sleep(1)
                
            search_program = self.wait.until(EC.element_to_be_clickable((By.ID, "cod_programa")))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", search_program)
            search_program.click()
            search_program.send_keys(Keys.CONTROL + "a")
            search_program.send_keys(Keys.DELETE)
            search_program.send_keys(cod_tela)
            search_program.send_keys(Keys.ENTER)
            time.sleep(1)
            self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "socframe")))
            return OperationResult.ok(f"Tela {cod_tela} acessada.")
        except Exception as e:
            return OperationResult.fail(f"❌ Não foi possível acessar a tela {cod_tela}.")

    def fechar_sessao(self):
        """Sempre chame isso ao final ou em erro crítico."""
        if self.driver:
            self.driver.quit()
        
    def selecionar_tipo_relatorio(self, valor="11") -> OperationResult:
        try:
            dropdown = self.wait.until(EC.element_to_be_clickable((By.ID, "dat001_codTipoLocalPersonalizacao")))
            select = Select(dropdown)
            select.select_by_value(valor)
            time.sleep(1)
            return OperationResult.ok("Tipo de relatório selecionado.")
        except Exception as e:
            return OperationResult.fail(f"❌ Erro ao selecionar tipo de relatório: {str(e)}")

    def selecionar_checkboxes(self, checkbox_ids=None) -> OperationResult:
        try:
            if checkbox_ids is None:
                checkbox_ids = ["inativos", "dat001_sinaisVitais"]
            
            for checkbox_id in checkbox_ids:
                checkbox = self.wait.until(EC.element_to_be_clickable((By.ID, checkbox_id)))
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                checkbox.click()
                time.sleep(0.5)
            return OperationResult.ok("Checkboxes marcados.")
        except Exception as e:
            return OperationResult.fail(f"❌ Erro ao marcar filtros: {str(e)}")

    def gerar_relatorio_excel(self) -> OperationResult:
        """Solicita a geração do relatório e trata o alerta de sucesso."""
        try:
            botao_excel = self.wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, \"doAcao('excel')\")]")))
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_excel)
            time.sleep(1)
            self.driver.execute_script("arguments[0].click();", botao_excel)
            
            try:
                WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                self.driver.switch_to.alert.accept()
            except:
                pass 
                
            logger.info("📤 Exportação solicitada com sucesso!")
            return OperationResult.ok("Exportação iniciada.")
        except Exception as e:
            return OperationResult.fail(f"❌ Erro ao solicitar Excel: {str(e)}")

    def baixar_ultimo_relatorio(self, tentativas=5) -> OperationResult:
        """
        Tenta buscar e baixar o último relatório com lógica de repetição.
        """
        for i in range(tentativas):
            try:
                logger.info(f"🔍 Buscando relatório (Tentativa {i+1}/{tentativas})...")
                try:
                    botao_procurar = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[name='botao-pesquisar-padrao-soc']")))
                    botao_procurar.click()
                except:
                    self.driver.execute_script("document.getElementsByName('botao-pesquisar-padrao-soc')[0].click();")
                
                time.sleep(3) 

                self.wait.until(EC.presence_of_element_located((By.ID, "tableProcessos")))
                linhas = self.driver.find_elements(By.XPATH, "//table[@id='tableProcessos']//tr[contains(@id, 'linha-pedido-')]")
                
                if not linhas:
                    continue 

                ultima_linha = linhas[-1] 
                
                try:
                    botao_download = ultima_linha.find_element(By.XPATH, ".//a[contains(text(), 'Download')]")
                    self.driver.execute_script("arguments[0].click();", botao_download)
                    return OperationResult.ok("✅ Download iniciado!")
                except:
                    logger.info(f"⏳ Relatório ainda em processamento...")
                    time.sleep(10) 
                    
            except Exception as e:
                logger.info(f"⚠️ Erro na tentativa {i+1}: {e}")
        
        return OperationResult.fail("❌ O relatório não ficou pronto para download a tempo.")
            
    def descompactar_e_renomear_relatorio(self, diretorio) -> OperationResult:
        """
        Aguarda o download, descompacta, organiza em pastas por data e renomeia o XLS.
        """
        tentativas = 45 

        while tentativas > 0:
            arquivos = [f for f in os.listdir(diretorio) if f.endswith('.zip') and not f.endswith('.crdownload')]

            if arquivos:
                caminho_zip = os.path.join(diretorio, arquivos[0])
                time.sleep(1) 
                try:
                    with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
                        nomes_arquivos = zip_ref.namelist()
                        zip_ref.extractall(diretorio)

                    arquivo_extraido = next((f for f in nomes_arquivos if f.endswith('.xls')), None)

                    if not arquivo_extraido:
                        return OperationResult.fail("❌ O ZIP do SOC foi baixado, mas não continha um arquivo .xls")
                    else:
                        logger.info(f"📦 Arquivo extraído: {arquivo_extraido}")

                    caminho_antigo = os.path.join(diretorio, arquivo_extraido)

                    try:
                        df_temp = pd.read_excel(caminho_antigo, skiprows=4, nrows=1)
                        data_valor = df_temp['Data Ficha Clínica'].iloc[0]
                        
                        if isinstance(data_valor, datetime):
                            data_str = data_valor.strftime("%d-%m-%Y")
                        else:
                            data_str = str(data_valor).replace('/', '-').strip()
                    except Exception as e:
                        logger.info(f"⚠️ Não foi possível ler a data do cabeçalho, usando data atual: {e}")
                        data_str = datetime.now().strftime("%d-%m-%Y")

                    pasta_data = os.path.join(diretorio, data_str)
                    if not os.path.exists(pasta_data):
                        os.makedirs(pasta_data)
                        logger.info(f"📁 Pasta criada: {pasta_data}")

                    novo_nome = f"Relatorio_licensas_medicas_{data_str}.xls"
                    caminho_novo = os.path.join(pasta_data, novo_nome)

                    if os.path.exists(caminho_novo):
                       try:
                            os.remove(caminho_novo)
                       except PermissionError:
                            return OperationResult.fail(f"❌ O arquivo '{novo_nome}' está aberto. Feche o Excel e tente novamente.")

                    try:
                        import shutil
                        shutil.move(caminho_antigo, caminho_novo)
                    except Exception as e:
                        return OperationResult.fail(f"❌ Erro ao mover arquivo: {str(e)}")
                    
                    if os.path.exists(caminho_zip):
                        try: os.remove(caminho_zip)
                        except: pass 

                    return OperationResult.ok(f"✅ Relatório processado: {novo_nome}", data=caminho_novo)

                except zipfile.BadZipFile:
                    return OperationResult.fail("❌ O arquivo baixado do SOC está corrompido (ZIP inválido).")
                except Exception as e:
                    return ErrorTranslator.traduzir(e)

            time.sleep(1)
            tentativas -= 1

        return OperationResult.fail("⏳ Tempo esgotado: O download do SOC não foi detectado na pasta.")
       
    def buscar_funcionario_por_codigo(self, codigo) -> OperationResult:
        """Busca um funcionário garantindo o formato de 10 dígitos."""
        try:
            try:
                codigo_limpo = str(int(float(codigo))).zfill(10)
            except (ValueError, TypeError):
                return OperationResult.fail(f"❌ Código de funcionário inválido: {codigo}")

            logger.info(f"🔍 Buscando funcionário: {codigo_limpo}")

            radio_codigo = self.wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[name='codigoPesquisaFuncionario'][value='1']")
            ))
            radio_codigo.click()

            campo_busca = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='nomeSeach']")))
            campo_busca.clear()
            campo_busca.send_keys(codigo_limpo)
            campo_busca.send_keys(Keys.ENTER)
            
            time.sleep(2) 

            xpath_link = f"//td[@class='codigo']//a[normalize-space(text())='{codigo_limpo}']"
            try:
                link_final = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath_link)))
                link_final.click()
                return OperationResult.ok(f"Funcionário {codigo_limpo} selecionado.")
            except TimeoutException:
                return OperationResult.fail(f"⚠️ Funcionário {codigo_limpo} não encontrado na listagem.")

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def obter_dados_ficha(self) -> OperationResult:
        """Captura todos os dados da ficha de uma vez (Sequencial, Médico, CID)."""
        try:
            dados = {
                "sequencial": None,
                "medico_nome": "Não encontrado",
                "medico_crm": "Não encontrado",
                "cid": "Não encontrado"
            }

            try:
                xpath_seq = "//label[contains(text(), 'Código Sequencial')]/following-sibling::span"
                dados["sequencial"] = self.driver.find_element(By.XPATH, xpath_seq).text.strip()
            except:
                logger.info("⚠️ Código Sequencial não localizado.")

            try:
                dados["medico_nome"] = self.driver.find_element(By.CSS_SELECTOR, "span[data-alterado-grava-tela='inputMedico_nomeSolicitante']").text.strip()
                dados["medico_crm"] = self.driver.find_element(By.CSS_SELECTOR, "span[data-alterado-grava-tela='atestadoVo.conselhoClasseSolicitante']").text.strip()
            except:
                logger.info("⚠️ Dados do médico não localizados.")

            cid_localizado = False
            for seletor in ["attestadoVo.cidEsocial", "cidDados"]:
                try:
                    elemento = self.driver.find_element(By.CSS_SELECTOR, f"span[data-alterado-grava-tela='{seletor}']")
                    texto = elemento.text.strip()
                    dados["cid"] = texto.split(" - ")[0] if " - " in texto else texto
                    cid_localizado = True
                    break
                except:
                    continue
            
            return OperationResult.ok("Dados da ficha capturados.", data=dados)

        except Exception as e:
            return ErrorTranslator.traduzir(e)
        
    def download_anexos_atestado(self, nome_funcionario, ficha_clinica, data_ficha, diretorio_base) -> OperationResult:
        """
        Baixa os anexos de um atestado gerenciando janelas e downloads dinâmicos.
        """
        frame_id = "socframe"

        try:
            data_string = str(data_ficha)[:10]
            data_limpa = re.sub(r'\D', '-', data_string)
            nome_pasta = f"{self.normalizar_nome(nome_funcionario)}_{ficha_clinica}"
            caminho_final_anexos = os.path.abspath(os.path.join(diretorio_base, data_limpa, nome_pasta))

            if not os.path.exists(caminho_final_anexos):
                os.makedirs(caminho_final_anexos)
                logger.info(f"📁 Pasta criada: {caminho_final_anexos}")

            self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": caminho_final_anexos
            })

            botao_pasta = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[img[contains(@src, 'pasta')]] | //*[@id='botoes']//td[6]/a")
            ))
            self.driver.execute_script("arguments[0].click();", botao_pasta)
            
            self.wait.until(EC.visibility_of_element_located((By.ID, "arquivosGed")))
            time.sleep(1.5)

            icone_visualizar_xpath = "//table[@id='arquivosGed']//span[@class='icone-visualizar-arquivo icones']"
            total_anexos = len(self.driver.find_elements(By.XPATH, icone_visualizar_xpath))

            if total_anexos == 0:
                logger.info(f"ℹ️ Nenhum anexo encontrado para ficha {ficha_clinica}")
                return OperationResult.ok("Sem anexos", data=caminho_final_anexos)
            
            janela_principal = self.driver.current_window_handle

            for indice in range(total_anexos):
                try:
                    xpath_especifico = f"({icone_visualizar_xpath})[{indice + 1}]"
                    icone = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_especifico)))
                    nome_arquivo = self.driver.execute_script("return arguments[0].parentNode.innerText;", icone).strip().lower()
                    
                    self.driver.execute_script("arguments[0].click();", icone)
                    time.sleep(3)

                    if len(self.driver.window_handles) > 1:
                        self.driver.switch_to.window(self.driver.window_handles[-1])
                        
                        if any(ext in nome_arquivo for ext in ['.jpg', '.jpeg', '.png']):
                            self.driver.execute_script("""
                                var link = document.createElement('a');
                                link.href = window.location.href;
                                link.download = '';
                                document.body.appendChild(link);
                                link.click();
                                document.body.removeChild(link);
                            """)
                        else:
                            time.sleep(2)

                        self.driver.close()
                        self.driver.switch_to.window(janela_principal)
                        self.driver.switch_to.default_content()
                        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, frame_id)))
                    
                    logger.info(f"✅ Anexo {indice + 1}/{total_anexos} baixado em: {nome_pasta}")

                except Exception as e:
                    logger.info(f"⚠️ Erro ao baixar anexo {indice + 1}: {e}")
                    self.driver.switch_to.window(janela_principal)
                    self.driver.switch_to.default_content()
                    try: self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, frame_id)))
                    except: pass

            return OperationResult.ok("Anexos processados", data=caminho_final_anexos)

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def normalizar_nome(self, texto: str) -> str:
        if not texto:
            return ""
        texto = unicodedata.normalize('NFKD', texto)
        texto = texto.encode('ASCII', 'ignore').decode('ASCII')
        texto = re.sub(r'[^A-Za-z0-9 ]+', '', texto)
        texto = re.sub(r'\s+', ' ', texto).strip().upper()

        return texto
            
    def processar_relatorio_licensas(self, caminho_excel, output_dir) -> OperationResult:
        """
        Processa o relatório de licenças médicas, extraindo informações adicionais (CID, Médico)
        diretamente da ficha clínica no SOC.
        """
        try:
            df = pd.read_excel(caminho_excel, skiprows=4)
            df.columns = df.columns.str.strip()
            
            for col in ['Médico assistente', 'CRM Médico assistente', 'CID', 'Pasta de anexos']:
                if col not in df.columns:
                    df[col] = ""

            df['Código Funcionário'] = df['Código Funcionário'].astype(str).str.replace('.0', '', regex=False)
            
            lista_funcionarios = [c for c in df['Código Funcionário'].unique() if str(c).lower() not in ['nan', 'nat', '']]
            total_func = len(lista_funcionarios)

            for i, cod_func in enumerate(lista_funcionarios):
                logger.info(f"\n👥 [{i+1}/{total_func}] Processando Funcionário: {cod_func}")
                
                res_busca = self.buscar_funcionario_por_codigo(cod_func)
                if not res_busca.success:
                    logger.info(f"⚠️ Pulando funcionário {cod_func}: {res_busca.message}")
                    continue
                    
                fichas_do_func = df[df['Código Funcionário'] == cod_func]
                indices_web_clicados = set() 

                for index_excel, row in fichas_do_func.iterrows():
                    try:
                        def formatar_data(v):
                            if pd.isna(v) or str(v).strip().lower() in ['nan', 'nat', '']: return ""
                            try: return pd.to_datetime(v, dayfirst=True).strftime('%d/%m/%Y')
                            except: return str(v).strip()

                        data_f = formatar_data(row['Data Ficha Clínica'])
                        data_i = formatar_data(row['Data de Afastamento (de)'])
                        data_a = formatar_data(row['Data de Afastamento (até)'])
                        nome_func = row['Nome Funcionário']
                        cod_ficha_excel = row['Código Ficha Clínica']
                        
                        logger.info(f"🔎 Buscando na Web: Ficha {data_f} | Início {data_i}")

                        self.wait.until(EC.presence_of_element_located((By.ID, "tabelaFichas")))
                        linhas_web = self.driver.find_elements(By.XPATH, "//table[@id='tabelaFichas']//tr[td]")
                        
                        linha_alvo_index = -1
                        for idx, tr in enumerate(linhas_web):
                            if idx in indices_web_clicados: continue
                            
                            texto_linha = tr.text.replace('\n', ' ').strip()
                            
                            match_ficha = data_f in texto_linha
                            match_inicio = data_i in texto_linha
                            match_tipo = 'Atestado' in texto_linha
                            match_fim = (data_a in texto_linha) if data_a else True

                            if match_ficha and match_inicio and match_fim and match_tipo:
                                linha_alvo_index = idx
                                break
                        
                        if linha_alvo_index != -1:
                            logger.info(f"🎯 Correspondência encontrada na linha web {linha_alvo_index}")
                            
                            
                            link = linhas_web[linha_alvo_index].find_element(By.XPATH, ".//a[contains(@class, 'llinha2')]")
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link)
                            self.driver.execute_script("arguments[0].click();", link)
                            indices_web_clicados.add(linha_alvo_index)

                            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-alterado-grava-tela='inputMedico_nomeSolicitante']")))
                           
                            dados_medico = self.obter_medico_assistente()
                            cid_v = self.obter_cid_principal()
                            anexos_dir = self.download_anexos_atestado(nome_func, cod_ficha_excel, data_f, output_dir)

                            if anexos_dir.success:
                                caminho_para_planilha = anexos_dir.data
                            else:
                                caminho_para_planilha = ""
                                logger.info(f"⚠️ Aviso de anexo: {anexos_dir.message}")

                            df.at[index_excel, 'Médico assistente'] = dados_medico.get('nome', '')
                            df.at[index_excel, 'CRM Médico assistente'] = dados_medico.get('crm', '')
                            df.at[index_excel, 'CID'] = cid_v
                            df.at[index_excel, 'Pasta de anexos'] = caminho_para_planilha
                            
                            logger.info(f"✅ Sucesso: {caminho_para_planilha}")

                            self.driver.back()
                            self.driver.switch_to.default_content()
                            self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "socframe")))
                            self.wait.until(EC.presence_of_element_located((By.ID, "tabelaFichas")))
                        else:
                            logger.info(f"❌ Ficha não encontrada na tabela web.")
                        self.navegar_para_tela('1084')

                    except Exception as e:
                        logger.info(f"⚠️ Erro na ficha {index_excel}: {e}")
                
            novo_caminho = caminho_excel.replace('.xls', '.xlsx')
            df.to_excel(novo_caminho, index=False)
            return OperationResult.ok("Processamento concluído com sucesso!", data=novo_caminho)

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def obter_medico_assistente(self):
        """
        Captura dados do médico assistente
        
        Returns:
            dict: Dicionário com 'nome' e 'crm' do médico
        """
        try:
            nome_elem = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-alterado-grava-tela='inputMedico_nomeSolicitante']")))
            crm_elem = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-alterado-grava-tela='atestadoVo.conselhoClasseSolicitante']")))
            
            return {
                'nome': nome_elem.text.strip(),
                'crm': crm_elem.text.strip()
            }
        except Exception as e:
            logger.info(f"❌ Erro ao capturar dados do médico: {e}")
            return {'nome': 'Erro', 'crm': 'Erro'}

    def obter_cid_principal(self):
        """
        Captura o CID principal do atestado
        
        Returns:
            str: Código do CID
        """
        try:
            logger.info("🔍 Tentando seletor prioritário (atestadoVo.cidEsocial)...")
            elemento = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-alterado-grava-tela='atestadoVo.cidEsocial']"))
            )
        except TimeoutException:
            try:
                logger.info("⚠️ Primeiro seletor não encontrado. Tentando secundário (cidDados)...")
                elemento = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-alterado-grava-tela='cidDados']"))
                )
            except TimeoutException:
                logger.info("❌ Nenhum dos seletores de CID foi encontrado na tela.")
                return "Não encontrado"

        texto_completo = elemento.text.strip()
        codigo_cid = texto_completo.split(" - ")[0] if " - " in texto_completo else texto_completo
        return codigo_cid

    def obter_codigo_sequencial(self):
        """Obtém o código sequencial para validar troca de ficha"""
        try:
            el_seq = self.driver.find_elements(By.XPATH, "//label[contains(text(), 'Código Sequencial')]/following-sibling::span")
            return el_seq[0].text.strip() if el_seq else ""
        except:
            return ""


    def _formatar_data_soc(self, valor):
        """Garante que a data esteja no formato string dd/mm/yyyy para comparação."""
        if pd.isna(valor) or str(valor).strip().lower() in ['nan', 'nat', '']:
            return ""
        try:
            return pd.to_datetime(valor, dayfirst=True).strftime('%d/%m/%Y')
        except:
            return str(valor).strip()

    def _voltar_ao_frame(self):
        """Helper para garantir que o Selenium está sempre no socframe."""
        self.driver.switch_to.default_content()
        self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "socframe")))

    def fechar(self):
        """Fecha o navegador"""
        if self.driver:
            self.driver.quit()
            
    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.fechar()


    def configurar_periodo(self, data_inicio=None, data_fim=None) -> OperationResult:
        """
        Configura o período de datas no relatório e retorna OperationResult.
        """
        try:
            if not data_inicio or not data_fim:
                hoje = datetime.now()
                dia_da_semana = hoje.weekday()

                if dia_da_semana == 0:  
                    inicio_dt = hoje - timedelta(days=3)
                    fim_dt = hoje - timedelta(days=1)
                else:
                    inicio_dt = hoje - timedelta(days=1)
                    fim_dt = hoje - timedelta(days=1)

                data_inicio = inicio_dt.strftime("%d/%m/%Y")
                data_fim = fim_dt.strftime("%d/%m/%Y")

            logger.info(f"📅 Configurando período: {data_inicio} até {data_fim}")

            data_inicial = self.wait.until(EC.presence_of_element_located((By.ID, "dataInicioPeriodo")))
            data_final = self.wait.until(EC.presence_of_element_located((By.ID, "dataFimPeriodo")))

            self.driver.execute_script("arguments[0].value = arguments[1];", data_inicial, data_inicio)
            self.driver.execute_script("arguments[0].value = arguments[1];", data_final, data_fim)
            
            self.driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", data_inicial)
            self.driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", data_final)

            return OperationResult.ok(f"Período configurado: {data_inicio} - {data_fim}", data={"inicio": data_inicio, "fim": data_fim})

        except Exception as e:
            logger.info(f"❌ Erro ao configurar datas: {e}")
            return ErrorTranslator.traduzir(e)


def gerar_relatorio_licensas_medicas(
    url_soc, usuario, senha_texto, senha_virtual_clicks, 
    output_dir, data_inicio=None, data_fim=None, processar_detalhes=True
) -> OperationResult:
    """
    Função principal orquestradora do SOC.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        with SOCService(url_soc) as soc:
            res_driver = soc._inicializar_driver(output_dir)
            if not res_driver.success: return res_driver
            
            res_login = soc.login(usuario, senha_texto, senha_virtual_clicks)
            if not res_login.success: return res_login
            
            soc.navegar_para_tela('237')
            soc.configurar_periodo(data_inicio, data_fim)
            soc.selecionar_tipo_relatorio()
            soc.selecionar_checkboxes()
            soc.gerar_relatorio_excel()
            
            tempo_total = 45
            for i in range(tempo_total):
                # sys.stdout.write(f"\rCronômetro: {i+1} segundos")
                segundos = i +1
                if segundos % 10 == 0 or segundos == 1:
                    logger.info(f"⏳ Aguardando download... ({segundos}s passados)")
                # sys.stdout.flush()
                time.sleep(1)
            logger.info("✅ Tempo de espera finalizado!")
            soc.navegar_para_tela('271')
            
            res_download = soc.baixar_ultimo_relatorio()
            if not res_download.success:
                return OperationResult.fail("❌ O relatório não apareceu na lista de downloads.")
            
            res_caminho = soc.descompactar_e_renomear_relatorio(output_dir)
            if not res_caminho.success:
                return res_caminho
            caminho_xls = res_caminho.data

            if processar_detalhes:
                soc.navegar_para_tela('1084')
                res_final = soc.processar_relatorio_licensas(caminho_xls, output_dir)
                
                if os.path.exists(caminho_xls):
                    os.remove(caminho_xls)
                
                return res_final 
            
            else:
                novo_caminho = caminho_xls.replace('.xls', '.xlsx')
                pd.read_excel(caminho_xls, skiprows=4).to_excel(novo_caminho, index=False)
                os.remove(caminho_xls)
                return OperationResult.ok("Relatório simples gerado", data=novo_caminho)

    except Exception as e:
        logger.info(f" Erro ao contar")
        return ErrorTranslator.traduzir(e)




# def executar_fluxo_soc():
#     """Lógica para gerar relatório SOC e enviar para o Sheets"""
#     logger.info("\n🚀 Iniciando Automação SOC...")
    
#     perfil = "adriana"
#     # perfil = "admin"

#     user_settings = settings['soc']['user'][perfil]
    
#     LOGIN = user_settings['LOGIN']
#     SENHA_TEXTO = user_settings['PASSWORD']
#     URL_SOC = settings['soc']['URL_SOC']
#     clicks_raw = user_settings['SENHA_VIRTUAL']
#     SENHA_VIRTUAL_CLICKS = clicks_raw.split(',') if clicks_raw else []
    
#     OUTPUT_DIR = settings['paths']['downloads']

#     # caminho_relatorio = gerar_relatorio_licensas_medicas(
#     #     url_soc=URL_SOC,
#     #     usuario=LOGIN,
#     #     senha_texto=SENHA_TEXTO,
#     #     senha_virtual_clicks=SENHA_VIRTUAL_CLICKS,
#     #     output_dir=OUTPUT_DIR,
#     #     processar_detalhes=True
#     # )

#     caminho_relatorio = gerar_relatorio_licensas_medicas(
#         url_soc=URL_SOC,
#         usuario=LOGIN,
#         senha_texto=SENHA_TEXTO,
#         senha_virtual_clicks=SENHA_VIRTUAL_CLICKS,
#         output_dir=OUTPUT_DIR,
#         data_inicio = '22/01/2026',
#         data_fim = '22/01/2026',
#         processar_detalhes=True
#     )

#     if caminho_relatorio:
#         logger.info(f"🎉 Relatório gerado com sucesso: {caminho_relatorio}")
        
#         # --- INTEGRAÇÃO COM GOOGLE SHEETS ---
#         try:
#             NOME_ABA = 'licensas_medicas'
            
#             logger.info(f"📤 Enviando arquivo de {OUTPUT_DIR} para a aba {NOME_ABA}...")
            
#             # Usamos o objeto global 'sheets' instanciado no config_global
#             # Note que não passamos mais o NOME_PLANILHA, pois ele já foi definido no config_global
#             sheets.importar_excel_para_aba(OUTPUT_DIR, NOME_ABA)
            
#             logger.info("✅ Integração com Sheets concluída com sucesso!")
#         except Exception as e:
#             logger.info(f"⚠️ Erro ao integrar com Sheets: {e}")
            
#         return caminho_relatorio
#     else:
#         logger.info("❌ Falha ao gerar o relatório SOC.")
#         return None


#   perfil = "adriana" ou perfil = "admin"


def executar_fluxo_soc(perfil_selecionado="admin", data_ini=None, data_fim=None) -> OperationResult:
    """
    Orquestra o download do SOC e a exportação para o Google Sheets.
    """
    logger.info(f"\n🚀 Iniciando Automação SOC - Perfil: {perfil_selecionado}")
    
    try:
        if perfil_selecionado not in get_config('soc', 'user'):
            return OperationResult.fail(f"Perfil '{perfil_selecionado}' não encontrado no config.")


        clicks_raw = descriptografar(get_config('soc','user', perfil_selecionado, 'SENHA_VIRTUAL'))
        senha_virtual = [c.strip() for c in clicks_raw.split(',')] if clicks_raw else []

        resultado_op = gerar_relatorio_licensas_medicas(
            url_soc=get_config('soc', 'URL_SOC'),
            usuario=descriptografar(get_config('soc','user', perfil_selecionado, 'LOGIN')),
            senha_texto=descriptografar(get_config('soc','user', perfil_selecionado, 'PASSWORD')),
            senha_virtual_clicks=(senha_virtual),
            output_dir=get_config('paths', 'downloads'),
            data_inicio=data_ini,
            data_fim=data_fim,
            processar_detalhes=True
        )

        if not resultado_op or not hasattr(resultado_op, 'success') or not resultado_op.success:
            return OperationResult.fail(f"❌ O processo do SOC falhou: {resultado_op.message if hasattr(resultado_op, 'message') else 'Erro desconhecido'}")

        resultado_caminho = resultado_op.data 

        logger.info(f"🎉 Relatório extraído: {os.path.basename(resultado_caminho)}")

        return OperationResult.ok(f"Sucesso! Relatório gerado!", data=resultado_caminho
            )

    except Exception as e:
        logger.info(f"❌ Erro crítico no fluxo: {e}")
        return OperationResult.fail(f"Erro inesperado: {str(e)}")