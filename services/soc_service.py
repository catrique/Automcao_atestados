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
            options.add_argument("--disable-notifications")
            prefs = {
                "download.default_directory": os.path.abspath(output_dir),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.automatic_downloads": 1,
                "profile.default_content_setting_values.notifications": 2,
            }
            options.add_experimental_option("prefs", prefs)
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 40)
            return OperationResult.ok("Driver inicializado.")
        except Exception as e:
            return OperationResult.fail(f"❌ Falha ao iniciar Chrome: {str(e)}")

    def login(self, usuario, senha_texto, senha_virtual_clicks) -> OperationResult:
        """Realiza login e aguarda a confirmação real da entrada no sistema."""
        try:
            logger.info(f"🔐 Acessando SOC: {self.url_soc}")
            self.driver.get(self.url_soc)
            configurar_e_autenticar_proxy()
            try:
                self.wait.until(EC.presence_of_element_located((By.ID, "bt_entrar")))
            except:
                return OperationResult.fail(
                    "⏳ O site do SOC demorou muito para responder."
                )

            self.driver.find_element(By.ID, "usu").send_keys(usuario)
            self.driver.find_element(By.ID, "senha").send_keys(senha_texto)
            self.driver.find_element(By.ID, "empsoc").click()

            self.wait.until(EC.visibility_of_element_located((By.ID, "teclado")))
            for val in senha_virtual_clicks:
                botao = self.driver.find_element(
                    By.XPATH, f"//div[@id='teclado']//input[@value='{val}']"
                )
                botao.click()
                time.sleep(0.3)

            self.driver.find_element(By.ID, "bt_entrar").click()

            timeout_login = 60
            start_time = time.time()

            while time.time() - start_time < timeout_login:
                if self.driver.find_elements(
                    By.ID, "g-recaptcha"
                ) or self.driver.find_elements(By.CLASS_NAME, "captcha-modal"):
                    logger.warning(
                        "⚠️ CAPTCHA detectado! Aguardando resolução manual..."
                    )
                    time.sleep(2)
                    continue

                if self.driver.find_elements(By.CSS_SELECTOR, "a.menu-icon"):
                    logger.info("✅ Login confirmado: Elemento da home detectado.")
                    return OperationResult.ok("Login realizado com sucesso.")

                time.sleep(1)

            return OperationResult.fail(
                "❌ Timeout: O login não foi concluído após 60 segundos."
            )

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def navegar_para_tela(self, cod_tela) -> OperationResult:
        """Navega entre frames com segurança."""
        try:
            self.driver.switch_to.default_content()
            btn_menu = self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "a.menu-icon[data-target='slide-out']")
                )
            )
            self.driver.execute_script("arguments[0].click();", btn_menu)
            self.wait.until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, ".sidenav-overlay[style*='opacity: 1']")
                )
            )
            time.sleep(0.5)

            search_program = self.wait.until(
                EC.element_to_be_clickable((By.ID, "ipt-text-busca-programa-menu"))
            )
            self.driver.execute_script(
                "arguments[0].scrollIntoView(true);", search_program
            )
            search_program.click()
            search_program.send_keys(Keys.CONTROL + "a")
            search_program.send_keys(Keys.DELETE)
            search_program.send_keys(cod_tela)
            search_program.send_keys(Keys.ENTER)
            time.sleep(1)
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "novosocFrame"))
            )
            return OperationResult.ok(f"Tela {cod_tela} acessada.")
        except Exception as e:
            return OperationResult.fail(
                f"❌ Não foi possível acessar a tela {cod_tela}."
            )

    def fechar_sessao(self):
        """Sempre chame isso ao final ou em erro crítico."""
        if self.driver:
            self.driver.quit()

    def selecionar_tipo_relatorio(self, valor="11") -> OperationResult:
        try:
            dropdown = self.wait.until(
                EC.element_to_be_clickable((By.ID, "dat001_codTipoLocalPersonalizacao"))
            )
            select = Select(dropdown)
            select.select_by_value(valor)
            time.sleep(1)
            return OperationResult.ok("Tipo de relatório selecionado.")
        except Exception as e:
            return OperationResult.fail(
                f"❌ Erro ao selecionar tipo de relatório: {str(e)}"
            )

    def selecionar_checkboxes(self, checkbox_ids=None) -> OperationResult:
        try:
            if checkbox_ids is None:
                checkbox_ids = ["inativos", "dat001_sinaisVitais"]

            for checkbox_id in checkbox_ids:
                checkbox = self.wait.until(
                    EC.element_to_be_clickable((By.ID, checkbox_id))
                )
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", checkbox
                )
                checkbox.click()
                time.sleep(0.5)
            return OperationResult.ok("Checkboxes marcados.")
        except Exception as e:
            return OperationResult.fail(f"❌ Erro ao marcar filtros: {str(e)}")

    def gerar_relatorio_excel(self) -> OperationResult:
        """Solicita a geração do relatório e trata o alerta de sucesso."""
        try:
            botao_excel = self.wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(@href, \"doAcao('excel')\")]")
                )
            )
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", botao_excel
            )
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

    def baixar_ultimo_relatorio(self, tentativas=10) -> OperationResult:
        """
        Tenta baixar o relatório, recarregando a página se ainda estiver processando.
        """
        janela_principal = self.driver.current_window_handle

        for i in range(tentativas):
            try:
                logger.info(
                    f"🔍 Verificando relatório (Tentativa {i+1}/{tentativas})..."
                )

                if len(self.driver.window_handles) > 1:
                    for window in self.driver.window_handles:
                        if window != janela_principal:
                            self.driver.switch_to.window(window)
                            self.driver.close()
                    self.driver.switch_to.window(janela_principal)

                try:
                    botao_procurar = self.wait.until(
                        EC.element_to_be_clickable(
                            (By.NAME, "botao-pesquisar-padrao-soc")
                        )
                    )
                    botao_procurar.click()
                    time.sleep(3)
                except Exception:
                    self.driver.execute_script(
                        "document.getElementsByName('botao-pesquisar-padrao-soc')[0].click();"
                    )
                    time.sleep(3)

                self.wait.until(
                    EC.presence_of_element_located((By.ID, "tableProcessos"))
                )
                linhas = self.driver.find_elements(
                    By.XPATH,
                    "//table[@id='tableProcessos']//tr[contains(@id, 'linha-pedido-')]",
                )

                if not linhas:
                    continue

                ultima_linha = linhas[-1]

                carregando = ultima_linha.find_elements(By.CLASS_NAME, "div-carregando")

                if carregando and carregando[0].is_displayed():
                    logger.info(
                        "⏳ Relatório ainda em processamento (ícone de load ativo). Aguardando 5s..."
                    )
                    time.sleep(5)
                    continue
                try:
                    botao_download = ultima_linha.find_element(
                        By.CSS_SELECTOR, "a.download-vault"
                    )
                    self.driver.execute_script("arguments[0].click();", botao_download)
                    time.sleep(2)

                    if len(self.driver.window_handles) > 1:
                        logger.info("📄 Fechando aba extra de download...")
                        for window in self.driver.window_handles:
                            if window != janela_principal:
                                self.driver.switch_to.window(window)
                                self.driver.close()
                        self.driver.switch_to.window(janela_principal)

                    return OperationResult.ok("✅ Download solicitado com sucesso!")

                except Exception:
                    logger.warning(
                        "⚠️ Link de download não encontrado na linha. Tentando novamente..."
                    )
                    time.sleep(5)

            except Exception as e:
                logger.error(f"❌ Erro na tentativa {i+1}: {str(e)}")
                time.sleep(5)

        return OperationResult.fail(
            "❌ O relatório não ficou pronto dentro do tempo limite."
        )

    def descompactar_e_renomear_relatorio(self, diretorio) -> OperationResult:
        """
        Aguarda o download, descompacta, organiza em pastas por data e renomeia o XLS.
        """
        tentativas = 45

        while tentativas > 0:
            arquivos = [
                f
                for f in os.listdir(diretorio)
                if f.endswith(".zip") and not f.endswith(".crdownload")
            ]

            if arquivos:
                caminho_zip = os.path.join(diretorio, arquivos[0])
                time.sleep(1)
                try:
                    with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
                        nomes_arquivos = zip_ref.namelist()
                        zip_ref.extractall(diretorio)

                    arquivo_extraido = next(
                        (f for f in nomes_arquivos if f.endswith(".xls")), None
                    )

                    if not arquivo_extraido:
                        return OperationResult.fail(
                            "❌ O ZIP do SOC foi baixado, mas não continha um arquivo .xls"
                        )
                    else:
                        logger.info(f"📦 Arquivo extraído: {arquivo_extraido}")

                    caminho_antigo = os.path.join(diretorio, arquivo_extraido)

                    try:
                        df_temp = pd.read_excel(caminho_antigo, skiprows=4, nrows=1)
                        data_valor = df_temp["Data Ficha Clínica"].iloc[0]

                        if isinstance(data_valor, datetime):
                            data_str = data_valor.strftime("%d-%m-%Y")
                        else:
                            data_str = str(data_valor).replace("/", "-").strip()
                    except Exception as e:
                        logger.info(
                            f"⚠️ Não foi possível ler a data do cabeçalho, usando data atual: {e}"
                        )
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
                            return OperationResult.fail(
                                f"❌ O arquivo '{novo_nome}' está aberto. Feche o Excel e tente novamente."
                            )

                    try:
                        import shutil

                        shutil.move(caminho_antigo, caminho_novo)
                    except Exception as e:
                        return OperationResult.fail(
                            f"❌ Erro ao mover arquivo: {str(e)}"
                        )

                    if os.path.exists(caminho_zip):
                        try:
                            os.remove(caminho_zip)
                        except:
                            pass

                    return OperationResult.ok(
                        f"✅ Relatório processado: {novo_nome}", data=caminho_novo
                    )

                except zipfile.BadZipFile:
                    return OperationResult.fail(
                        "❌ O arquivo baixado do SOC está corrompido (ZIP inválido)."
                    )
                except Exception as e:
                    return ErrorTranslator.traduzir(e)

            time.sleep(1)
            tentativas -= 1

        return OperationResult.fail(
            "⏳ Tempo esgotado: O download do SOC não foi detectado na pasta."
        )

    def buscar_funcionario_por_codigo(self, codigo) -> OperationResult:
        """Busca um funcionário garantindo o formato de 10 dígitos."""
        try:
            try:
                codigo_limpo = str(int(float(codigo))).zfill(10)
            except (ValueError, TypeError):
                return OperationResult.fail(
                    f"❌ Código de funcionário inválido: {codigo}"
                )

            logger.info(f"🔍 Buscando funcionário: {codigo_limpo}")

            radio_codigo = self.wait.until(
                EC.element_to_be_clickable(
                    (
                        By.CSS_SELECTOR,
                        "input[name='codigoPesquisaFuncionario'][value='1']",
                    )
                )
            )
            radio_codigo.click()

            campo_busca = self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[name='nomeSeach']")
                )
            )
            campo_busca.clear()
            campo_busca.send_keys(codigo_limpo)
            campo_busca.send_keys(Keys.ENTER)

            time.sleep(2)

            xpath_link = (
                f"//td[@class='codigo']//a[normalize-space(text())='{codigo_limpo}']"
            )
            try:
                link_final = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, xpath_link))
                )
                link_final.click()
                return OperationResult.ok(f"Funcionário {codigo_limpo} selecionado.")
            except TimeoutException:
                return OperationResult.fail(
                    f"⚠️ Funcionário {codigo_limpo} não encontrado na listagem."
                )

        except Exception as e:
            return OperationResult.fail(
                f"❌ Erro ao buscar funcionário: {str(e)[:150]}"
            )

    def obter_dados_ficha(self) -> OperationResult:
        """Captura todos os dados da ficha de uma vez (Sequencial, Médico, CID)."""
        try:
            dados = {
                "sequencial": None,
                "medico_nome": "Não encontrado",
                "medico_crm": "Não encontrado",
                "cid": "Não encontrado",
            }

            try:
                xpath_seq = "//label[contains(text(), 'Código Sequencial')]/following-sibling::span"
                dados["sequencial"] = self.driver.find_element(
                    By.XPATH, xpath_seq
                ).text.strip()
            except:
                logger.info("⚠️ Código Sequencial não localizado.")

            try:
                dados["medico_nome"] = self.driver.find_element(
                    By.CSS_SELECTOR,
                    "span[data-alterado-grava-tela='inputMedico_nomeSolicitante']",
                ).text.strip()
                dados["medico_crm"] = self.driver.find_element(
                    By.CSS_SELECTOR,
                    "span[data-alterado-grava-tela='atestadoVo.conselhoClasseSolicitante']",
                ).text.strip()
            except:
                logger.info("⚠️ Dados do médico não localizados.")

            cid_localizado = False
            for seletor in ["attestadoVo.cidEsocial", "cidDados"]:
                try:
                    elemento = self.driver.find_element(
                        By.CSS_SELECTOR, f"span[data-alterado-grava-tela='{seletor}']"
                    )
                    texto = elemento.text.strip()
                    dados["cid"] = texto.split(" - ")[0] if " - " in texto else texto
                    cid_localizado = True
                    break
                except:
                    continue

            return OperationResult.ok("Dados da ficha capturados.", data=dados)

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def download_anexos_atestado(
        self, nome_funcionario, ficha_clinica, data_ficha, diretorio_base
    ) -> OperationResult:
        """
        Baixa os anexos de um atestado gerenciando janelas e downloads dinâmicos.
        """
        frame_id = "novosocFrame"

        try:
            data_string = str(data_ficha)[:10]
            data_limpa = re.sub(r"\D", "-", data_string)
            nome_pasta = f"{self.normalizar_nome(nome_funcionario)}_{ficha_clinica}"
            caminho_final_anexos = os.path.abspath(
                os.path.join(diretorio_base, data_limpa, nome_pasta)
            )

            if not os.path.exists(caminho_final_anexos):
                os.makedirs(caminho_final_anexos)
                logger.info(f"📁 Pasta criada: {caminho_final_anexos}")

            self.driver.execute_cdp_cmd(
                "Page.setDownloadBehavior",
                {"behavior": "allow", "downloadPath": caminho_final_anexos},
            )

            botao_pasta = self.wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//a[img[contains(@src, 'pasta')]] | //*[@id='botoes']//td[6]/a",
                    )
                )
            )
            self.driver.execute_script("arguments[0].click();", botao_pasta)

            self.wait.until(EC.visibility_of_element_located((By.ID, "arquivosGed")))
            time.sleep(1.5)

            icone_visualizar_xpath = "//table[@id='arquivosGed']//span[@class='icone-visualizar-arquivo icones']"
            total_anexos = len(
                self.driver.find_elements(By.XPATH, icone_visualizar_xpath)
            )

            if total_anexos == 0:
                logger.info(f"ℹ️ Nenhum anexo encontrado para ficha {ficha_clinica}")
                return OperationResult.ok("Sem anexos", data=caminho_final_anexos)

            janela_principal = self.driver.current_window_handle

            for indice in range(total_anexos):
                try:
                    xpath_especifico = f"({icone_visualizar_xpath})[{indice + 1}]"
                    icone = self.wait.until(
                        EC.presence_of_element_located((By.XPATH, xpath_especifico))
                    )
                    nome_arquivo = (
                        self.driver.execute_script(
                            "return arguments[0].parentNode.innerText;", icone
                        )
                        .strip()
                        .lower()
                    )

                    self.driver.execute_script("arguments[0].click();", icone)
                    time.sleep(3)

                    if len(self.driver.window_handles) > 1:
                        self.driver.switch_to.window(self.driver.window_handles[-1])

                        if any(
                            ext in nome_arquivo for ext in [".jpg", ".jpeg", ".png"]
                        ):
                            self.driver.execute_script(
                                """
                                var link = document.createElement('a');
                                link.href = window.location.href;
                                link.download = '';
                                document.body.appendChild(link);
                                link.click();
                                document.body.removeChild(link);
                            """
                            )
                        else:
                            time.sleep(2)

                        self.driver.close()
                        self.driver.switch_to.window(janela_principal)
                        self.driver.switch_to.default_content()
                        self.wait.until(
                            EC.frame_to_be_available_and_switch_to_it((By.ID, frame_id))
                        )

                    logger.info(
                        f"✅ Anexo {indice + 1}/{total_anexos} baixado em: {nome_pasta}"
                    )

                except Exception as e:
                    logger.info(f"⚠️ Erro ao baixar anexo {indice + 1}: {e}")
                    self.driver.switch_to.window(janela_principal)
                    self.driver.switch_to.default_content()
                    try:
                        self.wait.until(
                            EC.frame_to_be_available_and_switch_to_it((By.ID, frame_id))
                        )
                    except:
                        pass

            return OperationResult.ok("Anexos processados", data=caminho_final_anexos)

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def normalizar_nome(self, texto: str) -> str:
        if not texto:
            return ""
        texto = unicodedata.normalize("NFKD", texto)
        texto = texto.encode("ASCII", "ignore").decode("ASCII")
        texto = re.sub(r"[^A-Za-z0-9 ]+", "", texto)
        texto = re.sub(r"\s+", " ", texto).strip().upper()

        return texto

    def processar_relatorio_licensas(
        self, caminho_excel, output_dir
    ) -> OperationResult:
        """
        Processa o relatório de licenças médicas, extraindo informações adicionais (CID, Médico)
        diretamente da ficha clínica no SOC.
        """
        try:
            df = pd.read_excel(caminho_excel, skiprows=4)
            df.columns = df.columns.str.strip()

            novo_caminho = caminho_excel.replace(".xls", ".xlsx")

            for col in [
                "Médico assistente",
                "CRM Médico assistente",
                "CID",
                "Pasta de anexos",
            ]:
                if col not in df.columns:
                    df[col] = ""

            df["Código Funcionário"] = (
                df["Código Funcionário"].astype(str).str.replace(".0", "", regex=False)
            )

            lista_funcionarios = [
                c
                for c in df["Código Funcionário"].unique()
                if str(c).lower() not in ["nan", "nat", ""]
            ]
            total_func = len(lista_funcionarios)
            for i, cod_func in enumerate(lista_funcionarios):
                logger.info(
                    f"\n👥 [{i+1}/{total_func}] Processando Funcionário: {cod_func}"
                )

                fichas_do_func = df[df["Código Funcionário"] == cod_func]
                indices_web_clicados = set()

                for index_excel, row in fichas_do_func.iterrows():
                    try:
                        self.navegar_para_tela("1084")
                        res_busca = self.buscar_funcionario_por_codigo(cod_func)
                        if not res_busca.success:
                            logger.info(
                                f"⚠️ Pulando ficha {index_excel} do funcionário {cod_func}: {res_busca.message}"
                            )
                            continue

                        def formatar_data(v):
                            if pd.isna(v) or str(v).strip().lower() in [
                                "nan",
                                "nat",
                                "",
                            ]:
                                return ""
                            try:
                                return pd.to_datetime(v, dayfirst=True).strftime(
                                    "%d/%m/%Y"
                                )
                            except:
                                return str(v).strip()

                        data_f = formatar_data(row["Data Ficha Clínica"])
                        data_i = formatar_data(row["Data de Afastamento (de)"])
                        data_a = formatar_data(row["Data de Afastamento (até)"])
                        nome_func = row["Nome Funcionário"]
                        cod_ficha_excel = row["Código Ficha Clínica"]

                        logger.info(
                            f"🔎 Buscando na Web: Ficha {data_f} | Início {data_i}"
                        )
                        self.wait.until(
                            EC.presence_of_element_located((By.ID, "tabelaFichas"))
                        )
                        linhas_web = self.driver.find_elements(
                            By.XPATH, "//table[@id='tabelaFichas']//tr[td]"
                        )

                        linha_alvo_index = -1
                        for idx, tr in enumerate(linhas_web):
                            if idx in indices_web_clicados:
                                continue

                            texto_linha = tr.text.replace("\n", " ").strip()

                            match_ficha = data_f in texto_linha
                            match_inicio = data_i in texto_linha
                            match_tipo = "Atestado" in texto_linha
                            match_fim = (data_a in texto_linha) if data_a else True

                            if (
                                match_ficha
                                and match_inicio
                                and match_fim
                                and match_tipo
                            ):
                                linha_alvo_index = idx
                                break

                        if linha_alvo_index != -1:
                            logger.info(
                                f"🎯 Correspondência encontrada na linha web {linha_alvo_index}"
                            )

                            link = linhas_web[linha_alvo_index].find_element(
                                By.XPATH, ".//a[contains(@class, 'llinha2')]"
                            )
                            self.driver.execute_script(
                                "arguments[0].scrollIntoView({block: 'center'});", link
                            )
                            self.driver.execute_script("arguments[0].click();", link)
                            indices_web_clicados.add(linha_alvo_index)
                            self.wait.until(
                                EC.presence_of_element_located(
                                    (
                                        By.CSS_SELECTOR,
                                        "span[data-alterado-grava-tela='inputMedico_nomeSolicitante']",
                                    )
                                )
                            )

                            dados_medico = self.obter_medico_assistente()
                            cid_v = self.obter_cid_principal()

                            identificador_atestado = str(cod_ficha_excel)
                            anexos_dir = self.download_anexos_atestado(
                                nome_func, identificador_atestado, data_f, output_dir
                            )
                            if anexos_dir.success:
                                caminho_para_planilha = anexos_dir.data
                            else:
                                caminho_para_planilha = ""
                                logger.info(f"⚠️ Aviso de anexo: {anexos_dir.message}")

                            df.at[index_excel, "Médico assistente"] = dados_medico.get(
                                "nome", ""
                            )
                            df.at[index_excel, "CRM Médico assistente"] = (
                                dados_medico.get("crm", "")
                            )
                            df.at[index_excel, "CID"] = cid_v
                            df.at[index_excel, "Pasta de anexos"] = (
                                caminho_para_planilha
                            )

                            logger.info(f"✅ Sucesso: {caminho_para_planilha}")
                        else:
                            logger.info(f"❌ Ficha não encontrada na tabela web.")

                    except Exception as e:
                        logger.info(f"⚠️ Erro na ficha {index_excel}: {e}")
                        self.navegar_para_tela("1084")

            df.to_excel(novo_caminho, index=False)
            return OperationResult.ok(
                "Processamento concluído com sucesso!", data=novo_caminho
            )

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def obter_medico_assistente(self):
        """
        Captura dados do médico assistente

        Returns:
            dict: Dicionário com 'nome' e 'crm' do médico
        """
        try:
            nome_elem = self.wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "span[data-alterado-grava-tela='inputMedico_nomeSolicitante']",
                    )
                )
            )
            crm_elem = self.wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "span[data-alterado-grava-tela='atestadoVo.conselhoClasseSolicitante']",
                    )
                )
            )

            return {"nome": nome_elem.text.strip(), "crm": crm_elem.text.strip()}
        except Exception as e:
            logger.info(f"❌ Erro ao capturar dados do médico: {e}")
            return {"nome": "Erro", "crm": "Erro"}

    def obter_cid_principal(self):
        """
        Captura o CID principal do atestado

        Returns:
            str: Código do CID
        """
        try:
            logger.info("🔍 Tentando seletor prioritário (atestadoVo.cidEsocial)...")
            elemento = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "span[data-alterado-grava-tela='atestadoVo.cidEsocial']",
                    )
                )
            )
        except TimeoutException:
            try:
                logger.info(
                    "⚠️ Primeiro seletor não encontrado. Tentando secundário (cidDados)..."
                )
                elemento = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "span[data-alterado-grava-tela='cidDados']")
                    )
                )
            except TimeoutException:
                logger.info("❌ Nenhum dos seletores de CID foi encontrado na tela.")
                return "Não encontrado"

        texto_completo = elemento.text.strip()
        codigo_cid = (
            texto_completo.split(" - ")[0]
            if " - " in texto_completo
            else texto_completo
        )
        return codigo_cid

    def obter_codigo_sequencial(self):
        """Obtém o código sequencial para validar troca de ficha"""
        try:
            el_seq = self.driver.find_elements(
                By.XPATH,
                "//label[contains(text(), 'Código Sequencial')]/following-sibling::span",
            )
            return el_seq[0].text.strip() if el_seq else ""
        except:
            return ""

    def _formatar_data_soc(self, valor):
        """Garante que a data esteja no formato string dd/mm/yyyy para comparação."""
        if pd.isna(valor) or str(valor).strip().lower() in ["nan", "nat", ""]:
            return ""
        try:
            return pd.to_datetime(valor, dayfirst=True).strftime("%d/%m/%Y")
        except:
            return str(valor).strip()

    def _clicar_botao_consultar(self) -> OperationResult:
        """
        Clica no botão de lupa (Consultar) para voltar à tela de busca de funcionário.
        Funciona tanto na tela de ficha quanto na listagem de fichas.
        """
        try:
            btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(@href, \"doAcao('browse')\")]")
                )
            )
            self.driver.execute_script("arguments[0].click();", btn)
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[name='nomeSeach']")
                )
            )
            logger.info("🔙 Voltou para tela de busca via botão Consultar.")
            return OperationResult.ok("Voltou para busca.")
        except Exception as e:
            return OperationResult.fail(f"⚠️ Botão Consultar não encontrado: {e}")

    def _voltar_para_busca_e_reabrir_vazio(self) -> bool:
        """
        Clica no botão Consultar (lupa) para voltar à tela de busca de funcionário,
        deixando o campo em branco. A busca do próximo funcionário é feita separadamente.
        Fallback: navega para tela 1084 via menu.
        Retorna True se chegou na tela de busca, False se falhou.
        """
        try:
            btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(@href, \"doAcao('browse')\")]")
                )
            )
            self.driver.execute_script("arguments[0].click();", btn)
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[name='nomeSeach']")
                )
            )
            logger.info("🔙 Tela de busca pronta para próxima ficha.")
            return True
        except Exception:
            logger.info("⚠️ Fallback: navegando para tela 1084...")
            try:
                self.navegar_para_tela("1084")
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "input[name='nomeSeach']")
                    )
                )
                return True
            except Exception as e:
                logger.info(f"❌ Não foi possível voltar para tela de busca: {e}")
                return False

    def _voltar_para_busca_e_reabrir(self, cod_func: str) -> bool:
        """
        Garante que o driver está na tela de busca de funcionário e faz
        a busca completa (digita código + pesquisa + clica no funcionário).
        Tenta: botão Consultar → se falhar, navega para tela 1084.
        Retorna True se chegou na listagem de fichas, False se falhou.
        """
        res = self._clicar_botao_consultar()
        if not res.success:
            logger.info("⚠️ Fallback: navegando para tela 1084...")
            try:
                self.navegar_para_tela("1084")
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "input[name='nomeSeach']")
                    )
                )
            except Exception as e:
                logger.info(f"❌ Não foi possível chegar na tela de busca: {e}")
                return False

        res_busca = self.buscar_funcionario_por_codigo(cod_func)
        if not res_busca.success:
            logger.info(
                f"❌ Falha ao rebuscar funcionário {cod_func}: {res_busca.message}"
            )
            return False

        try:
            self.wait.until(EC.presence_of_element_located((By.ID, "tabelaFichas")))
            return True
        except Exception as e:
            logger.info(f"❌ tabelaFichas não apareceu após rebusca: {e}")
            return False

    def _voltar_ao_frame(self):
        """Helper para garantir que o Selenium está sempre no novosocFrame."""
        self.driver.switch_to.default_content()
        self.wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "novosocFrame"))
        )

    def fechar(self):
        """Fecha o navegador"""
        if self.driver:
            self.driver.quit()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.fechar()

    def descompactar_para_relatorio_geral(
        self, diretorio, data_inicio, data_fim
    ) -> OperationResult:
        """
        Igual a descompactar_e_renomear_relatorio, mas salva na pasta 'relatorios'
        com o nome relatorio_geral_<data_inicio>_<data_fim>.xlsx.
        Não altera o método original.
        """
        import shutil

        tentativas = 45

        while tentativas > 0:
            arquivos = [
                f
                for f in os.listdir(diretorio)
                if f.endswith(".zip") and not f.endswith(".crdownload")
            ]

            if arquivos:
                caminho_zip = os.path.join(diretorio, arquivos[0])
                time.sleep(1)
                try:
                    with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
                        nomes_arquivos = zip_ref.namelist()
                        zip_ref.extractall(diretorio)

                    arquivo_extraido = next(
                        (f for f in nomes_arquivos if f.endswith(".xls")), None
                    )
                    if not arquivo_extraido:
                        return OperationResult.fail(
                            "❌ O ZIP do SOC não continha um arquivo .xls"
                        )

                    caminho_antigo = os.path.join(diretorio, arquivo_extraido)

                    pasta_destino = os.path.join(diretorio, "relatorios")
                    if not os.path.exists(pasta_destino):
                        os.makedirs(pasta_destino)
                        logger.info(f"📁 Pasta criada: {pasta_destino}")

                    ini_fmt = data_inicio.replace("/", "-")
                    fim_fmt = data_fim.replace("/", "-")
                    novo_nome = f"relatorio_geral_{ini_fmt}_{fim_fmt}.xls"
                    caminho_novo = os.path.join(pasta_destino, novo_nome)

                    if os.path.exists(caminho_novo):
                        try:
                            os.remove(caminho_novo)
                        except PermissionError:
                            return OperationResult.fail(
                                f"❌ O arquivo '{novo_nome}' está aberto. Feche o Excel e tente novamente."
                            )

                    shutil.move(caminho_antigo, caminho_novo)

                    if os.path.exists(caminho_zip):
                        try:
                            os.remove(caminho_zip)
                        except:
                            pass

                    logger.info(f"📦 Arquivo salvo: {caminho_novo}")
                    return OperationResult.ok(
                        f"✅ Relatório salvo: {novo_nome}", data=caminho_novo
                    )

                except zipfile.BadZipFile:
                    return OperationResult.fail(
                        "❌ O arquivo baixado do SOC está corrompido (ZIP inválido)."
                    )
                except Exception as e:
                    return ErrorTranslator.traduzir(e)

            time.sleep(1)
            tentativas -= 1

        return OperationResult.fail(
            "⏳ Tempo esgotado: O download do SOC não foi detectado na pasta."
        )

    def processar_relatorio_sem_anexos(
        self, caminho_excel, output_dir, data_inicio, data_fim
    ) -> OperationResult:
        """
        Versão Corrigida: Processa dados do SOC com salvamento seguro e detecção de colunas.
        """
        import shutil

        pasta_destino = os.path.join(output_dir, "relatorios")
        os.makedirs(pasta_destino, exist_ok=True)

        ini_fmt = data_inicio.replace("/", "-")
        fim_fmt = data_fim.replace("/", "-")

        caminho_final = os.path.join(
            pasta_destino, f"relatorio_geral_{ini_fmt}_{fim_fmt}.xlsx"
        )

        def salvar_seguro(df_alvo):
            try:
                temp_file = caminho_final.replace(".xlsx", ".tmp_save.xlsx")

                df_to_save = df_alvo.copy()

                for col in df_to_save.columns:
                    df_to_save[col] = (
                        df_to_save[col].astype(str).replace(["nan", "None", "NAT"], "")
                    )

                df_to_save.to_excel(temp_file, index=False, engine="openpyxl")

                if os.path.exists(temp_file):
                    if os.path.exists(caminho_final):
                        try:
                            os.remove(caminho_final)
                        except:
                            pass
                    shutil.move(temp_file, caminho_final)
                return True
            except Exception as e:
                logger.error(f"❌ Erro ao gravar no disco: {e}")
                return False

        try:
            if os.path.exists(caminho_final):
                logger.info(f"🔄 Retomando processamento existente: {caminho_final}")
                dtype_settings = {
                    "Código Funcionário": str,
                    "CID": str,
                    "CRM Médico assistente": str,
                    "status_processamento": str,
                }
                df = pd.read_excel(caminho_final, dtype=dtype_settings)
            else:
                logger.info(f"📄 Lendo relatório original do SOC: {caminho_excel}")
                df = pd.read_excel(caminho_excel, skiprows=4)
                df.columns = df.columns.str.strip()

                if df.empty or len(df.columns) < 3:
                    logger.warning(
                        "⚠️ Cabeçalho na linha 4 parece incorreto. Tentando linha 0..."
                    )
                    df = pd.read_excel(caminho_excel)
                    df.columns = df.columns.str.strip()

            if "Código Funcionário" in df.columns:
                df["Código Funcionário"] = (
                    df["Código Funcionário"]
                    .astype(str)
                    .str.replace(r"\.0$", "", regex=True)
                    .str.strip()
                )

            if df.empty:
                return OperationResult.fail(
                    "❌ O Excel carregado está vazio. Verifique o arquivo baixado."
                )

            colunas_necessarias = [
                "Médico assistente",
                "CRM Médico assistente",
                "CID",
                "status_processamento",
            ]
            for col in colunas_necessarias:
                if col not in df.columns:
                    df[col] = ""

            if "Código Funcionário" in df.columns:
                df["Código Funcionário"] = (
                    df["Código Funcionário"]
                    .astype(str)
                    .str.replace(r"\.0$", "", regex=True)
                    .str.strip()
                )
            else:
                return OperationResult.fail(
                    "❌ Coluna 'Código Funcionário' não encontrada no Excel."
                )

            df.loc[
                ~df["status_processamento"].isin(["ok", "nao_encontrado"]),
                "status_processamento",
            ] = "pendente"

            if not salvar_seguro(df):
                return OperationResult.fail(
                    "❌ Não foi possível criar o arquivo na rede. Verifique permissões."
                )

            lista_funcionarios = [
                c
                for c in df["Código Funcionário"].unique()
                if str(c).lower() not in ["nan", "nat", ""]
            ]
            total_func = len(lista_funcionarios)

            logger.info(f"📊 Total de funcionários para validar: {total_func}")

            for i, cod_func in enumerate(lista_funcionarios):
                pendentes = df[
                    (df["Código Funcionário"] == cod_func)
                    & (df["status_processamento"] == "pendente")
                ]

                if pendentes.empty:
                    continue

                logger.info(f"\n👥 [{i+1}/{total_func}] Funcionário: {cod_func}")

                self.navegar_para_tela("1084")
                res_busca = self.buscar_funcionario_por_codigo(cod_func)

                if not res_busca.success:
                    logger.warning(f"⚠️ Funcionário {cod_func} não localizado no SOC.")
                    df.loc[
                        df["Código Funcionário"] == cod_func, "status_processamento"
                    ] = "nao_encontrado"
                    salvar_seguro(df)
                    continue

                for index_excel, row in pendentes.iterrows():
                    try:
                        data_f = ""
                        for col_data in [
                            "Data Ficha Clínica",
                            "Data de Emissão",
                            "Data de Sugestão",
                        ]:
                            if col_data in row and pd.notna(row[col_data]):
                                val = row[col_data]
                                data_f = (
                                    val.strftime("%d/%m/%Y")
                                    if hasattr(val, "strftime")
                                    else str(val)
                                )
                                break

                        if not data_f:
                            logger.info(
                                f"⏭️ Linha {index_excel} sem data válida. Pulando."
                            )
                            continue

                        self.wait.until(
                            EC.presence_of_element_located((By.ID, "tabelaFichas"))
                        )
                        linhas_web = self.driver.find_elements(
                            By.XPATH, "//table[@id='tabelaFichas']//tr[td]"
                        )

                        encontrou_na_web = False
                        for tr in linhas_web:
                            texto_tr = tr.text.replace("\n", " ")
                            if data_f in texto_tr and "Atestado" in texto_tr:
                                link = tr.find_element(
                                    By.XPATH, ".//a[contains(@class, 'llinha2')]"
                                )
                                self.driver.execute_script(
                                    "arguments[0].click();", link
                                )
                                encontrou_na_web = True
                                break

                        if encontrou_na_web:
                            self.wait.until(
                                EC.presence_of_element_located(
                                    (
                                        By.CSS_SELECTOR,
                                        "span[data-alterado-grava-tela='inputMedico_nomeSolicitante']",
                                    )
                                )
                            )

                            dados_medico = self.obter_medico_assistente()
                            cid_v = self.obter_cid_principal()

                            df.at[index_excel, "Médico assistente"] = dados_medico.get(
                                "nome", ""
                            )
                            df.at[index_excel, "CRM Médico assistente"] = (
                                dados_medico.get("crm", "")
                            )
                            df.at[index_excel, "CID"] = cid_v
                            df.at[index_excel, "status_processamento"] = "ok"

                            logger.info(f"✅ Ficha {data_f} atualizada.")

                            try:
                                btn_consultar = WebDriverWait(self.driver, 10).until(
                                    EC.element_to_be_clickable(
                                        (By.XPATH, "//img[contains(@src, 'busca.png')]")
                                    )
                                )
                                self.driver.execute_script(
                                    "arguments[0].click();", btn_consultar
                                )
                                self.wait.until(
                                    EC.presence_of_element_located(
                                        (By.ID, "tabelaFichas")
                                    )
                                )
                                logger.info("🔙 Voltou para listagem de fichas.")
                            except Exception as e_voltar:
                                logger.warning(
                                    f"⚠️ Não conseguiu voltar via lupa, tentando doAcao browse: {e_voltar}"
                                )
                                try:
                                    btn2 = WebDriverWait(self.driver, 10).until(
                                        EC.element_to_be_clickable(
                                            (
                                                By.XPATH,
                                                "//a[contains(@href, \"doAcao('browse')\")]",
                                            )
                                        )
                                    )
                                    self.driver.execute_script(
                                        "arguments[0].click();", btn2
                                    )
                                    self.wait.until(
                                        EC.presence_of_element_located(
                                            (By.ID, "tabelaFichas")
                                        )
                                    )
                                    logger.info(
                                        "🔙 Voltou para listagem via doAcao browse."
                                    )
                                except Exception as e_browse:
                                    logger.warning(
                                        f"⚠️ Fallback: rebuscando funcionário {cod_func}: {e_browse}"
                                    )
                                    self._voltar_para_busca_e_reabrir(cod_func)
                        else:
                            df.at[index_excel, "status_processamento"] = (
                                "nao_encontrado"
                            )
                            logger.info(f"❌ Ficha {data_f} não vista na web.")

                        salvar_seguro(df)

                    except Exception as e:
                        logger.error(f"⚠️ Erro na linha {index_excel}: {e}")
                        df.at[index_excel, "status_processamento"] = "erro"
                        salvar_seguro(df)
                        self._voltar_ao_frame()

            return OperationResult.ok("Processamento finalizado!", data=caminho_final)

        except Exception as e:
            return ErrorTranslator.traduzir(e)

    def configurar_periodo(self, data_inicio, data_fim) -> OperationResult:
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

            data_inicial = self.wait.until(
                EC.presence_of_element_located((By.ID, "dataInicioPeriodo"))
            )
            data_final = self.wait.until(
                EC.presence_of_element_located((By.ID, "dataFimPeriodo"))
            )

            self.driver.execute_script(
                "arguments[0].value = arguments[1];", data_inicial, data_inicio
            )
            self.driver.execute_script(
                "arguments[0].value = arguments[1];", data_final, data_fim
            )

            self.driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change'));", data_inicial
            )
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change'));", data_final
            )

            return OperationResult.ok(
                f"Período configurado: {data_inicio} - {data_fim}",
                data={"inicio": data_inicio, "fim": data_fim},
            )

        except Exception as e:
            logger.info(f"❌ Erro ao configurar datas: {e}")
            return ErrorTranslator.traduzir(e)
    def _buscar_por_cpf(self, cpf: str) -> OperationResult:
        """
        Preenche o campo de busca com o CPF, seleciona o radio 'CPF' e executa
        a pesquisa. Aguarda a tabela de resultados aparecer.
        """
        try:
            radio_cpf = self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[name='codigoPesquisaFuncionario'][value='3']")
                )
            )
            radio_cpf.click()
 
            campo_busca = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='nomeSeach']"))
            )
            campo_busca.clear()
            campo_busca.send_keys(cpf)
            campo_busca.send_keys(Keys.ENTER)
 
            time.sleep(2)
 
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.resultados")))
            logger.info(f"🔍 Busca por CPF '{cpf}' executada.")
            return OperationResult.ok("Busca por CPF concluída.")
        except Exception as e:
            return OperationResult.fail(f"❌ Erro ao buscar por CPF '{cpf}': {str(e)[:150]}")
 
    def _localizar_matricula_na_tabela(self, matricula_alvo: str) -> OperationResult:
        """
        Varre as linhas da tabela de resultados e clica na linha cuja coluna
        'Matrícula' corresponda exatamente à matrícula informada.
        """
        try:
            linhas = self.driver.find_elements(
                By.CSS_SELECTOR, "table.resultados tbody tr:not(:first-child)"
            )
 
            if not linhas:
                return OperationResult.fail("⚠️ Nenhum resultado na tabela de busca.")
 
            for linha in linhas:
                celulas = linha.find_elements(By.TAG_NAME, "td")
                if len(celulas) < 6:
                    continue
 
                matricula_web = celulas[5].text.strip()
                if matricula_web == matricula_alvo:
                    link_codigo = celulas[0].find_element(By.TAG_NAME, "a")
                    self.driver.execute_script("arguments[0].click();", link_codigo)
                    logger.info(f"✅ Matrícula '{matricula_alvo}' localizada e selecionada.")
                    return OperationResult.ok("Matrícula localizada.")
 
            return OperationResult.fail(
                f"⚠️ Matrícula '{matricula_alvo}' não encontrada nos resultados."
            )
        except Exception as e:
            return OperationResult.fail(
                f"❌ Erro ao localizar matrícula na tabela: {str(e)[:150]}"
            )
 
    def _clicar_alterar(self) -> OperationResult:
        """Clica no botão 'Alterar' (ícone de edição) no cadastro do funcionário."""
        try:
            btn_alterar = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(@href, \"doAcao('alt')\")]")
                )
            )
            self.driver.execute_script("arguments[0].click();", btn_alterar)
            logger.info("✏️ Modo de edição ativado.")
            return OperationResult.ok("Modo de edição ativado.")
        except Exception as e:
            return OperationResult.fail(f"❌ Erro ao clicar em Alterar: {str(e)[:150]}")
 
    def _preencher_data_demissao(self, data_demissao: str) -> OperationResult:
        """
        Preenche o campo 'dataDemissao' com a data fornecida no formato dd/mm/yyyy.
        Usa JavaScript para garantir que o datepicker não interfira.
        """
        try:
            campo = self.wait.until(
                EC.presence_of_element_located((By.ID, "dataDemissao"))
            )
            self.driver.execute_script("arguments[0].removeAttribute('readonly');", campo)
            self.driver.execute_script(
                "arguments[0].value = arguments[1];", campo, data_demissao
            )
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change'));", campo
            )
            logger.info(f"📅 Data de demissão preenchida: {data_demissao}")
            return OperationResult.ok("Data de demissão preenchida.")
        except Exception as e:
            return OperationResult.fail(
                f"❌ Erro ao preencher data de demissão: {str(e)[:150]}"
            )
 
    def _selecionar_situacao_inativo(self) -> OperationResult:
        """Seleciona a opção 'Inativo' no campo situação do funcionário."""
        try:
            dropdown = self.wait.until(
                EC.element_to_be_clickable((By.ID, "situacao"))
            )
            Select(dropdown).select_by_value("Inativo")
            logger.info("🔴 Situação definida como 'Inativo'.")
            return OperationResult.ok("Situação definida como Inativo.")
        except Exception as e:
            return OperationResult.fail(
                f"❌ Erro ao selecionar situação Inativo: {str(e)[:150]}"
            )
 
    def _gravar_cadastro(self) -> OperationResult:
        """Clica no botão 'Gravar' para salvar as alterações no cadastro."""
        try:
            btn_gravar = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(@href, \"doAcao('save')\")]")
                )
            )
            self.driver.execute_script("arguments[0].click();", btn_gravar)
            time.sleep(2)
            logger.info("💾 Cadastro gravado.")
            return OperationResult.ok("Cadastro gravado com sucesso.")
        except Exception as e:
            return OperationResult.fail(f"❌ Erro ao gravar cadastro: {str(e)[:150]}")
 
    def _voltar_para_busca_funcionario(self) -> OperationResult:
        """
        Clica no botão 'Consultar' (lupa) para retornar à tela de busca de funcionário.
        Fallback: navega para a tela via menu.
        """
        try:
            btn_consultar = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(@href, \"doAcao('browse')\")]")
                )
            )
            self.driver.execute_script("arguments[0].click();", btn_consultar)
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='nomeSeach']"))
            )
            logger.info("🔙 Voltou para a tela de busca de funcionário.")
            return OperationResult.ok("Tela de busca pronta.")
        except Exception:
            logger.warning("⚠️ Fallback: renavegando para tela 232...")
            try:
                self.navegar_para_tela("232")
                self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='nomeSeach']"))
                )
                return OperationResult.ok("Tela de busca restaurada via menu.")
            except Exception as e:
                return OperationResult.fail(
                    f"❌ Não foi possível voltar para a busca: {str(e)[:150]}"
                )
 
    def inativar_servidor_no_soc(
        self, cpf: str, matricula: str, data_demissao: str
    ) -> OperationResult:
        """
        Executa o fluxo completo de inativação de um servidor no SOC:
        busca pelo CPF → localiza a matrícula → edita → preenche demissão →
        define como Inativo → grava → volta para busca.
        """
        res_busca = self._buscar_por_cpf(cpf)
        if not res_busca.success:
            return res_busca
 
        res_matricula = self._localizar_matricula_na_tabela(matricula)
        if not res_matricula.success:
            return res_matricula
 
        time.sleep(1.5)
 
        res_alterar = self._clicar_alterar()
        if not res_alterar.success:
            return res_alterar
 
        res_data = self._preencher_data_demissao(data_demissao)
        if not res_data.success:
            return res_data
 
        res_situacao = self._selecionar_situacao_inativo()
        if not res_situacao.success:
            return res_situacao
 
        self._voltar_para_busca_funcionario()
        return OperationResult.ok(f"✅ Servidor matrícula '{matricula}' inativado com sucesso.")

def gerar_relatorio_licensas_medicas(
    url_soc,
    usuario,
    senha_texto,
    senha_virtual_clicks,
    output_dir,
    data_inicio=None,
    data_fim=None,
    processar_detalhes=True,
) -> OperationResult:
    """
    Função principal orquestradora do SOC.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        with SOCService(url_soc) as soc:
            res_driver = soc._inicializar_driver(output_dir)
            if not res_driver.success:
                return res_driver

            res_login = soc.login(usuario, senha_texto, senha_virtual_clicks)
            if not res_login.success:
                return res_login
            time.sleep(1)
            soc.navegar_para_tela("237")
            soc.configurar_periodo(data_inicio, data_fim)
            soc.selecionar_tipo_relatorio()
            soc.selecionar_checkboxes()
            soc.gerar_relatorio_excel()

            tempo_total = 45
            for i in range(tempo_total):
                segundos = i + 1
                if segundos % 10 == 0 or segundos == 1:
                    logger.info(f"⏳ Aguardando download... ({segundos}s passados)")
                time.sleep(1)
            logger.info("✅ Tempo de espera finalizado!")
            soc.navegar_para_tela("271")

            res_download = soc.baixar_ultimo_relatorio()
            if not res_download.success:
                return OperationResult.fail(
                    "❌ O relatório não apareceu na lista de downloads."
                )

            res_caminho = soc.descompactar_e_renomear_relatorio(output_dir)
            if not res_caminho.success:
                return res_caminho
            caminho_xls = res_caminho.data

            if processar_detalhes:
                soc.navegar_para_tela("1084")
                res_final = soc.processar_relatorio_licensas(caminho_xls, output_dir)

                if os.path.exists(caminho_xls):
                    os.remove(caminho_xls)

                return res_final

            else:
                novo_caminho = caminho_xls.replace(".xls", ".xlsx")
                pd.read_excel(caminho_xls, skiprows=4).to_excel(
                    novo_caminho, index=False
                )
                os.remove(caminho_xls)
                return OperationResult.ok("Relatório simples gerado", data=novo_caminho)

    except Exception as e:
        logger.info(f" Erro ao contar")
        return ErrorTranslator.traduzir(e)


def executar_fluxo_soc(
    data_ini, data_fim, perfil_selecionado="admin"
) -> OperationResult:
    """
    Orquestra o download do SOC e a exportação para o Google Sheets.
    """
    logger.info(f"\n🚀 Iniciando Automação SOC - Perfil: {perfil_selecionado}")

    try:
        if perfil_selecionado not in get_config("soc", "user"):
            return OperationResult.fail(
                f"Perfil '{perfil_selecionado}' não encontrado no config."
            )

        clicks_raw = descriptografar(
            get_config("soc", "user", perfil_selecionado, "SENHA_VIRTUAL")
        )
        senha_virtual = [c.strip() for c in clicks_raw.split(",")] if clicks_raw else []

        resultado_op = gerar_relatorio_licensas_medicas(
            url_soc=get_config("soc", "URL_SOC"),
            usuario=descriptografar(
                get_config("soc", "user", perfil_selecionado, "LOGIN")
            ),
            senha_texto=descriptografar(
                get_config("soc", "user", perfil_selecionado, "PASSWORD")
            ),
            senha_virtual_clicks=(senha_virtual),
            output_dir=get_config("paths", "downloads"),
            data_inicio=data_ini,
            data_fim=data_fim,
            processar_detalhes=True,
        )

        if isinstance(resultado_op, str):
            return OperationResult.fail(
                f"❌ Erro inesperado (retorno string): {resultado_op}"
            )

        if not resultado_op or not resultado_op.success:
            msg = (
                resultado_op.message
                if hasattr(resultado_op, "message")
                else "Erro desconhecido"
            )
            return OperationResult.fail(f"❌ O processo do SOC falhou: {msg}")

        resultado_caminho = resultado_op.data

        logger.info(f"🎉 Relatório extraído: {os.path.basename(resultado_caminho)}")

        return OperationResult.ok(f"Sucesso! Relatório gerado!", data=resultado_caminho)

    except Exception as e:
        logger.info(f"❌ Erro crítico no fluxo: {e}")
        return OperationResult.fail(f"Erro inesperado: {str(e)}")


def listar_relatorios_pendentes() -> list:
    """
    Retorna lista de dicts com relatórios que ainda têm fichas por processar.
    Cada dict: { 'caminho': str, 'nome': str, 'pendentes': int, 'total': int }
    """
    output_dir = get_config("paths", "downloads")
    pasta_relatorios = os.path.join(output_dir, "relatorios")

    if not os.path.exists(pasta_relatorios):
        return []

    resultado = []
    for arquivo in os.listdir(pasta_relatorios):
        if not arquivo.endswith(".xlsx") or not arquivo.startswith("relatorio_geral_"):
            continue
        caminho = os.path.join(pasta_relatorios, arquivo)
        try:
            df = pd.read_excel(caminho)
            if "status_processamento" not in df.columns:
                continue
            total = len(df)
            pendentes = len(
                df[~df["status_processamento"].isin(["ok", "nao_encontrado"])]
            )
            if pendentes > 0:
                resultado.append(
                    {
                        "caminho": caminho,
                        "nome": arquivo,
                        "pendentes": pendentes,
                        "total": total,
                    }
                )
        except Exception:
            continue

    return resultado


def retomar_relatorio_por_periodo(
    caminho_excel: str, perfil_selecionado: str = "admin"
) -> OperationResult:
    """
    Retoma o processamento de um relatório anterior que foi interrompido,
    continuando a partir das fichas ainda com status 'pendente' ou 'erro'.
    Não baixa novo relatório do SOC.
    """
    logger.info(f"\n🔄 Retomando relatório: {os.path.basename(caminho_excel)}")

    try:
        if not os.path.exists(caminho_excel):
            return OperationResult.fail(f"❌ Arquivo não encontrado: {caminho_excel}")

        nome = os.path.basename(caminho_excel).replace(".xlsx", "")
        partes = nome.split("_")
        try:
            data_ini = partes[2].replace("-", "/")
            data_fim = partes[3].replace("-", "/")
        except IndexError:
            return OperationResult.fail(
                "❌ Não foi possível extrair as datas do nome do arquivo."
            )

        if perfil_selecionado not in get_config("soc", "user"):
            return OperationResult.fail(
                f"❌ Perfil '{perfil_selecionado}' não encontrado no config."
            )

        clicks_raw = descriptografar(
            get_config("soc", "user", perfil_selecionado, "SENHA_VIRTUAL")
        )
        senha_virtual = [c.strip() for c in clicks_raw.split(",")] if clicks_raw else []
        output_dir = get_config("paths", "downloads")

        with SOCService(get_config("soc", "URL_SOC")) as soc:
            res_driver = soc._inicializar_driver(output_dir)
            if not res_driver.success:
                return res_driver

            res_login = soc.login(
                usuario=descriptografar(
                    get_config("soc", "user", perfil_selecionado, "LOGIN")
                ),
                senha_texto=descriptografar(
                    get_config("soc", "user", perfil_selecionado, "PASSWORD")
                ),
                senha_virtual_clicks=senha_virtual,
            )
            if not res_login.success:
                return res_login

            time.sleep(1)
            soc.navegar_para_tela("1084")

            res_final = soc.processar_relatorio_sem_anexos(
                caminho_excel, output_dir, data_ini, data_fim
            )

            if not res_final.success:
                return res_final

            logger.info(f"🎉 Retomada concluída: {res_final.data}")
            return OperationResult.ok(
                "Retomada concluída com sucesso!", data=res_final.data
            )

    except Exception as e:
        logger.info(f"❌ Erro crítico em retomar_relatorio_por_periodo: {e}")
        return OperationResult.fail(f"Erro inesperado: {str(e)}")


def exportar_relatorio_por_periodo(
    data_ini: str, data_fim: str, perfil_selecionado: str = "admin"
) -> OperationResult:
    """
    Acessa o SOC, gera o relatório no intervalo informado, lê os dados de cada
    ficha clínica (CID, CRM, médico) e salva o Excel na pasta 'relatorios' com
    o nome relatorio_geral_<data_ini>_<data_fim>.xlsx.
    Não baixa anexos e não importa para o Google Sheets.
    """
    logger.info(f"\n🚀 Exportando relatório SOC por período: {data_ini} → {data_fim}")

    try:
        for label, valor in [("Data Inicial", data_ini), ("Data Final", data_fim)]:
            try:
                datetime.strptime(valor, "%d/%m/%Y")
            except ValueError:
                return OperationResult.fail(
                    f"❌ {label} inválida: '{valor}'. Use o formato dd/mm/aaaa."
                )

        if perfil_selecionado not in get_config("soc", "user"):
            return OperationResult.fail(
                f"❌ Perfil '{perfil_selecionado}' não encontrado no config."
            )

        clicks_raw = descriptografar(
            get_config("soc", "user", perfil_selecionado, "SENHA_VIRTUAL")
        )
        senha_virtual = [c.strip() for c in clicks_raw.split(",")] if clicks_raw else []
        output_dir = get_config("paths", "downloads")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        with SOCService(get_config("soc", "URL_SOC")) as soc:
            res_driver = soc._inicializar_driver(output_dir)
            if not res_driver.success:
                return res_driver

            res_login = soc.login(
                usuario=descriptografar(
                    get_config("soc", "user", perfil_selecionado, "LOGIN")
                ),
                senha_texto=descriptografar(
                    get_config("soc", "user", perfil_selecionado, "PASSWORD")
                ),
                senha_virtual_clicks=senha_virtual,
            )
            if not res_login.success:
                return res_login

            time.sleep(1)
            soc.navegar_para_tela("237")
            soc.configurar_periodo(data_ini, data_fim)
            soc.selecionar_tipo_relatorio()
            soc.selecionar_checkboxes()
            soc.gerar_relatorio_excel()

            tempo_total = 15
            for i in range(tempo_total):
                segundos = i + 1
                if segundos % 5 == 0 or segundos == 1:
                    logger.info(f"⏳ Aguardando download... ({segundos}s passados)")
                time.sleep(1)
            logger.info("✅ Tempo de espera finalizado!")

            soc.navegar_para_tela("271")
            res_download = soc.baixar_ultimo_relatorio()
            if not res_download.success:
                return OperationResult.fail(
                    "❌ O relatório não apareceu na lista de downloads."
                )

            res_xls = soc.descompactar_para_relatorio_geral(
                output_dir, data_ini, data_fim
            )
            if not res_xls.success:
                return res_xls
            caminho_xls = res_xls.data

            soc.navegar_para_tela("1084")
            res_final = soc.processar_relatorio_sem_anexos(
                caminho_xls, output_dir, data_ini, data_fim
            )

            if os.path.exists(caminho_xls):
                os.remove(caminho_xls)

            if not res_final.success:
                return res_final

            logger.info(f"🎉 Relatório gerado: {res_final.data}")
            return OperationResult.ok(
                "Relatório exportado com sucesso!", data=res_final.data
            )

    except Exception as e:
        logger.info(f"❌ Erro crítico em exportar_relatorio_por_periodo: {e}")
        return OperationResult.fail(f"Erro inesperado: {str(e)}")


 
def inativar_servidores(excel_path: str, perfil_selecionado: str = "admin") -> OperationResult:
    """
    Lê o arquivo de demissões, busca CPFs na aba 'Servidores' do Google Sheets,
    acessa o SOC e inativa cada servidor preenchendo a data de demissão e
    alterando a situação para 'Inativo'.
 
    Suporta retomada: registros com 'Inativado' = 'ok' ou 'nao_encontrado'
    são ignorados nas próximas execuções.
    """
    logger.info(f"\n🚀 Iniciando inativação de servidores. Arquivo: {excel_path}")
 
    df, erro = _preparar_dataframe_inativacao(excel_path)
    if erro:
        return OperationResult.fail(erro)
 
    pendentes = df[~df["Inativado"].astype(str).str.strip().isin(["ok", "nao_encontrado"])]
    if pendentes.empty:
        logger.info("✅ Todos os servidores já foram processados.")
        return OperationResult.ok("Nenhum servidor pendente.")
 
    logger.info(f"📊 Servidores pendentes: {len(pendentes)}")
 
    caminho_saida = (
        excel_path if excel_path.endswith(".xlsx")
        else os.path.splitext(excel_path)[0] + ".xlsx"
    )
 
    if perfil_selecionado not in get_config("soc", "user"):
        return OperationResult.fail(f"❌ Perfil '{perfil_selecionado}' não encontrado no config.")
 
    clicks_raw = descriptografar(get_config("soc", "user", perfil_selecionado, "SENHA_VIRTUAL"))
    senha_virtual = [c.strip() for c in clicks_raw.split(",")] if clicks_raw else []
    output_dir = get_config("paths", "downloads")
    os.makedirs(output_dir, exist_ok=True)
 
    try:
        with SOCService(get_config("soc", "URL_SOC")) as soc:
            res_driver = soc._inicializar_driver(output_dir)
            if not res_driver.success:
                return res_driver
 
            res_login = soc.login(
                usuario=descriptografar(get_config("soc", "user", perfil_selecionado, "LOGIN")),
                senha_texto=descriptografar(get_config("soc", "user", perfil_selecionado, "PASSWORD")),
                senha_virtual_clicks=senha_virtual,
            )
            if not res_login.success:
                return res_login
 
            time.sleep(1)
            soc.navegar_para_tela("232")
 
            total = len(pendentes)
            for contador, (idx, linha) in enumerate(pendentes.iterrows(), start=1):
                matricula     = str(linha["Matrícula"]).strip()
                data_demissao = str(linha["Data Demissão"]).strip()
                nome          = str(linha.get("Nome", "")).strip()
                cpf           = str(linha.get("CPF", "")).strip()
 
                logger.info(f"\n👤 [{contador}/{total}] {nome} | Matrícula: {matricula} | CPF: {cpf}")
 
                if not cpf or cpf.lower() in ("nan", "none", ""):
                    logger.warning(f"⚠️ CPF não encontrado para '{nome}' (matrícula {matricula}). Pulando.")
                    df.at[idx, "Inativado"] = "nao_encontrado"
                    df.to_excel(caminho_saida, index=False)
                    continue
 
                if not data_demissao or data_demissao.lower() in ("nan", "nat", "none", ""):
                    logger.warning(f"⚠️ Data de demissão ausente para '{nome}'. Pulando.")
                    df.at[idx, "Inativado"] = "nao_encontrado"
                    df.to_excel(caminho_saida, index=False)
                    continue
 
                res = soc.inativar_servidor_no_soc(cpf, matricula, data_demissao)
 
                df.at[idx, "Inativado"] = "ok" if res.success else "erro"
                if res.success:
                    logger.info(f"✅ {nome} inativado.")
                else:
                    logger.error(f"❌ Falha ao inativar {nome}: {res.message}")
 
                try:
                    df.to_excel(caminho_saida, index=False)
                except Exception as e_save:
                    logger.warning(f"⚠️ Não foi possível salvar progresso: {e_save}")
 
        total_ok    = len(df[df["Inativado"] == "ok"])
        total_falha = len(df[df["Inativado"].isin(["erro", "nao_encontrado"])])
        logger.info(f"\n🎉 Concluído. Sucesso: {total_ok} | Falhas: {total_falha}")
        return OperationResult.ok(
            f"Inativação concluída. Sucesso: {total_ok} | Falhas: {total_falha}",
            data=caminho_saida,
        )
 
    except Exception as e:
        logger.error(f"❌ Erro crítico em inativar_servidores: {e}")
        return OperationResult.fail(f"Erro inesperado: {str(e)}")
    

def _normalizar_matricula(matricula_raw) -> str:
    """
    Remove o sufixo '/0' da matrícula, mantendo qualquer outro sufixo.
        '123456789/0' → '123456789'
        '987654321/2' → '987654321/2'
    """
    matricula = str(matricula_raw).strip()
    if matricula.endswith("/0"):
        return matricula[:-2]
    return matricula
 
 
 
def _buscar_cpf_por_matricula(matricula: str) -> str | None:
    try:
        df = sheets.obter_coluna_aba(nome_aba="SERVIDORES")

        if df.empty:
            logger.warning("⚠️ Aba 'SERVIDORES' vazia ou inacessível.")
            return None

        col_matricula = next(
            (c for c in df.columns if "matr" in c.lower()), None
        )
        col_cpf = next(
            (c for c in df.columns if "cpf" in c.lower()), None
        )

        if not col_matricula or not col_cpf:
            logger.warning(f"⚠️ Colunas esperadas não encontradas. Disponíveis: {list(df.columns)}")
            return None

        correspondencia = df[df[col_matricula].astype(str).str.strip() == matricula]

        if correspondencia.empty:
            logger.warning(f"⚠️ Matrícula '{matricula}' não localizada na aba SERVIDORES.")
            return None

        cpf = str(correspondencia.iloc[0][col_cpf]).strip()
        logger.info(f"✅ CPF encontrado para matrícula '{matricula}': {cpf}")
        return cpf

    except Exception as e:
        logger.error(f"❌ Erro ao buscar CPF na aba SERVIDORES: {e}")
        return None
 
def _preparar_dataframe_inativacao(excel_path: str) -> tuple[pd.DataFrame | None, str | None]:
    """
    Lê o arquivo (.xlsx/.xls/.csv com ';'), normaliza matrículas, busca CPFs
    na aba 'Servidores' e garante a coluna de controle 'Inativado'.
    Retorna (df, None) em sucesso ou (None, mensagem_erro) em falha.
    """
    if not os.path.exists(excel_path):
        return None, f"❌ Arquivo não encontrado: {excel_path}"
 
    extensao = os.path.splitext(excel_path)[1].lower()
    try:
        if extensao == ".csv":
            df = pd.read_csv(excel_path, sep=";", dtype=str, encoding="utf-8-sig")
        elif extensao in (".xlsx", ".xls"):
            df = pd.read_excel(excel_path, dtype=str)
        else:
            return None, f"❌ Formato não suportado: '{extensao}'. Use .csv, .xlsx ou .xls."
    except Exception as e:
        return None, f"❌ Erro ao ler o arquivo: {e}"
 
    df.columns = df.columns.str.strip()
 
    colunas_obrigatorias = ["Matrícula", "Data Demissão", "Nome"]
    ausentes = [c for c in colunas_obrigatorias if c not in df.columns]
    if ausentes:
        return None, f"❌ Colunas obrigatórias ausentes: {ausentes}"
 
    df["Matrícula"] = df["Matrícula"].apply(_normalizar_matricula)
 
    if "Inativado" not in df.columns:
        df["Inativado"] = ""
 
    if "CPF" not in df.columns:
        df["CPF"] = ""
 
    sem_cpf = df[df["CPF"].astype(str).str.strip().isin(["", "nan", "None"])].index
    for idx in sem_cpf:
        cpf = _buscar_cpf_por_matricula(df.at[idx, "Matrícula"])
        df.at[idx, "CPF"] = cpf if cpf else ""
 
    try:
        df.to_excel(excel_path if excel_path.endswith(".xlsx") else os.path.splitext(excel_path)[0] + ".xlsx", index=False)
        logger.info(f"📄 Arquivo preparado e salvo com CPFs.")
    except Exception as e:
        logger.warning(f"⚠️ Não foi possível salvar o arquivo preparado: {e}")
 
    return df, None