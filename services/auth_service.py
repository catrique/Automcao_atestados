import json
import os
import tempfile
import time
import hashlib
import base64
import platform
import getpass
from cryptography.fernet import Fernet
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.common.keys import Keys
from services.utils_service import logger
from services.utils_service import ErrorTranslator, OperationResult
from config_network import PROXIES_OFF
from config.loaders import get_config, reload_settings, update_settings

def _gerar_chave_unica():
    semente = f"{platform.node()}_{getpass.getuser()}"
    hash_obj = hashlib.sha256(semente.encode()).digest()
    return base64.urlsafe_b64encode(hash_obj)

cipher = Fernet(_gerar_chave_unica())

def atualizar_token_betha() -> OperationResult:
    logger.info("\n🚀 Iniciando atualização do token Betha via Selenium...")
    reload_settings()

    login_cripto = get_config('betha', 'user', 'admin', 'LOGIN')
    senha_cripto = get_config('betha', 'user', 'admin', 'PASSWORD')

    login_real = descriptografar(login_cripto)
    senha_real = descriptografar(senha_cripto)

    logger.info(f"Usuário Betha capturado: {login_real}")

    if not login_real or not senha_real:
        return OperationResult.fail("❌ Credenciais inválidas ou não encontradas.")

    driver = None
    
    try:
        options = Options()
        options.add_argument("--window-size=1920,1080")
        options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        caminho_driver = os.path.join(os.getenv('APPDATA'), 'AutomacaoAtestadosDriver')
        if not os.path.exists(caminho_driver):
            os.makedirs(caminho_driver, exist_ok=True)

        pasta_temp = os.path.join(tempfile.gettempdir(), "chromedriver_data")
        if not os.path.exists(pasta_temp):
            os.makedirs(pasta_temp, exist_ok=True)

        os.environ['WDM_LOCAL'] = '0' 
        os.environ['WDM_PATH'] = pasta_temp 

        logger.info(f"🚗 Verificando driver em: {pasta_temp}")
        
        driver_path = ChromeDriverManager().install()
        
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        driver.execute_cdp_cmd('Network.enable', {})
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        wait = WebDriverWait(driver, 35) 
        url_aso = "https://rh.betha.cloud/#/entidades/ZGF0YWJhc2U6MTE5NyxlbnRpdHk6MTAwNjU=/modulos/sst/executando/processos/aso"
        
        logger.info("🌐 Abrindo página...")
        driver.get(url_aso)
        
        configurar_e_autenticar_proxy()

        if "login" in driver.current_url:
            logger.info("🔑 Realizando login automático (dados descriptografados)...")
            wait.until(EC.presence_of_element_located((By.ID, "login:btAcessar")))
            
            campo_user = driver.find_element(By.ID, "login:iUsuarios")
            campo_user.send_keys(Keys.CONTROL + "a")
            campo_user.send_keys(login_real)
            
            campo_pass = driver.find_element(By.ID, "login:senha")
            campo_pass.send_keys(senha_real)   
            
            driver.find_element(By.ID, "login:btAcessar").click()
            wait.until(lambda d: "login" not in d.current_url)
            driver.get(url_aso)

        token_capturado = None
        user_access_capturado = None

        for tentativa in range(1, 5):
            logger.info(f"🔄 Tentativa {tentativa}: Buscando headers...")
            time.sleep(5)
            
            try:
                btn_refresh = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[ng-click='vm.search()']")))
                btn_refresh.click()
                time.sleep(5)
            except: pass

            logs = driver.get_log('performance')
            for entry in logs:
                log_data = json.loads(entry['message'])['message']
                if log_data['method'] == 'Network.requestWillBeSent':
                    headers = log_data.get('params', {}).get('request', {}).get('headers', {})
                    auth = headers.get('Authorization') or headers.get('authorization')
                    u_acc = headers.get('User-Access') or headers.get('user-access')

                    if auth and "Bearer" in auth and u_acc:
                        token_capturado = auth
                        user_access_capturado = u_acc
                        break
            if token_capturado: break

        if token_capturado:
            token_limpo = str(token_capturado).replace('"', '').strip()
            user_access_limpo = str(user_access_capturado).replace('"', '').strip()
            update_settings("betha,api,authorization", token_limpo)
            update_settings("betha,api,user_access", user_access_limpo)

            print("✅ Token e User-Access atualizados com sucesso!")

            return OperationResult.ok("✅ Token Betha capturado e atualizado!")
        
        return OperationResult.fail("❌ Não conseguiu capturar o Token.")

    except SessionNotCreatedException:
        return OperationResult.fail("❌ Erro: Verifique se o Chrome está aberto em outra tarefa.")
    except Exception as e:
        return ErrorTranslator.traduzir(e)
    finally:
        if driver:
            driver.quit()

def atualizar_credenciais(betha_login, betha_senha, soc_email, soc_senha, soc_virtual, proxy_user, proxy_senha) -> OperationResult:
    """Criptografa e salva as credenciais usando a função dinâmica update_settings."""
    try:
        credenciais = {
            "betha,user,admin,LOGIN": betha_login,
            "betha,user,admin,PASSWORD": betha_senha,
            "soc,user,admin,LOGIN": soc_email,
            "soc,user,admin,PASSWORD": soc_senha,
            "soc,user,admin,SENHA_VIRTUAL": soc_virtual,
            "proxy,PROXY_USER": proxy_user,
            "proxy,PROXY_PASS": proxy_senha
        }

        for path, valor in credenciais.items():
            if valor:  
                update_settings(path, criptografar(valor))

        return OperationResult.ok("✅ Credenciais salvas e criptografadas com sucesso!")

    except Exception as e:
        return ErrorTranslator.traduzir(e)

def criptografar(texto: str) -> str:
    if not texto: return ""
    return cipher.encrypt(str(texto).strip().encode()).decode()

def descriptografar(texto_cripto: str) -> str:
    if not texto_cripto: return ""
    try:
        return cipher.decrypt(str(texto_cripto).encode()).decode()
    except Exception:
        return ""

def configurar_e_autenticar_proxy():
    print("Configurando proxy...")
    PROXY_HOST = get_config('proxy', 'PROXY_HOST', default='')
    PROXY_PORT = get_config('proxy', 'PROXY_PORT', default='')
    PROXY_USER = descriptografar(get_config('proxy', 'PROXY_USER', default=''))
    PROXY_PASS = descriptografar(get_config('proxy', 'PROXY_PASS', default=''))

    if not PROXY_HOST or not PROXY_PORT:
        logger.warning("⚠️ Proxy não configurado no settings.json.")
        return None

    logger.info(f"🔐 Usando proxy {PROXY_HOST}:{PROXY_PORT}")
    proxy_argument = f"--proxy-server=http://{PROXY_HOST}:{PROXY_PORT}"
    if PROXY_USER and PROXY_PASS:
        logger.info("⏳ Aguardando popup de autenticação do proxy...")
        logger.info("🚨 NÃO MEXA NO MOUSE/TECLADO nos próximos segundos!")

        try:
            time.sleep(3)

            screen_width, screen_height = pyautogui.size()
            pyautogui.click(screen_width // 2, screen_height // 2)
            time.sleep(0.5)

            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('delete')
            pyautogui.write(PROXY_USER, interval=0.05)
            pyautogui.press('tab')
            time.sleep(0.3)

            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('delete')
            pyautogui.write(PROXY_PASS, interval=0.05)
            pyautogui.press('enter')

            logger.info("✅ Proxy autenticado com sucesso")
            time.sleep(3)

        except Exception as e:
            logger.warning(f"⚠️ Erro ao autenticar proxy: {e}")
            time.sleep(5)
    return proxy_argument
