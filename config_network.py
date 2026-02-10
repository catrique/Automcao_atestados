import sys
import os
import requests

sys.modules['OpenSSL'] = None
sys.modules['OpenSSL.crypto'] = None

PROXIES_OFF = {
    'http': None,
    'https': None
}

sessao_limpa = requests.Session()
sessao_limpa.trust_env = False
sessao_limpa.proxies = PROXIES_OFF

os.environ['NO_PROXY'] = '*'