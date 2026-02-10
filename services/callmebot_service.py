import requests
from config.loaders import settings
from config_network import PROXIES_OFF


class _SendMessageService:
    """
    Serviço de envio de mensagens via CallMeBot.
    Objeto chamável + métodos semânticos (Success, Fail, Error, etc.).
    """

    def __init__(self):
        cfg = settings.get("callmeBot", {})

        self.api_key = cfg.get("API_KEY")
        self.phone = cfg.get("phone")
        self.url = cfg.get("URL")

        if not all([self.api_key, self.phone, self.url]):
            raise RuntimeError("Configurações do CallMeBot incompletas no settings.json")

        self._message = None

    def __call__(self, message: str):
        if not message or not isinstance(message, str):
            raise ValueError("A mensagem deve ser uma string não vazia")

        self._message = message
        return self 

    def _send(self, prefix: str) -> bool:
        if not self._message:
            raise RuntimeError("Nenhuma mensagem definida. Use sendMessage('texto') antes.")

        full_message = f"{prefix} {self._message}"

        params = {
            "phone": self.phone,
            "text": full_message,
            "apikey": self.api_key
        }

        try:
            response = requests.get(self.url, params=params, timeout=10, proxies=PROXIES_OFF)

            if response.status_code == 200:
                self._message = None 
                return True

            print(
                f"[CALLMEBOT] Erro HTTP {response.status_code} | "
                f"Resposta: {response.text}"
            )
            return False

        except requests.exceptions.RequestException as e:
            print(f"[CALLMEBOT] Erro de comunicação: {e}")
            return False

    def Success(self) -> bool:
        return self._send("SUCESSO:")

    def Fail(self) -> bool:
        return self._send("FALHA:")

    def Error(self) -> bool:
        return self._send("ERRO:")

    def Warning(self) -> bool:
        return self._send("AVISO:")

    def Info(self) -> bool:
        return self._send("INFO:")

sendMessage = _SendMessageService()
