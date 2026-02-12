import json
import os
import sys

settings = {}
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    SETTINGS_PATH = os.path.join(BASE_DIR, "config", "settings.json")
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    SETTINGS_PATH = os.path.join(BASE_DIR, "config", "settings.json")


class ConfigLoader:
    @staticmethod
    def load_settings():
        """Lê o arquivo JSON do disco e retorna os dados."""
        if not os.path.exists(SETTINGS_PATH):
            print(f"📂 Arquivo não encontrado: {SETTINGS_PATH}")
            return None

        try:
            with open(SETTINGS_PATH, "r", encoding="utf8") as f:
                return json.load(f)
        except Exception as e:
            print(f"❌ Erro ao carregar JSON: {e}")
            return None


def reload_settings():
    """Atualiza a memória mantendo a mesma referência de objeto."""
    global settings

    if not os.path.exists(SETTINGS_PATH):
        return

    try:
        with open(SETTINGS_PATH, "r", encoding="utf8") as f:
            dados = json.load(f)
            settings.clear()
            settings.update(dados)
    except Exception as e:
        print(f"❌ Erro ao carregar JSON: {e}")


def get_config(*keys, default=None):
    """
    Busca valores aninhados no dicionário de settings.
    Recarrega se settings estiver vazio ou se for uma busca de credenciais.
    """
    global settings
    
    if not settings:
        reload_settings()

    val = settings
    try:
        for key in keys:
            if not isinstance(val, dict):
                return default
            val = val.get(key)
        
        return val if val is not None else default
    except Exception:
        return default


def update_settings(path_string, novo_valor, salvar_no_disco=True):
    """
    Atualiza um valor no dicionário global e opcionalmente salva no JSON.
    Uso: update_settings("betha,user,admin,LOGIN", "novo_usuario")
    """
    global settings

    keys = [k.strip() for k in path_string.split(",")]

    ptr = settings
    try:
        for key in keys[:-1]:
            if key not in ptr:
                ptr[key] = {}
            ptr = ptr[key]

        ptr[keys[-1]] = novo_valor

        if salvar_no_disco:
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
        reload_settings()
        return True
    except Exception as e:
        print(f"❌ Erro ao atualizar configuração: {e}")
        return False


reload_settings()
