import pandas as pd
import unicodedata
import re
import os
from config_global import sheets
from services.utils_service import OperationResult, ErrorTranslator
from services.utils_service import logger
from config_network import sessao_limpa

def normalizar_texto(texto):
    if not texto or pd.isna(texto): return ""
    texto = str(texto).upper().strip()
    for titulo in ["DR. ", "DRA. ", "DR ", "DRA "]:
        if texto.startswith(titulo):
            texto = texto.replace(titulo, "", 1)
    nfkd_form = unicodedata.normalize('NFKD', texto)
    texto_sem_acentos = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return " ".join(texto_sem_acentos.replace('Ç', 'C').split())

def limpar_crm(crm):
    if crm is None or pd.isna(crm): return ""
    return str(crm).split('.')[0].strip()

def limpar_cid(cid):
    if not cid or pd.isna(cid): return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(cid)).lower()


def validar_medico_crm(df, mapa_medicos):
    """Valida médico e CRM, adicionando erros à lista da linha."""
    total_erros = 0
    erros_resultado = [] 

    for idx, row in df.iterrows():
        erro_medico = None
        nome_excel = normalizar_texto(row.get("Médico assistente"))
        crm_excel = limpar_crm(row.get("CRM Médico assistente"))

        if nome_excel or crm_excel:
            if crm_excel in mapa_medicos:
                nome_correto = mapa_medicos[crm_excel]
                if nome_excel != nome_correto and (nome_correto not in nome_excel and nome_excel not in nome_correto):
                    erro_medico = f"Médico divergente (Betha: {nome_correto})"
                    total_erros += 1
            elif crm_excel:
                erro_medico = f"CRM {crm_excel} não localizado"
                total_erros += 1
        
        erros_resultado.append(erro_medico)
    
    return erros_resultado, total_erros

def validar_cid(df, set_cids_validos):
    """Valida o código CID, adicionando erros à lista da linha."""
    total_erros = 0
    erros_resultado = []

    for idx, row in df.iterrows():
        erro_cid = None
        cid_excel_raw = row.get("CID")
        cid_excel_limpo = limpar_cid(cid_excel_raw)

        if cid_excel_limpo:
            if cid_excel_limpo not in set_cids_validos:
                erro_cid = f"CID {cid_excel_raw} inválido"
                total_erros += 1
        else:
            erro_cid = "CID não informado"
            total_erros += 1
            
        erros_resultado.append(erro_cid)

    return erros_resultado, total_erros

def processar_validacoes_excel(caminho_excel) -> OperationResult:
    """
    Orquestra as validações e salva no Excel. Retorna OperationResult.
    """
    try:
        if not os.path.exists(caminho_excel):
            return OperationResult.fail(f"📂 Arquivo Excel não encontrado: {os.path.basename(caminho_excel)}")

        logger.info(f"🧐 Iniciando validações: {os.path.basename(caminho_excel)}")

        mapa_medicos_bruto = sheets.obter_mapa_validacao("MEDICOS", "numeroConselho", "nome")
        mapa_cid_bruto = sheets.obter_mapa_validacao("CID", "codigo", "descricao")

        if not mapa_medicos_bruto or not mapa_cid_bruto:
             return OperationResult.fail("📊 Falha ao carregar bases de validação da planilha. Verifique a conexão.")

        mapa_medicos_ref = {}
        for crm, nome_bruto in mapa_medicos_bruto.items():
            crm_limpo = limpar_crm(crm)
            nome_norm = normalizar_texto(nome_bruto)
            if crm_limpo:
                if crm_limpo in mapa_medicos_ref:
                    nome_atual = mapa_medicos_ref[crm_limpo]
                    if (nome_norm.startswith('J') and not nome_atual.startswith('J')) or (len(nome_norm) > len(nome_atual)):
                        mapa_medicos_ref[crm_limpo] = nome_norm
                else:
                    mapa_medicos_ref[crm_limpo] = nome_norm

        set_cids_ref = {limpar_cid(codigo) for codigo in mapa_cid_bruto.keys() if codigo}

        try:
            df = pd.read_excel(caminho_excel, header=0)
        except PermissionError:
            return OperationResult.fail(f"🚫 O arquivo '{os.path.basename(caminho_excel)}' está aberto. Feche-o para validar.")

        df.columns = df.columns.str.strip()

        lista_erros_medico, count_m = validar_medico_crm(df, mapa_medicos_ref)
        lista_erros_cid, count_c = validar_cid(df, set_cids_ref)

        erros_finais = []
        for e_med, e_cid in zip(lista_erros_medico, lista_erros_cid):
            erros_da_linha = [e for e in [e_med, e_cid] if e] 
            erros_finais.append(" | ".join(erros_da_linha))

        df["ERROS"] = erros_finais

        df.to_excel(caminho_excel, index=False)
        
        msg_sucesso = f"✅ Validação concluída! Médico: {count_m} erros | CID: {count_c} erros."
        return OperationResult.ok(msg_sucesso, data=caminho_excel)

    except Exception as e:
        return ErrorTranslator.traduzir(e)