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

def limpar_cpf(cpf):
    """Remove pontuação do CPF para comparação (ex: 123.456.789-00 -> 12345678900)."""
    if not cpf or pd.isna(cpf): return ""
    return re.sub(r'[^0-9]', '', str(cpf))

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

def validar_duplicidade_matricula(df_excel, df_servidores):
    """
    Verifica se o CPF do Excel possui mais de um vínculo (matrícula) na aba SERVIDORES.
    """
    total_erros = 0
    erros_resultado = []

    contagem_servidores = {}
    
    if not df_servidores.empty:
        col_cpf_ref = None
        for col in ['cpf', 'CPF', 'Cpf']:
            if col in df_servidores.columns:
                col_cpf_ref = col
                break
        
        if col_cpf_ref:
            serie_cpfs_limpos = df_servidores[col_cpf_ref].apply(limpar_cpf)
            contagem_servidores = serie_cpfs_limpos.value_counts().to_dict()
        else:
            logger.error("⚠️ Coluna 'cpf' não encontrada na aba SERVIDORES")

    for idx, row in df_excel.iterrows():
        erro_msg = None
        cpf_valor = row.get("CPF") or row.get("C.P.F") or row.get("cpf") or row.get("Cpf")
        cpf_limpo_excel = limpar_cpf(cpf_valor)

        if cpf_limpo_excel:
            qtd_na_base = contagem_servidores.get(cpf_limpo_excel, 0)
            if qtd_na_base > 1:
                erro_msg = f"Servidor possui {qtd_na_base} matriculas"
                total_erros += 1
        
        erros_resultado.append(erro_msg)

    return erros_resultado, total_erros

def validar_vinculo_matricula(df_excel, df_servidores):
    """
    Busca o vínculo na aba SERVIDORES baseado na Matrícula do Excel.
    Retorna uma lista com os vínculos encontrados para atualizar o DataFrame.
    """
    mapa_vinculos = {
        str(row['Matricula']).strip(): str(row['Vínculo']).strip()
        for _, row in df_servidores.iterrows()
        if 'Matricula' in row and 'Vínculo' in row
    }

    vinculos_resultado = []

    for _, row in df_excel.iterrows():
        matricula_excel = str(row.get("Matrícula Funcionário") or "").strip()
        vinculo_encontrado = mapa_vinculos.get(matricula_excel)
        if vinculo_encontrado:
            vinculos_resultado.append(vinculo_encontrado)
        else:
            vinculos_resultado.append(row.get("Vínculo") or "")
            
    return vinculos_resultado

def validar_situacao_matricula(df_excel, df_servidores):
    mapa_situacoes = {
        str(row['Matricula']).strip().lstrip('0'): str(row['Situação']).strip()
        for _, row in df_servidores.iterrows()
        if 'Matricula' in row and 'Situação' in row
    }

    situacoes_resultado = []

    for _, row in df_excel.iterrows():
        matricula_excel = str(row.get("Matrícula Funcionário") or "").strip().lstrip('0')
        situacao_encontrada = mapa_situacoes.get(matricula_excel)
        
        if situacao_encontrada:
            situacoes_resultado.append(situacao_encontrada)
        else:
            situacoes_resultado.append(row.get("Situação BETHA") or "")  
            
    return situacoes_resultado
def processar_validacoes_excel(caminho_excel) -> OperationResult:
    """
    Orquestra as validações (Médico, CID e Múltiplos Vínculos) e salva no Excel.
    """
    try:
        if not os.path.exists(caminho_excel):
            return OperationResult.fail(f"📂 Arquivo Excel não encontrado: {os.path.basename(caminho_excel)}")

        logger.info(f"🧐 Iniciando validações: {os.path.basename(caminho_excel)}")

        mapa_medicos_bruto = sheets.obter_mapa_validacao("MEDICOS", "numeroConselho", "nome")
        mapa_cid_bruto = sheets.obter_mapa_validacao("CID", "codigo", "descricao")
        
        df_servidores_ref = sheets.obter_coluna_aba("SERVIDORES")

        if not mapa_medicos_bruto or not mapa_cid_bruto or df_servidores_ref.empty:
             logger.warning("⚠️ Alguma base de validação não pôde ser carregada completamente.")

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
        df["Vínculo"] = validar_vinculo_matricula(df, df_servidores_ref)
        df["Situação BETHA"] = validar_situacao_matricula(df, df_servidores_ref)
        lista_erros_medico, count_m = validar_medico_crm(df, mapa_medicos_ref)
        lista_erros_cid, count_c = validar_cid(df, set_cids_ref)
        lista_erros_matricula, count_mat = validar_duplicidade_matricula(df, df_servidores_ref)

        erros_finais = []
        for e_med, e_cid, e_mat in zip(lista_erros_medico, lista_erros_cid, lista_erros_matricula):
            erros_da_linha = [e for e in [e_med, e_cid, e_mat] if e] 
            erros_finais.append(" | ".join(erros_da_linha) if erros_da_linha else "")

        df["ERROS"] = erros_finais
        
        try:
            df.to_excel(caminho_excel, index=False)
        except Exception as e:
            return OperationResult.fail(f"Erro ao salvar arquivo validado: {str(e)}")

        total_problemas = count_m + count_c + count_mat
        if total_problemas > 0:
            return OperationResult.ok(f"Validação concluída: {total_problemas} alertas encontrados.", data=caminho_excel)
        
        return OperationResult.ok("Planilha validada com sucesso! Nenhum problema encontrado.", data=caminho_excel)

    except Exception as e:
        logger.error(f"Erro crítico na validação: {e}")
        return OperationResult.fail(f"Erro ao validar: {str(e)}")
