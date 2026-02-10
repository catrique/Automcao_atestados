import os
import sys
import time
from repositories.update_data import sincronizar_bases_betha
from services.soc_service import executar_fluxo_soc
from services.validation_service import processar_validacoes_excel
from config_global import sheets
from services.betha_service import BethaService
from services.auth_service import atualizar_credenciais, atualizar_token_betha

def menu():
    while True:
        print("\n========================================")
        print("      SISTEMA DE AUTOMAÇÃO - MENU")
        print("========================================")
        print("[1] Atualizar Bases de Dados (API Betha -> Sheets)")
        print("[2] Baixar Relatório SOC e importar para Google Sheets")
        print("[3] Importar Atestado para o Betha")
        print("[4] Atualizar Credenciais")
        print("[0] Sair")
        print("========================================")
        
        opcao = input("Escolha uma opção: ")

        if opcao == "1":
            print("\n🔄 Sincronizando bases Betha...")
            resultado = sincronizar_bases_betha()
            
            if resultado.success:
                print(f"✅ {resultado.message}")
            else:
                print(f"❌ Erro na Sincronização: {resultado.message}")

        elif opcao == "2":
            print("\n🚀 Executando atualização SOC...")
            res_soc = executar_fluxo_soc()
            if res_soc and res_soc.success:
                excel = res_soc.data
            else:
                print(f"❌ Falha no SOC: {res_soc.message if res_soc else 'Erro desconhecido'}")
                excel = None

            # --- LINHAS PARA TESTE MANUAL (Bypass do SOC) ---
            # Se precisar testar apenas a validação e upload, comente o bloco acima e use as linhas abaixo:
            # excel = "C:\\Users\\42706671840\\Documents\\Automacao_Atestados\\workspace\\Downloads\\Relatorio_licensas_medicas_16_01_2026.xlsx"
            # excel = "\\\\10.1.1.50\\ADM_Cresst\\Atestados_Laudar\\22-01-2026\\Relatorio_licensas_medicas_22-01-2026_Copia.xlsx"

            if excel:
                print(f"🔍 Iniciando validação do arquivo: {excel}")
                output_op = processar_validacoes_excel(excel)
                if output_op and hasattr(output_op, 'success') and output_op.success:
                    caminho_validado = output_op.data
                    print(f"✅ Validações concluídas com sucesso!")
                    
                    print(f"📤 Enviando para o Google Sheets...")
                    pasta_do_arquivo = os.path.dirname(caminho_validado)
                    sheets.importar_excel_para_aba(pasta_do_arquivo, 'licensas_medicas')
                    print("✅ Processo SOC -> Sheets finalizado!")
                else:
                    msg_erro = output_op.message if hasattr(output_op, 'message') else "Erro desconhecido"
                    print(f"❌ Falha ao processar as validações do Excel: {msg_erro}")

        elif opcao == "3":
            print("\n⚙️ Iniciando integração com API Betha...")
            atualizar_token_betha()
            betha = BethaService()
            
            resultado = sheets.ler_planilha_para_automacao("licensas_medicas", 33)
            
            if not resultado:
                print("ℹ️ Nenhum atestado pendente encontrado na planilha.")
                continue

            resultado_lote = betha.processar_lote_planilha(resultado)

            if resultado_lote.success:
                linhas_validas = resultado_lote.data
                res_payloads = betha.gerar_payloads_lote(linhas_validas)
                
                if not res_payloads.success:
                    print(f"⚠️ Interrompendo: {res_payloads.message}")
                    return
                    
                lista_final_payloads = res_payloads.data
                
            else:
                print(f"⚠️ Interrompendo: {resultado_lote.message}")
                return 

            print(f"📦 Total de atestados para processar: {len(lista_final_payloads)}")

            for payload in lista_final_payloads: 
                cod_ficha = "Desconhecido" 
                try:
                    time.sleep(3) 
                    cod_ficha = payload.get("numeroAtestado")
                    
                    payload_envio = payload.copy()
                    payload_envio["numeroAtestado"] = None 
                    result = betha.enviar_atestado(payload_envio)
                    status_envio = result.success if hasattr(result, 'success') else result

                    if status_envio:
                        sheets.marcar_status_na_planilha(
                            id_busca=cod_ficha,
                            erro=False
                        )
                        print(f"✅ Atestado {cod_ficha} enviado com sucesso!")
                    else:
                        msg_erro = result.message if hasattr(result, 'message') else "Erro desconhecido"
                        sheets.marcar_status_na_planilha(
                            id_busca=cod_ficha,
                            erro=True
                        )
                        print(f"❌ Falha no envio do atestado {cod_ficha}: {msg_erro}")

                except Exception as e:
                    print(f"❌ Erro crítico no loop de envio para ficha {cod_ficha}: {e}")

        elif opcao == "4":
            print("\n🔑 Atualização de Credenciais")
            email_betha = input("Digite o login do Betha: ")
            senha_betha = input("Digite a senha do Betha: ")
            email_soc = input("Digite o login do SOC: ")
            senha_soc = input("Digite a senha do SOC: ")
            senha_virtual_soc = input("Digite a senha virtual do SOC (ex: 1,2,3,4): ")
            resultado = atualizar_credenciais(email_betha, senha_betha, email_soc, senha_soc, senha_virtual_soc)
            
            if resultado and hasattr(resultado, 'success') and resultado.success:
                print("✅ Credenciais atualizadas com sucesso!")
            elif resultado and hasattr(resultado, 'success'):
                print(f"❌ Erro ao atualizar: {resultado.message}")
            else:
                print("✅ Credenciais gravadas (verifique se o arquivo JSON mudou).")

        elif opcao == "0":
            print("Encerrando o sistema...")
            sys.exit()
            
        else:
            print("⚠️ Opção inválida! Tente novamente.")


if __name__ == "__main__":
    try:
        menu()
    except KeyboardInterrupt:
        print("\n\nSaindo do sistema...")
        sys.exit()
    except Exception as e:
        print(f"\n💥 Erro crítico ao iniciar sistema: {e}")
        sys.exit()