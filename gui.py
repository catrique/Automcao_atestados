import customtkinter as ctk
import os
import threading
import time
from tkinter import messagebox
from services.utils_service import configurar_log_gui
from config_global import sheets
from services.betha_service import BethaService
from services.auth_service import atualizar_credenciais, atualizar_token_betha
from repositories.update_data import sincronizar_bases_betha
from services.soc_service import executar_fluxo_soc
from services.validation_service import processar_validacoes_excel

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerenciador de Atestados - CRESST")
        self.geometry("900x500")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="AUTOMAÇÃO RH", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 30))

        self.btn_sync = ctk.CTkButton(self.sidebar_frame, text="1. Atualizar Bases de consulta (CID, CRM..)", command=self.thread_sync)
        self.btn_sync.grid(row=1, column=0, padx=20, pady=10)

        self.btn_soc = ctk.CTkButton(self.sidebar_frame, text="2. Baixar relatório do soc e importar planilha", command=self.thread_soc)
        self.btn_soc.grid(row=2, column=0, padx=20, pady=10)

        self.btn_enviar = ctk.CTkButton(self.sidebar_frame, text="3. Ler planilha e Lançar atestado no Betha ", fg_color="#1f538d", command=self.confirmar_envio)
        self.btn_enviar.grid(row=3, column=0, padx=20, pady=10)

        self.btn_config = ctk.CTkButton(self.sidebar_frame, text="Atualizar Token Betha", fg_color="gray30", hover_color="gray20", command=self.thread_token_betha)
        self.btn_config.grid(row=5, column=0, padx=20, pady=(100, 10))

        self.btn_config = ctk.CTkButton(self.sidebar_frame, text="⚙️ Credenciais de Acesso (Betha e SOC)", fg_color="gray30", hover_color="gray20", command=self.mostrar_configuracoes)
        self.btn_config.grid(row=5, column=0, padx=20, pady=(10, 10))

        self.console_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.console_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        self.lbl_console = ctk.CTkLabel(self.console_frame, text="Console de Execução", font=ctk.CTkFont(weight="bold"))
        self.lbl_console.pack(anchor="w", pady=(0, 5))
        
        self.log_text = ctk.CTkTextbox(self.console_frame, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_text.pack(expand=True, fill="both")

        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="Credenciais de Acesso")
        configurar_log_gui(self.log)
        self.montar_campos_config()

    # def montar_campos_config(self):
    #     """Constrói os campos de input na tela de configurações."""
    #     ctk.CTkLabel(self.scroll_frame, text="Betha Cloud", font=ctk.CTkFont(weight="bold")).pack(pady=(10, 0))
    #     self.entry_betha_user = ctk.CTkEntry(self.scroll_frame, placeholder_text="Login Betha", width=350)
    #     self.entry_betha_user.pack(pady=5)
    #     self.entry_betha_pass = ctk.CTkEntry(self.scroll_frame, placeholder_text="Senha Betha", show="*", width=350)
    #     self.entry_betha_pass.pack(pady=5)

    #     ctk.CTkLabel(self.scroll_frame, text="SOC Integration", font=ctk.CTkFont(weight="bold")).pack(pady=(20, 0))
    #     self.entry_soc_user = ctk.CTkEntry(self.scroll_frame, placeholder_text="E-mail SOC", width=350)
    #     self.entry_soc_user.pack(pady=5)
    #     self.entry_soc_pass = ctk.CTkEntry(self.scroll_frame, placeholder_text="Senha SOC", show="*", width=350)
    #     self.entry_soc_pass.pack(pady=5)
    #     self.entry_soc_virt = ctk.CTkEntry(self.scroll_frame, placeholder_text="Senha Virtual (ex: 1,2,3,4)", width=350)
    #     self.entry_soc_virt.pack(pady=5)

    #     self.btn_save = ctk.CTkButton(self.scroll_frame, text="Salvar Credenciais", fg_color="#28a745", hover_color="#218838", command=self.salvar_dados)
    #     self.btn_save.pack(pady=20)

    def montar_campos_config(self):
        """Constrói os campos garantindo que os inputs de 350px fiquem centralizados no eixo."""
        
        ctk.CTkLabel(self.scroll_frame, text="Betha Cloud", font=ctk.CTkFont(weight="bold")).pack(pady=(10, 0))
        
        self.entry_betha_user = ctk.CTkEntry(self.scroll_frame, placeholder_text="Login Betha", width=350)
        self.entry_betha_user.pack(pady=5)

        f_betha_pass = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        f_betha_pass.pack(pady=5, fill="x")
        
        # Coluna 0: Espaço invisível para equilibrar (mesma largura do botão)
        # Coluna 1: O Input centralizado
        # Coluna 2: O Botão real
        f_betha_pass.grid_columnconfigure(0, weight=1) 
        f_betha_pass.grid_columnconfigure(2, weight=1) 
        
        ctk.CTkLabel(f_betha_pass, text="", width=80).grid(row=0, column=0, sticky="e")
        
        self.entry_betha_pass = ctk.CTkEntry(f_betha_pass, placeholder_text="Senha Betha", show="*", width=350)
        self.entry_betha_pass.grid(row=0, column=1, padx=0)
        
        self.btn_eye_betha = ctk.CTkButton(f_betha_pass, text="Mostrar", width=80,
                                        fg_color=("#dbdbdb", "#2b2b2b"), text_color=("#000", "#fff"),
                                        command=lambda: self.toggle_password(self.entry_betha_pass, self.btn_eye_betha))
        self.btn_eye_betha.grid(row=0, column=2, padx=(5, 0), sticky="w")

        ctk.CTkLabel(self.scroll_frame, text="SOC Integration", font=ctk.CTkFont(weight="bold")).pack(pady=(20, 0))
        
        self.entry_soc_user = ctk.CTkEntry(self.scroll_frame, placeholder_text="E-mail SOC", width=350)
        self.entry_soc_user.pack(pady=5)

        f_soc_pass = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        f_soc_pass.pack(pady=5, fill="x")
        
        f_soc_pass.grid_columnconfigure(0, weight=1)
        f_soc_pass.grid_columnconfigure(2, weight=1)
        
        ctk.CTkLabel(f_soc_pass, text="", width=80).grid(row=0, column=0, sticky="e")
        
        self.entry_soc_pass = ctk.CTkEntry(f_soc_pass, placeholder_text="Senha SOC", show="*", width=350)
        self.entry_soc_pass.grid(row=0, column=1, padx=0)
        
        self.btn_eye_soc = ctk.CTkButton(f_soc_pass, text="Mostrar", width=80,
                                        fg_color=("#dbdbdb", "#2b2b2b"), text_color=("#000", "#fff"),
                                        command=lambda: self.toggle_password(self.entry_soc_pass, self.btn_eye_soc))
        self.btn_eye_soc.grid(row=0, column=2, padx=(5, 0), sticky="w")

        self.entry_soc_virt = ctk.CTkEntry(self.scroll_frame, placeholder_text="Senha Virtual (ex: 1,2,3,4)", width=350)
        self.entry_soc_virt.pack(pady=5)


        ctk.CTkLabel(self.scroll_frame, text="Proxy", font=ctk.CTkFont(weight="bold")).pack(pady=(10, 0))
        
        self.entry_proxy_user = ctk.CTkEntry(self.scroll_frame, placeholder_text="Login Proxy", width=350)
        self.entry_proxy_user.pack(pady=5)

        f_proxy_pass = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        f_proxy_pass.pack(pady=5, fill="x")
        
        # Coluna 0: Espaço invisível para equilibrar (mesma largura do botão)
        # Coluna 1: O Input centralizado
        # Coluna 2: O Botão real
        f_proxy_pass.grid_columnconfigure(0, weight=1) 
        f_proxy_pass.grid_columnconfigure(2, weight=1) 
        
        ctk.CTkLabel(f_proxy_pass, text="", width=80).grid(row=0, column=0, sticky="e")
        
        self.entry_proxy_pass = ctk.CTkEntry(f_proxy_pass, placeholder_text="Senha Proxy", show="*", width=350)
        self.entry_proxy_pass.grid(row=0, column=1, padx=0)
        
        self.btn_eye_proxy = ctk.CTkButton(f_proxy_pass, text="Mostrar", width=80,
                                        fg_color=("#dbdbdb", "#2b2b2b"), text_color=("#000", "#fff"),
                                        command=lambda: self.toggle_password(self.entry_proxy_pass, self.btn_eye_proxy))
        self.btn_eye_proxy.grid(row=0, column=2, padx=(5, 0), sticky="w")




        self.btn_save = ctk.CTkButton(self.scroll_frame, text="Salvar Credenciais", fg_color="#28a745", 
                                    hover_color="#218838", command=self.salvar_dados, width=350)
        self.btn_save.pack(pady=20)

    def toggle_password(self, entry_field, button):
        """Alterna a visibilidade da senha e o texto do botão."""
        if entry_field.cget("show") == "*":
            entry_field.configure(show="")
            button.configure(text="Esconder")
        else:
            entry_field.configure(show="*")
            button.configure(text="Mostrar")

    def mostrar_console(self):
        self.scroll_frame.grid_forget()
        self.console_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

    def mostrar_configuracoes(self):
        self.console_frame.grid_forget()
        self.scroll_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

    def log(self, msg):
        self.log_text.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see("end")

    def salvar_dados(self):
        res = atualizar_credenciais(
            self.entry_betha_user.get(), self.entry_betha_pass.get(),
            self.entry_soc_user.get(), self.entry_soc_pass.get(), self.entry_soc_virt.get(),
            self.entry_proxy_user.get(), self.entry_proxy_pass.get()
        )
        messagebox.showinfo("Sucesso", "Configurações salvas localmente!")
        for entry in (
            self.entry_betha_user,
            self.entry_betha_pass,
            self.entry_soc_user,
            self.entry_soc_pass,
            self.entry_soc_virt,
            self.entry_proxy_user,
            self.entry_proxy_pass,
        ):
            entry.delete(0, "end")

        # --- garantir que campos de senha voltem a ficar ocultos ---
        self.entry_betha_pass.configure(show="*")
        self.entry_soc_pass.configure(show="*")
        self.entry_proxy_pass.configure(show="*")

        # --- resetar texto dos botões "olho", se existirem ---
        if hasattr(self, "btn_eye_betha"):
            self.btn_eye_betha.configure(text="Mostrar")
        if hasattr(self, "btn_eye_soc"):
            self.btn_eye_soc.configure(text="Mostrar")
        if hasattr(self, "btn_eye_proxy"):
            self.btn_eye_proxy.configure(text="Mostrar")

        # --- ação final existente ---
        self.mostrar_console()

    def confirmar_envio(self):
        if messagebox.askyesno("Revisão", "A planilha Sheets foi conferida e os 'Sim' foram marcados?"):
            self.thread_envio()

    
    def thread_sync(self):
        self.mostrar_console()
        threading.Thread(target=self.run_sync, daemon=True).start()

    def run_sync(self):
        self.log("⏳ Iniciando sincronização Betha -> Sheets...")
        try:
            res = sincronizar_bases_betha()
            if res.success: self.log(f"✅ {res.message}")
            else: self.log(f"❌ Erro: {res.message}")
        except Exception as e: self.log(f"💥 Erro crítico: {e}")
    
    def thread_soc(self):
        self.mostrar_console()
        threading.Thread(target=self.run_soc, daemon=True).start()

    def thread_token_betha(self):
        self.mostrar_console()
        threading.Thread(target=self.run_token_betha, daemon=True).start()
    def run_token_betha(self):
        try:
            res_token = atualizar_token_betha()
            
            if hasattr(res_token, 'success') and res_token.success:
                self.log("✅ Token Betha atualizado com sucesso!")
            else:
                msg = res_token.message if hasattr(res_token, 'message') else str(res_token)
                self.log(f"❌ Falha no token Betha: {msg}")
                
        except Exception as e:
            self.log(f"💥 Erro crítico na interface: {e}")

    def run_soc(self):
        try:
            res_soc = executar_fluxo_soc()
            if res_soc and res_soc.success:
            # if True:
                excel = res_soc.data
                # excel = "\\\\10.1.1.50\\ADM_Cresst\\Atestados_Laudar\\22-01-2026\\Relatorio_licensas_medicas_22-01-2026.xlsx"
                self.log(f"🔍 Validando arquivo: {os.path.basename(excel)}")
                
                output_op = processar_validacoes_excel(excel)
                if output_op.success:
                    caminho_validado = output_op.data
                    self.log("📤 Importando para Google Sheets...")
                    pasta = os.path.dirname(caminho_validado)
                    sheets.importar_excel_para_aba(pasta, 'licensas_medicas')
                    self.log("✅ Processo SOC -> Sheets concluído!")
                else:
                    self.log(f"❌ Erro na validação: {output_op.message}")
            else:
                self.log(f"❌ Falha no SOC: {res_soc.message if res_soc else 'Erro de conexão'}")
        except Exception as e: self.log(f"💥 Erro no SOC: {e}")

    def thread_envio(self):
        self.mostrar_console()
        threading.Thread(target=self.run_envio, daemon=True).start()

    def run_envio(self):
        dialogo = ctk.CTkInputDialog(text="A partir de qual linha ler a planilha?", title="Ponto de Início")
        entrada = dialogo.get_input()
        
        try:
            linha_inicio = int(entrada) if entrada else 2
        except:
            self.log("❌ Valor inválido. Use números.")
            return

        self.log(f"⚙️ Preparando API Betha (Linha {linha_inicio})...")
        try:
            atualizar_token_betha()
            betha = BethaService()

            resultado_bruto = sheets.ler_planilha_para_automacao("licensas_medicas", linha_inicio)
            if not resultado_bruto:
                self.log("ℹ️ Nenhum dado pendente encontrado.")
                return

            res_lote = betha.processar_lote_planilha(resultado_bruto)
            if not res_lote.success:
                self.log(f"⚠️ {res_lote.message}")
                return

            res_payloads = betha.gerar_payloads_lote(res_lote.data)
            if not res_payloads.success:
                self.log(f"⚠️ {res_payloads.message}")
                return

            lista_payloads = res_payloads.data
            self.log(f"📦 {len(lista_payloads)} atestados prontos para envio.")

            for p in lista_payloads:
                time.sleep(3)
                cod_ficha = p.get("numeroAtestado")
                p_envio = p.copy()
                p_envio["numeroAtestado"] = None 
                
                self.log(f"📤 Enviando Ficha: {cod_ficha}...")
                result = betha.enviar_atestado(p_envio)
                
                foi_sucesso = result.success if hasattr(result, 'success') else result
                
                if foi_sucesso:
                    sheets.marcar_status_na_planilha(id_busca=cod_ficha, erro=False)
                    self.log(f"✅ Ficha {cod_ficha}: Sucesso!")
                else:
                    erro_msg = result.message if hasattr(result, 'message') else "Erro na API"
                    sheets.marcar_status_na_planilha(id_busca=cod_ficha, erro=True)
                    self.log(f"❌ Ficha {cod_ficha}: Falhou ({erro_msg})")

            self.log("🏁 Fim do processo de envio.")

        except Exception as e:
            self.log(f"💥 Erro Crítico: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()