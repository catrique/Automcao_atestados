import os
import time
import threading

import customtkinter as ctk
from tkinter import messagebox

from config_global import sheets
from services.auth_service import atualizar_credenciais, atualizar_token_betha
from services.betha_service import BethaService
from services.soc_service import (
    executar_fluxo_soc,
    exportar_relatorio_por_periodo,
    retomar_relatorio_por_periodo,
    listar_relatorios_pendentes,
)
from services.utils_service import configurar_log_gui, registrar_limpar_gui
from services.validation_service import processar_validacoes_excel
from repositories.update_data import sincronizar_bases_betha

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self._configurar_janela()
        self._construir_sidebar()
        self._construir_console()
        self._construir_painel_configuracoes()
        self._construir_painel_exportar_periodo()
        self._construir_painel_datas_soc()
        self._registrar_servicos_log()


    def _configurar_janela(self):
        """Define título, tamanho e layout de colunas/linhas da janela principal."""
        self.title("Gerenciador de Atestados - CRESST")
        self.geometry("900x500")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)


    def _construir_sidebar(self):
        """Monta a barra lateral com todos os botões de ação."""
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")

        ctk.CTkLabel(
            self.sidebar_frame,
            text="AUTOMAÇÃO RH",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 30))

        self.btn_sincronizar_bases = ctk.CTkButton(
            self.sidebar_frame,
            text="1. Atualizar Bases de consulta (CID, CRM..)",
            command=self.ao_clicar_sincronizar_bases,
        )
        self.btn_sincronizar_bases.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.btn_baixar_soc = ctk.CTkButton(
            self.sidebar_frame,
            text="2. Baixar relatório do SOC e importar planilha",
            command=self.ao_clicar_baixar_soc,
        )
        self.btn_baixar_soc.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        self.btn_exportar_periodo = ctk.CTkButton(
            self.sidebar_frame,
            text="3. Exportar relatório por período",
            command=self.ao_clicar_exportar_periodo,
        )
        self.btn_exportar_periodo.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        self.btn_lancar_atestado = ctk.CTkButton(
            self.sidebar_frame,
            text="4. Ler planilha e Lançar atestado no Betha",
            command=self.ao_clicar_lancar_atestado,
        )
        self.btn_lancar_atestado.grid(row=4, column=0, padx=20, pady=10, sticky="ew")

        self.btn_limpar_console = ctk.CTkButton(
            self.sidebar_frame,
            text="Limpar Console",
            fg_color="gray25",
            hover_color="gray15",
            command=self.ao_clicar_limpar_console,
        )
        self.btn_limpar_console.grid(row=5, column=0, padx=20, pady=(90, 10), sticky="ew")

        self.btn_atualizar_token = ctk.CTkButton(
            self.sidebar_frame,
            text="Atualizar Token Betha",
            fg_color="gray30",
            hover_color="gray20",
            command=self.ao_clicar_atualizar_token,
        )
        self.btn_atualizar_token.grid(row=6, column=0, padx=20, pady=10, sticky="ew")

        self.btn_abrir_configuracoes = ctk.CTkButton(
            self.sidebar_frame,
            text="⚙️ Credenciais de Acesso (Betha e SOC)",
            fg_color="gray30",
            hover_color="gray20",
            command=self.ao_clicar_abrir_configuracoes,
        )
        self.btn_abrir_configuracoes.grid(row=7, column=0, padx=20, pady=(10, 20), sticky="ew")


    def _construir_console(self):
        """Monta o console de log que exibe o andamento das operações."""
        self.console_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.console_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(
            self.console_frame,
            text="Console de Execução",
            font=ctk.CTkFont(weight="bold"),
        ).pack(anchor="w", pady=(0, 5))

        self.log_text = ctk.CTkTextbox(
            self.console_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
        )
        self.log_text.pack(expand=True, fill="both")

    def _registrar_servicos_log(self):
        """Conecta o sistema de log e limpeza aos serviços externos."""
        configurar_log_gui(self.log)
        registrar_limpar_gui(self.ao_clicar_limpar_console)


    def _construir_painel_configuracoes(self):
        """Monta o painel de campos de credenciais (Betha, SOC, Proxy)."""
        self.scroll_frame = ctk.CTkScrollableFrame(
            self, label_text="Credenciais de Acesso"
        )
        self._construir_campos_credenciais_betha()
        self._construir_campos_credenciais_soc()
        self._construir_campos_credenciais_proxy()
        self._construir_botao_salvar_credenciais()

    def _construir_campos_credenciais_betha(self):
        """Campos de login e senha para o sistema Betha."""
        ctk.CTkLabel(
            self.scroll_frame, text="Betha Cloud", font=ctk.CTkFont(weight="bold")
        ).pack(pady=(10, 0))

        self.entry_betha_usuario = ctk.CTkEntry(
            self.scroll_frame, placeholder_text="Login Betha", width=350
        )
        self.entry_betha_usuario.pack(pady=5)

        frame_senha_betha = self._criar_frame_senha(self.scroll_frame)
        self.entry_betha_senha = ctk.CTkEntry(
            frame_senha_betha, placeholder_text="Senha Betha", show="*", width=350
        )
        self.entry_betha_senha.grid(row=0, column=1)
        self.btn_toggle_senha_betha = ctk.CTkButton(
            frame_senha_betha,
            text="Mostrar",
            width=80,
            fg_color=("#dbdbdb", "#2b2b2b"),
            text_color=("#000", "#fff"),
            command=lambda: self.ao_clicar_toggle_senha(
                self.entry_betha_senha, self.btn_toggle_senha_betha
            ),
        )
        self.btn_toggle_senha_betha.grid(row=0, column=2, padx=(5, 0), sticky="w")

    def _construir_campos_credenciais_soc(self):
        """Campos de login, senha e senha virtual para o SOC."""
        ctk.CTkLabel(
            self.scroll_frame, text="SOC Integration", font=ctk.CTkFont(weight="bold")
        ).pack(pady=(20, 0))

        self.entry_soc_usuario = ctk.CTkEntry(
            self.scroll_frame, placeholder_text="E-mail SOC", width=350
        )
        self.entry_soc_usuario.pack(pady=5)

        frame_senha_soc = self._criar_frame_senha(self.scroll_frame)
        self.entry_soc_senha = ctk.CTkEntry(
            frame_senha_soc, placeholder_text="Senha SOC", show="*", width=350
        )
        self.entry_soc_senha.grid(row=0, column=1)
        self.btn_toggle_senha_soc = ctk.CTkButton(
            frame_senha_soc,
            text="Mostrar",
            width=80,
            fg_color=("#dbdbdb", "#2b2b2b"),
            text_color=("#000", "#fff"),
            command=lambda: self.ao_clicar_toggle_senha(
                self.entry_soc_senha, self.btn_toggle_senha_soc
            ),
        )
        self.btn_toggle_senha_soc.grid(row=0, column=2, padx=(5, 0), sticky="w")

        self.entry_soc_senha_virtual = ctk.CTkEntry(
            self.scroll_frame,
            placeholder_text="Senha Virtual (ex: 1,2,3,4)",
            width=350,
        )
        self.entry_soc_senha_virtual.pack(pady=5)

    def _construir_campos_credenciais_proxy(self):
        """Campos de login e senha para o Proxy."""
        ctk.CTkLabel(
            self.scroll_frame, text="Proxy", font=ctk.CTkFont(weight="bold")
        ).pack(pady=(10, 0))

        self.entry_proxy_usuario = ctk.CTkEntry(
            self.scroll_frame, placeholder_text="Login Proxy", width=350
        )
        self.entry_proxy_usuario.pack(pady=5)

        frame_senha_proxy = self._criar_frame_senha(self.scroll_frame)
        self.entry_proxy_senha = ctk.CTkEntry(
            frame_senha_proxy, placeholder_text="Senha Proxy", show="*", width=350
        )
        self.entry_proxy_senha.grid(row=0, column=1)
        self.btn_toggle_senha_proxy = ctk.CTkButton(
            frame_senha_proxy,
            text="Mostrar",
            width=80,
            fg_color=("#dbdbdb", "#2b2b2b"),
            text_color=("#000", "#fff"),
            command=lambda: self.ao_clicar_toggle_senha(
                self.entry_proxy_senha, self.btn_toggle_senha_proxy
            ),
        )
        self.btn_toggle_senha_proxy.grid(row=0, column=2, padx=(5, 0), sticky="w")

    def _construir_botao_salvar_credenciais(self):
        """Botão de salvar credenciais na tela de configurações."""
        self.btn_salvar_credenciais = ctk.CTkButton(
            self.scroll_frame,
            text="Salvar Credenciais",
            fg_color="#28a745",
            hover_color="#218838",
            command=self.ao_clicar_salvar_credenciais,
            width=350,
        )
        self.btn_salvar_credenciais.pack(pady=20)

    def _criar_frame_senha(self, pai):
        """Cria um frame auxiliar centralizado para campo de senha + botão mostrar/esconder."""
        frame = ctk.CTkFrame(pai, fg_color="transparent")
        frame.pack(pady=5, fill="x")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(frame, text="", width=80).grid(row=0, column=0, sticky="e")
        return frame


    def _construir_painel_exportar_periodo(self):
        """Monta o painel completo de exportação por período (novo ou retomar)."""
        self.exportar_frame = ctk.CTkFrame(self, fg_color="transparent")

        ctk.CTkLabel(
            self.exportar_frame,
            text="Exportar Relatório SOC por Período",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w", pady=(0, 4))

        ctk.CTkLabel(
            self.exportar_frame,
            text="Gera o relatório no SOC, lê os dados de cada ficha (CID, CRM…) e salva o Excel na máquina.",
            text_color="gray60",
        ).pack(anchor="w", pady=(0, 20))

        self._construir_seletor_modo_exportacao()
        self._construir_subpainel_novo_relatorio()
        self._construir_subpainel_retomar_relatorio()

        # Exibe o sub-painel "novo" por padrão
        self.frame_novo_relatorio.pack(anchor="w", fill="x")

    def _construir_seletor_modo_exportacao(self):
        """Radio buttons para escolher entre novo relatório ou retomar um existente."""
        frame_escolha = ctk.CTkFrame(self.exportar_frame, fg_color="transparent")
        frame_escolha.pack(anchor="w", pady=(0, 20))

        ctk.CTkLabel(
            frame_escolha, text="O que deseja fazer?", font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self.modo_exportar = ctk.StringVar(value="novo")

        ctk.CTkRadioButton(
            frame_escolha,
            text="📥  Baixar novo relatório do SOC",
            variable=self.modo_exportar,
            value="novo",
            command=self._alternar_subpainel_exportacao,
        ).grid(row=1, column=0, sticky="w", padx=(0, 30))

        ctk.CTkRadioButton(
            frame_escolha,
            text="🔄  Retomar relatório anterior",
            variable=self.modo_exportar,
            value="retomar",
            command=self._alternar_subpainel_exportacao,
        ).grid(row=1, column=1, sticky="w")

    def _construir_subpainel_novo_relatorio(self):
        """Sub-painel com campos de data e botão para gerar novo relatório."""
        self.frame_novo_relatorio = ctk.CTkFrame(self.exportar_frame, fg_color="transparent")

        form = ctk.CTkFrame(self.frame_novo_relatorio, fg_color="transparent")
        form.pack(anchor="w", pady=(0, 15))

        ctk.CTkLabel(
            form, text="Data Inicial (dd/mm/aaaa):", width=200, anchor="w"
        ).grid(row=0, column=0, padx=(0, 10), pady=8, sticky="w")
        self.entry_periodo_data_ini = ctk.CTkEntry(
            form, placeholder_text="Ex: 01/06/2025", width=200
        )
        self.entry_periodo_data_ini.grid(row=0, column=1, pady=8)

        ctk.CTkLabel(
            form, text="Data Final (dd/mm/aaaa):", width=200, anchor="w"
        ).grid(row=1, column=0, padx=(0, 10), pady=8, sticky="w")
        self.entry_periodo_data_fim = ctk.CTkEntry(
            form, placeholder_text="Ex: 30/06/2025", width=200
        )
        self.entry_periodo_data_fim.grid(row=1, column=1, pady=8)

        ctk.CTkButton(
            self.frame_novo_relatorio,
            text="▶  Gerar Relatório",
            fg_color="#5a3e99",
            hover_color="#472e7a",
            width=220,
            command=self.ao_confirmar_exportar_periodo,
        ).pack(anchor="w")

    def _construir_subpainel_retomar_relatorio(self):
        """Sub-painel com dropdown de relatórios pendentes e botão para retomar."""
        self.frame_retomar_relatorio = ctk.CTkFrame(self.exportar_frame, fg_color="transparent")
        self._relatorios_pendentes = []

        ctk.CTkLabel(
            self.frame_retomar_relatorio,
            text="Relatórios com processamento pendente:",
            anchor="w",
        ).pack(anchor="w", pady=(0, 8))

        self.combo_relatorios_pendentes = ctk.CTkComboBox(
            self.frame_retomar_relatorio,
            values=["(nenhum encontrado)"],
            width=480,
            state="readonly",
        )
        self.combo_relatorios_pendentes.pack(anchor="w", pady=(0, 15))

        ctk.CTkButton(
            self.frame_retomar_relatorio,
            text="🔄  Retomar Processamento",
            fg_color="#5a3e99",
            hover_color="#472e7a",
            width=220,
            command=self.ao_confirmar_retomar_relatorio,
        ).pack(anchor="w")


    def _construir_painel_datas_soc(self):
        """Monta o painel de datas + console que aparece ao clicar no botão 2 (SOC)."""
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(1, weight=1)

        self.header_frame_soc = ctk.CTkFrame(self.main_container)

        ctk.CTkLabel(self.header_frame_soc, text="Data Início:").grid(
            row=0, column=0, padx=(15, 5), pady=10
        )
        self.entry_soc_data_ini = ctk.CTkEntry(
            self.header_frame_soc, placeholder_text="DD/MM/AAAA", width=120
        )
        self.entry_soc_data_ini.grid(row=0, column=1, padx=5, pady=10)

        ctk.CTkLabel(self.header_frame_soc, text="Data Fim:").grid(
            row=0, column=2, padx=(15, 5), pady=10
        )
        self.entry_soc_data_fim = ctk.CTkEntry(
            self.header_frame_soc, placeholder_text="DD/MM/AAAA", width=120
        )
        self.entry_soc_data_fim.grid(row=0, column=3, padx=5, pady=10)

        ctk.CTkButton(
            self.header_frame_soc,
            text="Baixar Agora",
            width=100,
            command=self.ao_confirmar_baixar_soc,
        ).grid(row=0, column=4, padx=15, pady=10)

        self.soc_console_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.soc_console_frame.grid_columnconfigure(0, weight=1)
        self.soc_console_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            self.soc_console_frame,
            text="Console de Execução",
            font=ctk.CTkFont(weight="bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.soc_log_text = ctk.CTkTextbox(
            self.soc_console_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
        )
        self.soc_log_text.grid(row=1, column=0, sticky="nsew")


    def _exibir_console(self):
        """Exibe o console e oculta os outros painéis."""
        self.scroll_frame.grid_forget()
        self.exportar_frame.grid_forget()
        self.main_container.grid_forget()
        self.console_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

    def _exibir_painel_configuracoes(self):
        """Exibe o painel de credenciais e oculta os outros."""
        self.console_frame.grid_forget()
        self.exportar_frame.grid_forget()
        self.main_container.grid_forget()
        self.scroll_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

    def _exibir_painel_exportar_periodo(self):
        """Exibe o painel de exportação por período e oculta os outros."""
        self.console_frame.grid_forget()
        self.scroll_frame.grid_forget()
        self.main_container.grid_forget()
        self.exportar_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

    def _exibir_painel_soc(self):
        """Exibe o painel SOC (datas + console) e oculta os outros."""
        self.console_frame.grid_forget()
        self.scroll_frame.grid_forget()
        self.exportar_frame.grid_forget()
        self.main_container.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(1, weight=1)
        self.header_frame_soc.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.soc_console_frame.grid(row=1, column=0, sticky="nsew")

    def _alternar_subpainel_exportacao(self):
        """Alterna entre sub-painel 'novo' e 'retomar' conforme o radio button selecionado."""
        if self.modo_exportar.get() == "novo":
            self.frame_retomar_relatorio.pack_forget()
            self.frame_novo_relatorio.pack(anchor="w", fill="x")
        else:
            self.frame_novo_relatorio.pack_forget()
            pendentes = listar_relatorios_pendentes()
            if pendentes:
                opcoes = [
                    f"{p['nome']}  ({p['pendentes']}/{p['total']} pendentes)"
                    for p in pendentes
                ]
                self._relatorios_pendentes = pendentes
            else:
                opcoes = ["(nenhum relatório pendente encontrado)"]
                self._relatorios_pendentes = []
            self.combo_relatorios_pendentes.configure(values=opcoes)
            self.combo_relatorios_pendentes.set(opcoes[0])
            self.frame_retomar_relatorio.pack(anchor="w", fill="x")


    def ao_clicar_sincronizar_bases(self):
        """Botão 1 → Inicia sincronização das bases Betha no Sheets em thread separada."""
        self._exibir_console()
        threading.Thread(target=self._executar_sincronizar_bases, daemon=True).start()


    def ao_clicar_baixar_soc(self):
        """Botão 2 → Exibe o painel de seleção de período + console para baixar o relatório SOC."""
        self._exibir_painel_soc()
        self.entry_soc_data_ini.delete(0, "end")
        self.entry_soc_data_fim.delete(0, "end")

    def ao_confirmar_baixar_soc(self):
        """[Baixar Agora] → Valida as datas e inicia o download SOC em thread separada."""
        data_ini = self.entry_soc_data_ini.get().strip()
        data_fim = self.entry_soc_data_fim.get().strip()

        if not data_ini or not data_fim:
            messagebox.showwarning("Atenção", "Preencha o período antes de baixar.")
            return

        self.soc_log_text.delete("1.0", "end")
        threading.Thread(
            target=self._executar_baixar_soc,
            args=(data_ini, data_fim),
            daemon=True,
        ).start()


    def ao_clicar_exportar_periodo(self):
        """Botão 3 → Exibe o painel de exportação por período."""
        self._exibir_painel_exportar_periodo()

    def ao_confirmar_exportar_periodo(self):
        """[Gerar Relatório] → Lê datas e inicia exportação em thread separada."""
        self._exibir_console()
        threading.Thread(
            target=self._executar_exportar_periodo,
            args=(
                self.entry_periodo_data_ini.get().strip(),
                self.entry_periodo_data_fim.get().strip(),
            ),
            daemon=True,
        ).start()

    def ao_confirmar_retomar_relatorio(self):
        """[Retomar Processamento] → Encontra o relatório selecionado e retoma em thread."""
        if not self._relatorios_pendentes:
            messagebox.showwarning("Aviso", "Nenhum relatório pendente encontrado.")
            return

        valor_selecionado = self.combo_relatorios_pendentes.get()
        caminho = next(
            (
                p["caminho"]
                for p in self._relatorios_pendentes
                if valor_selecionado.startswith(p["nome"])
            ),
            None,
        )

        if not caminho:
            messagebox.showwarning("Aviso", "Selecione um relatório válido.")
            return

        self._exibir_console()
        threading.Thread(
            target=self._executar_retomar_relatorio,
            args=(caminho,),
            daemon=True,
        ).start()


    def ao_clicar_lancar_atestado(self):
        """Botão 4 → Confirma com o usuário antes de iniciar o lançamento no Betha."""
        if messagebox.askyesno(
            "Revisão", "A planilha Sheets foi conferida e os 'Sim' foram marcados?"
        ):
            self._exibir_console()
            threading.Thread(target=self._executar_lancar_atestado, daemon=True).start()


    def ao_clicar_limpar_console(self):
        """[Limpar Console] → Apaga todo o conteúdo exibido no console."""
        self.log_text.delete("1.0", "end")


    def ao_clicar_atualizar_token(self):
        """[Atualizar Token Betha] → Inicia atualização do token em thread separada."""
        self._exibir_console()
        threading.Thread(target=self._executar_atualizar_token, daemon=True).start()


    def ao_clicar_abrir_configuracoes(self):
        """[⚙️ Credenciais] → Exibe o painel de configurações de acesso."""
        self._exibir_painel_configuracoes()

    def ao_clicar_salvar_credenciais(self):
        """[Salvar Credenciais] → Salva os dados preenchidos e limpa os campos."""
        atualizar_credenciais(
            self.entry_betha_usuario.get(),
            self.entry_betha_senha.get(),
            self.entry_soc_usuario.get(),
            self.entry_soc_senha.get(),
            self.entry_soc_senha_virtual.get(),
            self.entry_proxy_usuario.get(),
            self.entry_proxy_senha.get(),
        )
        messagebox.showinfo("Sucesso", "Configurações salvas localmente!")
        self._limpar_campos_credenciais()
        self._ocultar_senhas()

    def ao_clicar_toggle_senha(self, entry_field, button):
        """[Mostrar/Esconder] → Alterna a visibilidade do campo de senha."""
        if entry_field.cget("show") == "*":
            entry_field.configure(show="")
            button.configure(text="Esconder")
        else:
            entry_field.configure(show="*")
            button.configure(text="Mostrar")


    def _executar_sincronizar_bases(self):
        """Worker: sincroniza as bases Betha com o Google Sheets."""
        self.log("⏳ Iniciando sincronização Betha → Sheets...")
        try:
            res = sincronizar_bases_betha()
            if res.success:
                self.log(f"✅ {res.message}")
            else:
                self.log(f"❌ Erro: {res.message}")
        except Exception as e:
            self.log(f"💥 Erro crítico: {e}")

    def _executar_baixar_soc(self, data_ini, data_fim):
        """Worker: baixa relatório SOC, valida o Excel e importa para o Sheets."""
        self.log_soc(f"🚀 Iniciando busca SOC: {data_ini} até {data_fim}...")
        try:
            res_soc = executar_fluxo_soc(data_ini, data_fim)
            if not (res_soc and res_soc.success):
                self.log_soc(f"❌ Erro no SOC: {res_soc.message if res_soc else 'Falha na extração'}")
                return

            excel = res_soc.data
            self.log_soc(f"🔍 Validando arquivo: {os.path.basename(excel)}")
            output_op = processar_validacoes_excel(excel)

            if output_op.success:
                caminho_validado = output_op.data
                self.log_soc("📤 Importando para Google Sheets...")
                pasta = os.path.dirname(caminho_validado)
                sheets.importar_excel_para_aba(pasta)
                self.log_soc("✅ Processo SOC → Sheets concluído!")
            else:
                self.log_soc(f"❌ Erro na validação: {output_op.message}")


        # try:

        #     excel = r"\\10.1.1.50\ADM_Cresst\Atestados_Laudar\09-04-2026\Relatorio_licensas_medicas_09-04-2026.xlsx"
        #     self.log_soc(f"🔍 Validando arquivo: {os.path.basename(excel)}")
        #     output_op = processar_validacoes_excel(excel)

        #     if output_op.success:
        #         caminho_validado = output_op.data
        #         self.log_soc("📤 Importando para Google Sheets...")
        #         pasta = os.path.dirname(caminho_validado)
        #         sheets.importar_excel_para_aba(pasta)
        #         self.log_soc("✅ Processo SOC → Sheets concluído!")
        #     else:
        #         self.log_soc(f"❌ Erro na validação: {output_op.message}")

        except Exception as e:
            self.log_soc(f"💥 Erro no SOC: {e}")

    def _executar_exportar_periodo(self, data_ini, data_fim):
        """Worker: gera e salva relatório SOC para o período informado."""
        self.log(f"📅 Exportando relatório SOC: {data_ini} → {data_fim}")
        try:
            res = exportar_relatorio_por_periodo(data_ini, data_fim)
            if res.success:
                self.log(f"✅ Relatório salvo em: {res.data}")
            else:
                self.log(f"❌ {res.message}")
        except Exception as e:
            self.log(f"💥 Erro crítico: {e}")

    def _executar_retomar_relatorio(self, caminho):
        """Worker: retoma o processamento de um relatório parcialmente concluído."""
        self.log(f"🔄 Retomando: {os.path.basename(caminho)}")
        try:
            res = retomar_relatorio_por_periodo(caminho)
            if res.success:
                self.log(f"✅ Concluído: {res.data}")
            else:
                self.log(f"❌ {res.message}")
        except Exception as e:
            self.log(f"💥 Erro crítico: {e}")

    def _executar_atualizar_token(self):
        """Worker: solicita novo token de acesso ao Betha."""
        try:
            res_token = atualizar_token_betha()
            if hasattr(res_token, "success") and res_token.success:
                self.log("✅ Token Betha atualizado com sucesso!")
            else:
                msg = res_token.message if hasattr(res_token, "message") else str(res_token)
                self.log(f"❌ Falha no token Betha: {msg}")
        except Exception as e:
            self.log(f"💥 Erro crítico na interface: {e}")

    def _executar_lancar_atestado(self):
        """Worker: lê a planilha Sheets e lança cada atestado no sistema Betha."""
        dialogo = ctk.CTkInputDialog(
            text="A partir de qual linha ler a planilha?", title="Ponto de Início"
        )
        entrada = dialogo.get_input()

        if entrada is None:
            self.log("ℹ️ Operação cancelada pelo usuário.")
            return

        try:
            linha_inicio = int(entrada) if entrada else 2
        except ValueError:
            self.log("❌ Valor inválido. Digite apenas números.")
            return

        self.log(f"⚙️ Iniciando processamento (Linha inicial: {linha_inicio})...")

        try:
            betha = BethaService()
            if not betha.inicializado:
                self.log(f"❌ Erro de Configuração: {betha.erro_inicializacao}")
                return

            resultado_bruto = sheets.ler_planilha_para_automacao(linha_inicio)
            if not resultado_bruto:
                self.log("ℹ️ Nenhum dado pendente encontrado na planilha.")
                return

            res_lote = betha.processar_lote_planilha(resultado_bruto)
            if not res_lote.success:
                self.log(f"⚠️ Erro no processamento de lote: {res_lote.message}")
                return

            res_payloads = betha.gerar_payloads_lote(res_lote.data)
            if not res_payloads.success:
                self.log(f"⚠️ Erro ao gerar payloads: {res_payloads.message}")
                return

            lista_payloads = res_payloads.data
            total = len(lista_payloads)
            self.log(f"📦 {total} atestados prontos para envio.")
            sucessos, falhas = 0, 0

            for payload in lista_payloads:
                try:
                    time.sleep(2)
                    cod_ficha = payload.get("numeroAtestado")
                    payload_envio = payload.copy()
                    payload_envio["numeroAtestado"] = None

                    self.log(f"📤 Enviando Ficha: {cod_ficha}...")
                    result = betha.enviar_atestado(payload_envio)

                    if result.success:
                        sheets.marcar_status_na_planilha(cod_ficha, mensagem_status="ENVIADO")
                        self.log(f"✅ Ficha {cod_ficha}: Sucesso!")
                        sucessos += 1
                    else:
                        sheets.marcar_status_na_planilha(
                            cod_ficha, mensagem_status=f"ERRO | {result.message or result}"
                        )
                        self.log(f"❌ Ficha {cod_ficha}: Falhou ({result.message or result})")
                        falhas += 1

                except Exception as e_loop:
                    self.log(f"⚠️ Erro inesperado na ficha {payload.get('numeroAtestado')}: {e_loop}")
                    continue

            self.log(f"🏁 Fim do processo. Sucessos: {sucessos} | Falhas: {falhas}")

        except Exception as e:
            self.log(f"💥 Erro Crítico no sistema: {str(e)}")


    def log(self, msg):
        """Insere uma mensagem no console principal com timestamp."""
        self.log_text.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see("end")

    def log_soc(self, msg):
        """Insere uma mensagem no console do painel SOC com timestamp."""
        self.soc_log_text.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.soc_log_text.see("end")

    def _limpar_campos_credenciais(self):
        """Apaga o conteúdo de todos os campos de credenciais após salvar."""
        for entry in (
            self.entry_betha_usuario,
            self.entry_betha_senha,
            self.entry_soc_usuario,
            self.entry_soc_senha,
            self.entry_soc_senha_virtual,
            self.entry_proxy_usuario,
            self.entry_proxy_senha,
        ):
            entry.delete(0, "end")

    def _ocultar_senhas(self):
        """Restaura todos os campos de senha para o modo oculto (*) e reseta os botões."""
        for entry in (self.entry_betha_senha, self.entry_soc_senha, self.entry_proxy_senha):
            entry.configure(show="*")
        for btn in (self.btn_toggle_senha_betha, self.btn_toggle_senha_soc, self.btn_toggle_senha_proxy):
            btn.configure(text="Mostrar")


if __name__ == "__main__":
    app = App()
    app.mainloop()