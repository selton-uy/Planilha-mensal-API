import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

import limpa_planilha
import monta_tabela

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

PLANILHA_BASE = 'planilha_base.xlsx'


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Acordões")
        self.geometry("520x560")
        self.resizable(False, False)

        # ── Título ──────────────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text="Gerador de Planilha de Acordões",
            font=("Arial", 18, "bold"),
        ).pack(pady=(24, 12))

        frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        frame_botoes.pack(anchor="center")

        self.btn_limpar = ctk.CTkButton(
            frame_botoes,
            text="Limpar Planilha",
            width=200,
            command=self.acao_limpar,
        )
        self.btn_limpar.pack(padx=20,pady=10)

        self.btn_gerar = ctk.CTkButton(
            frame_botoes,
            text="Gerar Planilha Nova",
            width=200,
            command=self.acao_gerar,
        )
        self.btn_gerar.pack(padx=10,pady=10)

        # ── Status principal ─────────────────────────────────────────────
        self.status_label = ctk.CTkLabel(self, text="", font=("Arial", 13))
        self.status_label.pack(pady=(14, 0))

        # ── Status de falhas (amarelo, abaixo do status principal) ───────
        self.falhas_label = ctk.CTkLabel(self, text="", font=("Arial", 12), text_color="#FFD700")
        self.falhas_label.pack(pady=(2, 2))

        # ── Barra de progresso ───────────────────────────────────────────
        self.progress = ctk.CTkProgressBar(self, width=460, mode="indeterminate")

        # ── Área de logs ─────────────────────────────────────────────────
        ctk.CTkLabel(self, text="Logs", font=("Arial", 12, "bold")).pack(anchor="w", padx=28, pady=(12, 2))

        self.log_box = ctk.CTkTextbox(
            self,
            width=464,
            height=200,
            font=("Courier", 12),
            state="disabled",
            wrap="word",
        )
        self.log_box.pack(padx=28, pady=(0, 8))

        btn_limpar_log = ctk.CTkButton(
            self,
            text="Limpar logs",
            width=120,
            height=24,
            fg_color="transparent",
            border_width=1,
            command=self._limpar_logs,
        )
        btn_limpar_log.pack(anchor="e", padx=28)

    # ── Log ──────────────────────────────────────────────────────────────

    def _log(self, mensagem: str):
        def _inserir(*args):
            self.log_box.configure(state="normal")
            self.log_box.insert("end", mensagem + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _inserir)

    def _limpar_logs(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ── Helpers de interface ─────────────────────────────────────────────

    def _iniciar_ui(self, mensagem: str):
        self.btn_limpar.configure(state="disabled")
        self.btn_gerar.configure(state="disabled")
        self.status_label.configure(text=mensagem, text_color="white")
        self.falhas_label.configure(text="")
        self.progress.pack(pady=2)
        self.progress.start()

    def _finalizar_ui(self, mensagem: str, erro: bool = False, falhas: int = 0):
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_limpar.configure(state="normal")
        self.btn_gerar.configure(state="normal")
        cor = "#e05555" if erro else "#3fa06e"
        self.status_label.configure(text=mensagem, text_color=cor)
        if falhas > 0:
            self.falhas_label.configure(text=f"⚠️ {falhas} processo(s) ficaram em branco — marcados em vermelho na planilha.")
        else:
            self.falhas_label.configure(text="")

    def _executar_em_thread(self, funcao, msg_rodando, msg_ok, msg_erro):
        self.after(0, self._iniciar_ui, msg_rodando)

        def tarefa():
            try:
                resultado = funcao()
                falhas = resultado if isinstance(resultado, int) else 0
                self.after(0, self._finalizar_ui, msg_ok, False, falhas)
            except Exception:
                self.after(0, self._finalizar_ui, msg_erro, True, 0)


        threading.Thread(target=tarefa, daemon=True).start()

    # ── Ações dos botões ─────────────────────────────────────────────────

    def acao_limpar(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha para limpar",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
        )
        if not caminho:
            return

        self._executar_em_thread(
            funcao=lambda: limpa_planilha.main(caminho),
            msg_rodando="⏳ Limpando planilha...",
            msg_ok="✅ planilha_base.xlsx salva!",
            msg_erro="❌ Erro ao limpar planilha.",
        )

    def acao_gerar(self):
        # 1. Escolher arquivo de entrada
        if os.path.exists(PLANILHA_BASE):
            usar_base = messagebox.askyesno(
                "Arquivo de entrada",
                f"Foi encontrada uma '{PLANILHA_BASE}'.\n\nDeseja usá-la como entrada?\n\n"
                "Clique 'Não' para escolher outro arquivo.",
            )
            caminho_entrada = PLANILHA_BASE if usar_base else None
        else:
            caminho_entrada = None

        if caminho_entrada is None:
            caminho_entrada = filedialog.askopenfilename(
                title="Selecione a planilha de entrada",
                filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
            )
            if not caminho_entrada:
                return

        # 2. Escolher onde salvar
        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar planilha nova como...",
            defaultextension=".xlsx",
            initialfile="planilha_nova",
            filetypes=[("Arquivo Excel", "*.xlsx")],
        )
        if not caminho_saida:
            return

        self._executar_em_thread(
            funcao=lambda: monta_tabela.main(caminho_entrada, caminho_saida, self._log),
            msg_rodando="⏳ Consultando processos... (pode demorar)",
            msg_ok="✅ Planilha gerada com sucesso!",
            msg_erro="❌ Erro ao gerar planilha.",
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()