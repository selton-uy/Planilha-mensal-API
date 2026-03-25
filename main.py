import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

import limpa_planilha
import monta_tabela
import baixa_acordaos

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

PLANILHA_BASE = 'planilha_base.xlsx'


class DialogColuna(ctk.CTkToplevel):
    """Janela modal para o usuário digitar/escolher o nome da coluna."""

    def __init__(self, parent, colunas: list[str]):
        super().__init__(parent)
        self.title("Coluna de códigos")
        self.geometry("380x220")
        self.resizable(False, False)
        self.grab_set()  # modal
        self.coluna_escolhida = None

        ctk.CTkLabel(self, text="Selecione a coluna com os códigos de processo:",
                     wraplength=340, font=("Arial", 13)).pack(pady=(20, 10), padx=20)

        self.combo = ctk.CTkComboBox(self, values=colunas, width=340)
        self.combo.set(colunas[0] if colunas else "")
        self.combo.pack(padx=20, pady=4)

        ctk.CTkButton(self, text="Confirmar", width=160,
                      command=self._confirmar).pack(pady=(16, 0))

    def _confirmar(self):
        self.coluna_escolhida = self.combo.get()
        self.destroy()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Acordões")
        self.geometry("520x660")
        self.resizable(False, False)

        # ── Título ──────────────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text="Gerador de Planilha de Acordões",
            font=("Arial", 18, "bold"),
        ).pack(pady=(24, 12))

        # ── Botões linha 1 ───────────────────────────────────────────────
        frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        frame_botoes.pack(anchor="center")

        self.btn_limpar = ctk.CTkButton(
            frame_botoes,
            text="1 · Limpar Planilha",
            width=200,
            command=self.acao_limpar,
        )
        self.btn_limpar.pack(side="left", padx=10)

        self.btn_gerar = ctk.CTkButton(
            frame_botoes,
            text="2 · Gerar Planilha Nova",
            width=200,
            command=self.acao_gerar,
        )
        self.btn_gerar.pack(side="left", padx=10)

        # ── Botão linha 2 ────────────────────────────────────────────────
        frame_botoes2 = ctk.CTkFrame(self, fg_color="transparent")
        frame_botoes2.pack(anchor="center", pady=(8, 0))

        self.btn_baixar = ctk.CTkButton(
            frame_botoes2,
            text="3 · Baixar Acórdãos",
            width=200,
            command=self.acao_baixar,
        )
        self.btn_baixar.pack()

        # ── Status principal ─────────────────────────────────────────────
        self.status_label = ctk.CTkLabel(self, text="", font=("Arial", 13))
        self.status_label.pack(pady=(14, 0))

        # ── Status de falhas (amarelo) ───────────────────────────────────
        self.falhas_label = ctk.CTkLabel(self, text="", font=("Arial", 12), text_color="#FFD700")
        self.falhas_label.pack(pady=(2, 2))

        # ── Barra de progresso ───────────────────────────────────────────
        self.progress = ctk.CTkProgressBar(self, width=460, mode="indeterminate")

        # ── Área de logs ─────────────────────────────────────────────────
        ctk.CTkLabel(self, text="Logs", font=("Arial", 12, "bold")).pack(anchor="w", padx=28, pady=(12, 2))

        self.log_box = ctk.CTkTextbox(
            self,
            width=464,
            height=220,
            font=("Courier", 12),
            state="disabled",
            wrap="word",
        )
        self.log_box.pack(padx=28, pady=(0, 8))

        ctk.CTkButton(
            self,
            text="Limpar logs",
            width=120,
            height=24,
            fg_color="transparent",
            border_width=1,
            command=self._limpar_logs,
        ).pack(anchor="e", padx=28)

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

    def _todos_botoes(self, state: str):
        self.btn_limpar.configure(state=state)
        self.btn_gerar.configure(state=state)
        self.btn_baixar.configure(state=state)

    def _iniciar_ui(self, mensagem: str):
        self._todos_botoes("disabled")
        self.status_label.configure(text=mensagem, text_color="white")
        self.falhas_label.configure(text="")
        self.progress.pack(pady=2)
        self.progress.start()

    def _finalizar_ui(self, mensagem: str, erro: bool = False, falhas: int = 0):
        self.progress.stop()
        self.progress.pack_forget()
        self._todos_botoes("normal")
        cor = "#e05555" if erro else "#3fa06e"
        self.status_label.configure(text=mensagem, text_color=cor)
        if falhas > 0:
            self.falhas_label.configure(
                text=f"⚠️ {falhas} processo(s) ficaram em branco — marcados em vermelho na planilha."
            )
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

    def acao_baixar(self):
        # 1. Escolher planilha
        caminho_planilha = filedialog.askopenfilename(
            title="Selecione a planilha com os códigos",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
        )
        if not caminho_planilha:
            return

        # 2. Ler colunas da planilha e perguntar qual usar
        try:
            import pandas as pd
            colunas = list(pd.read_excel(caminho_planilha, nrows=0).columns)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler a planilha:\n{e}")
            return

        dialog = DialogColuna(self, colunas)
        self.wait_window(dialog)
        coluna = dialog.coluna_escolhida
        if not coluna:
            return

        # 3. Escolher pasta de destino
        pasta_destino = filedialog.askdirectory(title="Selecione a pasta para salvar os PDFs")
        if not pasta_destino:
            return

        self._executar_em_thread(
            funcao=lambda: baixa_acordaos.main(caminho_planilha, coluna, pasta_destino, self._log),
            msg_rodando="⏳ Baixando acórdãos... (pode demorar)",
            msg_ok="✅ Downloads concluídos!",
            msg_erro="❌ Erro ao baixar acórdãos.",
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()