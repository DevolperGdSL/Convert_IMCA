#!/usr/bin/env python3
"""
=============================================================================
CONVERT_IMCA - Interface Gráfica Desktop (v1.1.0)
Interface moderna de tela única para conversão de planilhas Excel (.xlsx)
para o formato ultra-leve .IMCA (cabeceira-pwa1).
Compatível com Linux e Windows.
=============================================================================
"""

import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# Importa o motor de conversão existente
try:
    from convert_imca import parse_xlsx, build_imca_content, __version__
except ImportError:
    # Se rodar como executável congelado (PyInstaller)
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from convert_imca import parse_xlsx, build_imca_content, __version__


class ConvertImcaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Convert_IMCA - Conversor de Planilhas para .IMCA")
        self.geometry("780x680")
        self.minsize(700, 600)
        self.configure(bg="#0f172a")  # Slate 900 escuro moderno

        # Tentar carregar ícone
        icon_path_png = os.path.join(os.path.dirname(__file__), "assets", "icon.png")
        icon_path_ico = os.path.join(os.path.dirname(__file__), "assets", "icon.ico")
        try:
            if sys.platform.startswith("win") and os.path.exists(icon_path_ico):
                self.iconbitmap(icon_path_ico)
            elif os.path.exists(icon_path_png):
                img = tk.PhotoImage(file=icon_path_png)
                self.iconphoto(True, img)
        except Exception:
            pass

        self.input_file_var = tk.StringVar()
        self.default_time_var = tk.StringVar(value="19:30")
        self.translate_days_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Selecione uma planilha (.xlsx) para começar.")
        self.output_file_path = None

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Configurações globais de cores
        style.configure(".", background="#0f172a", foreground="#f8fafc", font=("Inter", 10))

        # Cards / Frames
        style.configure("Card.TFrame", background="#1e293b", relief="flat")
        style.configure("InnerCard.TFrame", background="#334155", relief="flat")

        # Labels
        style.configure("Title.TLabel", background="#0f172a", foreground="#ffffff", font=("Inter", 16, "bold"))
        style.configure("Subtitle.TLabel", background="#0f172a", foreground="#94a3b8", font=("Inter", 10))
        style.configure("CardTitle.TLabel", background="#1e293b", foreground="#e2e8f0", font=("Inter", 11, "bold"))
        style.configure("CardDesc.TLabel", background="#1e293b", foreground="#94a3b8", font=("Inter", 9))
        style.configure("StatValue.TLabel", background="#334155", foreground="#38bdf8", font=("Inter", 13, "bold"))
        style.configure("StatLabel.TLabel", background="#334155", foreground="#cbd5e1", font=("Inter", 8))

        # Inputs
        style.configure("TEntry", fieldbackground="#0f172a", foreground="#ffffff", bordercolor="#475569")
        style.configure("TCheckbutton", background="#1e293b", foreground="#f1f5f9", font=("Inter", 10))

        # Botões
        style.configure("Primary.TButton", font=("Inter", 11, "bold"), background="#10b981", foreground="#ffffff")
        style.map("Primary.TButton",
                  background=[("active", "#059669"), ("pressed", "#047857")],
                  foreground=[("active", "#ffffff")])

        style.configure("Secondary.TButton", font=("Inter", 9, "bold"), background="#4f46e5", foreground="#ffffff")
        style.map("Secondary.TButton",
                  background=[("active", "#4338ca"), ("pressed", "#3730a3")],
                  foreground=[("active", "#ffffff")])

        style.configure("Outline.TButton", font=("Inter", 9), background="#334155", foreground="#f8fafc")
        style.map("Outline.TButton",
                  background=[("active", "#475569")],
                  foreground=[("active", "#ffffff")])

    def _build_ui(self):
        container = ttk.Frame(self, padding="20 15 20 20")
        container.pack(fill=tk.BOTH, expand=True)

        # 1. HEADER
        header_frame = ttk.Frame(container)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        title_label = ttk.Label(header_frame, text="⚡ Convert_IMCA", style="Title.TLabel")
        title_label.pack(anchor="w")

        sub_label = ttk.Label(
            header_frame,
            text="Conversor de planilhas Excel (.xlsx) para o formato ultra-leve .IMCA (cabeceira-pwa1)",
            style="Subtitle.TLabel"
        )
        sub_label.pack(anchor="w", pady=(2, 0))

        # 2. CARD: SELEÇÃO DE ARQUIVO
        file_card = ttk.Frame(container, style="Card.TFrame", padding="15")
        file_card.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(file_card, text="1. Planilha Excel de Entrada", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(file_card, text="Selecione o arquivo .xlsx contendo o calendário ou eventos", style="CardDesc.TLabel").pack(anchor="w", pady=(0, 10))

        file_row = ttk.Frame(file_card, style="Card.TFrame")
        file_row.pack(fill=tk.X)

        entry = ttk.Entry(file_row, textvariable=self.input_file_var, font=("Inter", 10))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4, padx=(0, 10))

        btn_browse = ttk.Button(file_row, text="📂 Procurar...", style="Secondary.TButton", command=self._browse_file)
        btn_browse.pack(side=tk.RIGHT, ipadx=10, ipady=3)

        # 3. CARD: CONFIGURAÇÕES RÁPIDAS
        opt_card = ttk.Frame(container, style="Card.TFrame", padding="15")
        opt_card.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(opt_card, text="2. Opções de Formatação", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))

        opt_row = ttk.Frame(opt_card, style="Card.TFrame")
        opt_row.pack(fill=tk.X)

        ttk.Label(opt_row, text="Horário Padrão:", style="CardDesc.TLabel").pack(side=tk.LEFT, padx=(0, 8))
        time_entry = ttk.Entry(opt_row, textvariable=self.default_time_var, width=8, font=("Inter", 10))
        time_entry.pack(side=tk.LEFT, padx=(0, 20))

        chk_translate = ttk.Checkbutton(
            opt_row,
            text="Traduzir dias da semana para Português (Segunda, Terça...)",
            variable=self.translate_days_var,
            style="TCheckbutton"
        )
        chk_translate.pack(side=tk.LEFT)

        # 4. BOTÃO PRINCIPAL DE CONVERSÃO
        self.btn_convert = ttk.Button(
            container,
            text="✨ CONVERTER PARA FORMATO .IMCA",
            style="Primary.TButton",
            command=self._convert
        )
        self.btn_convert.pack(fill=tk.X, ipady=8, pady=(0, 15))

        # 5. CARD: ARQUIVO TRANSFORMADO & RESULTADO
        self.result_card = ttk.Frame(container, style="Card.TFrame", padding="15")
        self.result_card.pack(fill=tk.BOTH, expand=True)

        res_header = ttk.Frame(self.result_card, style="Card.TFrame")
        res_header.pack(fill=tk.X, pady=(0, 8))

        self.lbl_result_status = ttk.Label(res_header, text="3. Arquivo Transformado & Estatísticas", style="CardTitle.TLabel")
        self.lbl_result_status.pack(side=tk.LEFT)

        self.btn_open_folder = ttk.Button(
            res_header,
            text="📁 Abrir na Pasta",
            style="Outline.TButton",
            command=self._open_folder,
            state=tk.DISABLED
        )
        self.btn_open_folder.pack(side=tk.RIGHT)

        # Caminho do Arquivo Gerado
        self.lbl_out_path = ttk.Label(
            self.result_card,
            text="Nenhum arquivo convertido ainda.",
            style="CardDesc.TLabel"
        )
        self.lbl_out_path.pack(anchor="w", pady=(0, 10))

        # Métricas (Cards em Grid)
        self.metrics_frame = ttk.Frame(self.result_card, style="Card.TFrame")
        self.metrics_frame.pack(fill=tk.X, pady=(0, 12))

        self.stat_events = self._create_stat_box(self.metrics_frame, "EVENTOS", "0")
        self.stat_orig = self._create_stat_box(self.metrics_frame, "TAMANHO .XLSX", "0 KB")
        self.stat_imca = self._create_stat_box(self.metrics_frame, "TAMANHO .IMCA", "0 KB")
        self.stat_savings = self._create_stat_box(self.metrics_frame, "ECONOMIA", "0%")

        # Prévia em Treeview
        ttk.Label(self.result_card, text="Prévia dos Registros Gerados:", style="CardDesc.TLabel").pack(anchor="w", pady=(0, 4))

        tree_frame = ttk.Frame(self.result_card, style="Card.TFrame")
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("data", "dia", "horario", "nome", "local", "solicitante")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=5)
        
        self.tree.heading("data", text="Data (ISO)")
        self.tree.heading("dia", text="Dia da Semana")
        self.tree.heading("horario", text="Horário")
        self.tree.heading("nome", text="Nome do Evento")
        self.tree.heading("local", text="Local")
        self.tree.heading("solicitante", text="Solicitante")

        self.tree.column("data", width=95, anchor="center")
        self.tree.column("dia", width=110, anchor="center")
        self.tree.column("horario", width=65, anchor="center")
        self.tree.column("nome", width=230, anchor="w")
        self.tree.column("local", width=90, anchor="w")
        self.tree.column("solicitante", width=140, anchor="w")

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_stat_box(self, parent, label_text, initial_value):
        box = ttk.Frame(parent, style="InnerCard.TFrame", padding="8 6")
        box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)

        val_lbl = ttk.Label(box, text=initial_value, style="StatValue.TLabel", anchor="center")
        val_lbl.pack(fill=tk.X)

        sub_lbl = ttk.Label(box, text=label_text, style="StatLabel.TLabel", anchor="center")
        sub_lbl.pack(fill=tk.X)

        return val_lbl

    def _browse_file(self):
        file_selected = filedialog.askopenfilename(
            title="Selecione a Planilha Excel (.xlsx)",
            filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
        )
        if file_selected:
            self.input_file_var.set(file_selected)

    def _convert(self):
        input_path = self.input_file_var.get().strip()
        if not input_path:
            messagebox.showwarning("Aviso", "Por favor, selecione uma planilha Excel (.xlsx) antes de converter.")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Erro", f"O arquivo '{input_path}' não foi encontrado.")
            return

        start_time = datetime.now()
        out_path = os.path.splitext(input_path)[0] + ".imca"
        default_time = self.default_time_var.get().strip() or "19:30"
        translate_days = self.translate_days_var.get()

        try:
            metadata, events = parse_xlsx(
                file_path=input_path,
                default_time=default_time,
                translate_days=translate_days
            )
        except Exception as e:
            messagebox.showerror("Falha na Conversão", f"Ocorreu um erro ao processar a planilha:\n\n{e}")
            return

        if not events:
            messagebox.showwarning("Atenção", "Nenhum evento válido foi encontrado na planilha informada.")
            return

        imca_content = build_imca_content(metadata, events)

        try:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(imca_content)
        except Exception as e:
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo .imca:\n\n{e}")
            return

        orig_size = os.path.getsize(input_path)
        imca_size = len(imca_content.encode("utf-8"))
        reduction = ((orig_size - imca_size) / orig_size) * 100 if orig_size else 0
        elapsed_ms = (datetime.now() - start_time).total_seconds() * 1000

        self.output_file_path = out_path
        self.lbl_result_status.config(text="✅ Arquivo Transformado com Sucesso!")
        self.lbl_out_path.config(text=f"Destino: {out_path} ({elapsed_ms:.1f} ms)")
        self.btn_open_folder.config(state=tk.NORMAL)

        # Atualizar Cards
        self.stat_events.config(text=str(len(events)))
        self.stat_orig.config(text=f"{orig_size / 1024:.1f} KB")
        self.stat_imca.config(text=f"{imca_size / 1024:.1f} KB")
        self.stat_savings.config(text=f"-{reduction:.1f}%")

        # Atualizar Prévia
        for item in self.tree.get_children():
            self.tree.delete(item)

        for ev in events[:20]:
            self.tree.insert("", tk.END, values=(
                ev["data"],
                ev["dia_semana"],
                ev["horario"],
                ev["nome"],
                ev["local"] or "-",
                ev["solicitante"] or "-"
            ))

    def _open_folder(self):
        if not self.output_file_path or not os.path.exists(self.output_file_path):
            return

        folder = os.path.dirname(os.path.abspath(self.output_file_path))
        if sys.platform.startswith("win"):
            os.system(f'explorer /select,"{os.path.abspath(self.output_file_path)}"')
        elif sys.platform.startswith("darwin"):
            subprocess.run(["open", "-R", self.output_file_path])
        else:
            subprocess.run(["xdg-open", folder])


def main():
    app = ConvertImcaApp()
    app.mainloop()


if __name__ == "__main__":
    main()
