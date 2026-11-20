import tkinter as tk

from tkinter import ttk
from tkinter import messagebox

from pathlib import Path

import os
import sys

from FolderService import FolderService
from DocxToExcelConverter import DocxToExcelConverter
from MergeService import MergeService


class MainWindow:

    def __init__(self):

        self.root = tk.Tk()

        self.root.title("Epimed Cultura")

        self.root.geometry("900x650")

        self.root.minsize(
            800,
            550
        )

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.exit_application
        )

        self.base_path = (
            Path(sys.executable).resolve().parent
            if getattr(sys, "frozen", False)
            else Path.cwd()
        )

        self.folder_service = (
            FolderService(
                self.base_path
            )
        )

        self.converter = (
            DocxToExcelConverter()
        )

        self.merge_service = (
            MergeService()
        )

        self._center_window()

        self._create_components()

    # ==================================
    # INTERFACE
    # ==================================

    def _center_window(self):

        self.root.update_idletasks()

        width = 900
        height = 650

        screen_width = (
            self.root.winfo_screenwidth()
        )

        screen_height = (
            self.root.winfo_screenheight()
        )

        x = int(
            (screen_width / 2)
            - (width / 2)
        )

        y = int(
            (screen_height / 2)
            - (height / 2)
        )

        self.root.geometry(
            f"{width}x{height}+{x}+{y}"
        )

    def _create_components(self):

        container = ttk.Frame(
            self.root,
            padding=20
        )

        container.pack(
            fill="both",
            expand=True
        )

        titulo = ttk.Label(
            container,
            text="EPIMED CULTURA",
            font=(
                "Segoe UI",
                18,
                "bold"
            )
        )

        titulo.pack(
            pady=(0, 15)
        )

        subtitulo = ttk.Label(
            container,
            text=(
                "Sistema de cruzamento "
                "entre Cultura e Epimed"
            )
        )

        subtitulo.pack(
            pady=(0, 20)
        )

        # -------------------------
        # Botões
        # -------------------------

        self.btn_create = ttk.Button(
            container,
            text="Criar Pastas",
            command=self.create_folders
        )

        self.btn_create.pack(
            fill="x",
            pady=3
        )

        self.btn_convert = ttk.Button(
            container,
            text="Realizar Conversão",
            command=self.convert_files
        )

        self.btn_convert.pack(
            fill="x",
            pady=3
        )

        self.btn_generate = ttk.Button(
            container,
            text="Gerar Base Final",
            command=self.generate_base
        )

        self.btn_generate.pack(
            fill="x",
            pady=3
        )

        self.btn_open = ttk.Button(
            container,
            text="Abrir Pasta do Projeto",
            command=self.open_project_folder
        )

        self.btn_open.pack(
            fill="x",
            pady=3
        )

        self.btn_exit = ttk.Button(
            container,
            text="Sair",
            command=self.exit_application
        )

        self.btn_exit.pack(
            fill="x",
            pady=(12, 3)
        )

        # -------------------------
        # Barra de progresso
        # -------------------------

        ttk.Label(
            container,
            text="Status"
        ).pack(
            pady=(20, 5)
        )

        self.progress = ttk.Progressbar(
            container,
            mode="indeterminate"
        )

        self.progress.pack(
            fill="x"
        )

        # -------------------------
        # Área de log
        # -------------------------

        ttk.Label(
            container,
            text="Log de Execução"
        ).pack(
            pady=(20, 5)
        )

        self.log_text = tk.Text(
            container,
            height=15
        )

        self.log_text.pack(
            fill="both",
            expand=True
        )

    # ==================================
    # LOG
    # ==================================

    def log(self, message):

        self.log_text.insert(
            tk.END,
            f"{message}\n"
        )

        self.log_text.see(
            tk.END
        )

        self.root.update()

    # ==================================
    # PASTAS
    # ==================================

    def create_folders(self):

        try:

            self.folder_service.create_folders()

            self.log(
                "Pastas limpas e criadas com sucesso."
            )

            os.startfile(
                self.base_path
            )

            messagebox.showinfo(
                "Pastas Criadas",
                (
                    "1. Coloque os DOCX em Dados_Epimed\n\n"
                    "2. Coloque o XLSX de Cultura em Culturas\n\n"
                    "3. Clique em Realizar Conversão"
                )
            )

        except Exception as ex:

            messagebox.showerror(
                "Erro",
                str(ex)
            )

    # ==================================
    # CONVERSÃO
    # ==================================

    def convert_files(self):

        try:

            self.progress.start()

            self.log(
                "Iniciando conversão..."
            )

            total = (
                self.converter.convert(
                    self.folder_service.dados_epimed,
                    self.folder_service.xlsx
                )
            )

            self.progress.stop()

            self.log(
                f"{total} arquivo(s) convertidos."
            )

            messagebox.showinfo(
                "Conversão Finalizada",
                (
                    f"{total} arquivo(s) "
                    f"convertidos."
                )
            )

        except Exception as ex:

            self.progress.stop()

            messagebox.showerror(
                "Erro",
                str(ex)
            )

    # ==================================
    # MERGE
    # ==================================

    def generate_base(self):

        try:

            self.progress.start()

            self.log(
                "Iniciando merge..."
            )

            output_file = (
                self.merge_service.execute(
                    cultura_folder=
                        self.folder_service.culturas,

                    epimed_xlsx_folder=
                        self.folder_service.xlsx,

                    output_folder=
                        self.folder_service.saida
                )
            )

            self.progress.stop()

            self.log(
                "Base final criada."
            )

            self.log(
                f"Arquivo: {output_file}"
            )

            messagebox.showinfo(
                "Sucesso",
                (
                    "Base final criada com sucesso.\n\n"
                    f"{output_file}"
                )
            )

        except Exception as ex:

            self.progress.stop()

            messagebox.showerror(
                "Erro",
                str(ex)
            )

    # ==================================
    # ABRIR PASTA
    # ==================================

    def open_project_folder(self):

        os.startfile(
            self.base_path
        )

    def exit_application(self):

        resposta = messagebox.askyesno(
            "Sair",
            "Deseja encerrar o sistema?"
        )

        if resposta:
            self.root.destroy()

    # ==================================
    # EXECUÇÃO
    # ==================================

    def run(self):

        self.root.mainloop()