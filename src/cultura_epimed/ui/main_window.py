import tkinter as tk

from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog

from pathlib import Path

import sys
import time

from cultura_epimed.services.docx_to_excel_converter import (
    DocxToExcelConverter
)
from cultura_epimed.services.folder_service import FolderService
from cultura_epimed.services.merge_service import MergeService


FILE_TYPE_LABELS = {
    "docx": "Word",
    "xls": "Excel (xls)",
    "pdf": "PDF",
    "csv": "CSV",
    "xlsx": "Excel"
}


class MainWindow:

    def __init__(self):

        self.root = tk.Tk()
        self.app_icon = self._find_app_icon()
        self.root.iconbitmap(str(self.app_icon))
        self.splash_started_at = time.monotonic()
        self._create_splash()

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

        self.folder_service.create_folders()

        self.converter = (
            DocxToExcelConverter()
        )

        self.merge_service = (
            MergeService()
        )

        self.selected_conversion_type = None
        self.selected_source_folder = None
        self.selected_files = []
        self.selected_file_count = 0
        self.cultura_xlsx_folder = None
        self.epimed_xlsx_folder = None
        self.progress_value = 0
        self.progress_description = "Processando"

    # ==================================
    # INTERFACE
    # ==================================

    def _find_app_icon(self) -> Path:
        filename = "cultura-epimed.ico"
        search_roots = [
            Path(__file__).resolve().parents[1]
        ]
        bundle_path = getattr(sys, "_MEIPASS", None)
        if bundle_path:
            search_roots.append(Path(bundle_path))
        if getattr(sys, "frozen", False):
            search_roots.append(Path(sys.executable).resolve().parent)

        candidates = [
            root / relative_path / filename
            for root in search_roots
            for relative_path in (
                Path("assets"),
                Path("cultura_epimed") / "assets"
            )
        ]
        for candidate in candidates:
            if candidate.is_file():
                return candidate

        raise FileNotFoundError(
            "Não foi possível encontrar o ícone do aplicativo. "
            "Inclua src\\cultura_epimed\\assets\\cultura-epimed.ico "
            "no pacote do aplicativo."
        )

    def _find_splash_logo(self) -> Path:
        filename = "logo-impacto-mr.png"
        search_roots = [
            Path(__file__).resolve().parents[1]
        ]
        bundle_path = getattr(sys, "_MEIPASS", None)
        if bundle_path:
            search_roots.append(Path(bundle_path))
        if getattr(sys, "frozen", False):
            search_roots.append(Path(sys.executable).resolve().parent)

        candidates = [
            root / relative_path / filename
            for root in search_roots
            for relative_path in (
                Path("assets"),
                Path("cultura_epimed") / "assets"
            )
        ]
        for candidate in candidates:
            if candidate.is_file():
                return candidate

        raise FileNotFoundError(
            "Não foi possível encontrar o logo da tela de abertura. "
            "Inclua src\\cultura_epimed\\assets\\logo-impacto-mr.png "
            "no pacote do aplicativo."
        )

    def _create_splash(self):
        self.root.title("Impacto MR")
        self.root.overrideredirect(True)
        self.root.configure(background="#F6F8FC")

        splash = tk.Frame(
            self.root,
            background="#F6F8FC"
        )
        splash.pack(
            fill="both",
            expand=True
        )

        tk.Frame(
            splash,
            height=8,
            background="#403A8D"
        ).pack(
            fill="x"
        )

        logo = tk.PhotoImage(
            file=str(self._find_splash_logo())
        )
        self.splash_logo = logo.subsample(2, 2)

        tk.Label(
            splash,
            image=self.splash_logo,
            background="#F6F8FC"
        ).pack(
            pady=(16, 2)
        )

        tk.Label(
            splash,
            text="CULTURA  |  EPIMED",
            font=("Segoe UI", 13, "bold"),
            foreground="#403A8D",
            background="#F6F8FC"
        ).pack()

        tk.Label(
            splash,
            text="Preparando o ambiente...",
            font=("Segoe UI", 10),
            foreground="#626478",
            background="#F6F8FC"
        ).pack(
            pady=(8, 0)
        )

        self.splash_progress = ttk.Progressbar(
            splash,
            mode="indeterminate",
            length=280
        )
        self.splash_progress.pack(
            pady=(14, 22)
        )
        self.splash_progress.start(12)

        width = 800
        height = 500
        x = int(
            (self.root.winfo_screenwidth() - width) / 2
        )
        y = int(
            (self.root.winfo_screenheight() - height) / 2
        )
        self.root.geometry(
            f"{width}x{height}+{x}+{y}"
        )
        self.root.update_idletasks()
        self.root.update()

    def _show_main_window(self):
        self.splash_progress.stop()
        for widget in self.root.winfo_children():
            widget.destroy()

        self.root.overrideredirect(False)
        self.root.title("Epimed Cultura")
        self.root.geometry("900x650")
        self.root.minsize(800, 550)
        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.exit_application
        )
        self._center_window()
        self.root.state("zoomed")
        self._create_components()

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
            text="CULTURA EPIMED",
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
                "Sistema para auxiliar na conversão e cruzamento "
                "entre Bases de dados"
            )
        )

        subtitulo.pack(
            pady=(0, 20)
        )

        # -------------------------
        # Botões
        # -------------------------

        self.btn_convert = ttk.Menubutton(
            container,
            text="Conversões"
        )

        self.conversion_menu = tk.Menu(
            self.btn_convert,
            tearoff=False
        )
        self.conversion_menu.add_command(
            label="Word para Excel",
            command=lambda: self.choose_conversion("docx")
        )
        self.conversion_menu.add_command(
            label="Excel (xls) para Excel",
            command=lambda: self.choose_conversion("xls")
        )
        self.conversion_menu.add_command(
            label="PDF para Excel",
            command=lambda: self.choose_conversion("pdf")
        )
        self.conversion_menu.add_command(
            label="CSV para Excel",
            command=lambda: self.choose_conversion("csv")
        )
        self.btn_convert["menu"] = self.conversion_menu

        self.btn_convert.pack(
            fill="x",
            pady=3
        )

        self.selected_files_label = ttk.Label(
            container,
            text="Nenhuma conversão ou arquivo selecionado.",
            wraplength=800
        )
        self.selected_files_label.pack(
            fill="x",
            pady=(5, 3)
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

        self.progress_label = ttk.Label(
            container,
            text="Status"
        )
        self.progress_label.pack(
            pady=(20, 5)
        )

        self.progress = ttk.Progressbar(
            container,
            mode="determinate",
            maximum=100
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
            height=10
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

    def _begin_progress(self, description):
        self.progress_description = description
        self.progress_value = 0
        self.progress.configure(value=0)
        self.progress_label.configure(
            text=f"{description}: 0%"
        )
        self.root.update_idletasks()

    def _update_progress(self, percent):
        self.progress_value = max(
            0,
            min(100, int(percent))
        )
        self.progress.configure(value=self.progress_value)
        self.progress_label.configure(
            text=f"{self.progress_description}: {self.progress_value}%"
        )
        self.root.update_idletasks()

    def _finish_progress(self, successful):
        if successful:
            self._update_progress(100)
            self.progress_label.configure(
                text=f"{self.progress_description}: concluído (100%)"
            )
        else:
            self.progress_label.configure(
                text=(
                    f"{self.progress_description}: interrompido em "
                    f"{self.progress_value}%"
                )
            )
        self.root.update_idletasks()

    # ==================================
    # SELEÇÃO DE ARQUIVOS
    # ==================================

    def choose_conversion(self, file_type):

        self.selected_conversion_type = file_type
        self.selected_source_folder = None
        self.selected_files = []
        self.selected_file_count = 0

        self.select_files()

    def select_files(self):

        if not self.selected_conversion_type:
            return

        file_types = ("docx", "xls", "pdf", "csv")
        ordered_file_types = (
            self.selected_conversion_type,
            *(
                file_type
                for file_type in file_types
                if file_type != self.selected_conversion_type
            )
        )
        selected_files = filedialog.askopenfilenames(
            title=(
                "Selecione um ou mais arquivos "
                f"{FILE_TYPE_LABELS[self.selected_conversion_type]}"
            ),
            initialdir=str(self.base_path),
            filetypes=[
                (
                    f"Arquivos {FILE_TYPE_LABELS[file_type]} "
                    f"(*.{file_type})",
                    f"*.{file_type}"
                )
                for file_type in ordered_file_types
            ] + [
                ("Todos os arquivos", "*.*")
            ]
        )

        if not selected_files:
            self.selected_files_label.configure(
                text="Nenhum arquivo selecionado."
            )
            return

        self.selected_files = [
            Path(file)
            for file in selected_files
        ]
        invalid_files = [
            file
            for file in self.selected_files
            if file.suffix.lower()
            != f".{self.selected_conversion_type}"
        ]
        if invalid_files:
            invalid_names = "\n".join(
                f"- {file.name}"
                for file in invalid_files
            )
            self.selected_files = []
            self.selected_files_label.configure(
                text="Os arquivos selecionados não correspondem ao formato."
            )
            messagebox.showerror(
                "Formato de arquivo incorreto",
                (
                    f"Para esta conversão, selecione somente arquivos "
                    f".{self.selected_conversion_type.upper()}:\n\n"
                    f"{invalid_names}"
                )
            )
            return

        source_folders = {
            file.parent.resolve()
            for file in self.selected_files
        }
        if len(source_folders) > 1:
            self.selected_files = []
            messagebox.showerror(
                "Selecione arquivos de uma pasta",
                "Os arquivos selecionados devem estar na mesma pasta."
            )
            return

        self.selected_source_folder = self.selected_files[0].parent
        self.selected_file_count = len(self.selected_files)
        selected_names = [
            file.name
            for file in self.selected_files[:3]
        ]
        if self.selected_file_count > 3:
            selected_names.append(
                f"e mais {self.selected_file_count - 3}"
            )

        self.selected_files_label.configure(
            text=(
                f"{self.selected_file_count} arquivo(s) "
                f"{FILE_TYPE_LABELS[self.selected_conversion_type]} "
                f"selecionado(s): "
                f"{', '.join(selected_names)}"
            )
        )

        self.log(
            f"{self.selected_file_count} arquivo(s) selecionado(s) "
            f"na pasta {self.selected_source_folder}."
        )

        if self.selected_conversion_type == "pdf":
            messagebox.showinfo(
                "Fluxo em construção",
                "Os arquivos PDF foram selecionados, mas a conversão para "
                "Excel ainda está em construção."
            )
            return

        self.convert_files()

    # ==================================
    # CONVERSÃO
    # ==================================

    def convert_files(self):

        if (
            self.selected_conversion_type == "pdf"
        ):
            messagebox.showinfo(
                "Fluxo em construção",
                "A conversão de PDF para Excel ainda está em construção."
            )
            return

        if (
            not self.selected_source_folder
            or not self.selected_files
            or not self.selected_conversion_type
        ):
            messagebox.showwarning(
                "Arquivos não selecionados",
                "Selecione uma conversão e os arquivos antes de processar."
            )
            return

        destination_folder = (
            self.selected_source_folder
            if self.selected_conversion_type == "xls"
            else self.selected_source_folder / "xlsx"
        )

        try:

            self._begin_progress("Convertendo arquivos")

            self.log(
                f"Iniciando conversão de "
                f"{FILE_TYPE_LABELS[self.selected_conversion_type]} para "
                "Excel..."
            )

            total = (
                self.converter.convert(
                    self.selected_source_folder,
                    destination_folder,
                    self.selected_conversion_type,
                    selected_files=self.selected_files,
                    progress_callback=self._update_progress
                )
            )

            self._finish_progress(True)

            self.log(
                f"{total} arquivo(s) convertido(s) para Excel em "
                f"{destination_folder}."
            )

            if self._is_cultura_folder(self.selected_source_folder):
                self.cultura_xlsx_folder = destination_folder
            else:
                self.epimed_xlsx_folder = destination_folder

            messagebox.showinfo(
                "Conversão Finalizada",
                (
                    f"{total} arquivo(s) convertido(s) para Excel.\n\n"
                    f"Pasta de destino:\n{destination_folder}"
                )
            )

        except Exception as ex:

            self._finish_progress(False)
            self.log(f"Falha na conversão: {ex}")

            messagebox.showerror(
                "Falha na conversão",
                str(ex)
            )

    # ==================================
    # MERGE
    # ==================================

    def generate_base(self):

        messagebox.showinfo(
            "Etapa 1 de 2 — Bases de Cultura",
            (
                "Selecione agora um ou mais arquivos Excel (.xlsx) "
                "com as bases de Cultura."
            )
        )
        cultura_initial_folder = (
            self.cultura_xlsx_folder
            or self.folder_service.culturas
        )
        cultura_files = self._ask_merge_files(
            "Cultura",
            cultura_initial_folder
        )
        if not cultura_files:
            return

        if not self._validate_merge_files("Cultura", cultura_files):
            return

        try:
            self._begin_progress("Lendo colunas de Cultura")
            cultura_columns = self.merge_service.get_cultura_columns(
                cultura_files,
                progress_callback=self._update_progress
            )
            self._finish_progress(True)
        except Exception as ex:
            self._finish_progress(False)
            messagebox.showerror(
                "Não foi possível ler as colunas de Cultura",
                str(ex)
            )
            return

        cultura_column = self._select_crossing_column(
            "Cultura",
            cultura_columns
        )
        if not cultura_column:
            return

        messagebox.showinfo(
            "Etapa 2 de 2 — Bases do Epimed",
            (
                "Agora selecione um ou mais arquivos Excel (.xlsx) "
                "com as bases do sistema Epimed."
            )
        )
        epimed_initial_folder = (
            self.epimed_xlsx_folder
            or self.folder_service.xlsx
        )
        epimed_files = self._ask_merge_files(
            "Epimed",
            epimed_initial_folder
        )
        if not epimed_files:
            return

        if not self._validate_merge_files("Epimed", epimed_files):
            return

        self.log(
            f"Arquivos selecionados: {len(cultura_files)} de Cultura e "
            f"{len(epimed_files)} de Epimed."
        )

        try:
            self._begin_progress("Lendo colunas do Epimed")
            _, epimed_columns = (
                self.merge_service.get_crossing_columns(
                    cultura_files,
                    epimed_files[0].parent,
                    epimed_files=epimed_files,
                    progress_callback=self._update_progress
                )
            )
            self._finish_progress(True)
        except Exception as ex:
            self._finish_progress(False)
            messagebox.showerror(
                "Não foi possível identificar os campos de cruzamento",
                str(ex)
            )
            return

        epimed_column = self._select_crossing_column(
            "Epimed",
            epimed_columns
        )
        if not epimed_column:
            return

        output_filename = self.merge_service.get_output_filename()
        messagebox.showinfo(
            "Etapa final — Salvar a base",
            (
                "Escolha a pasta onde deseja salvar a base final. "
                f"O arquivo será criado com o nome {output_filename}."
            )
        )
        output_folder = filedialog.askdirectory(
            title="Selecione onde salvar a base final",
            initialdir=str(
                self.folder_service.saida
                if self.folder_service.saida.exists()
                else self.base_path
            ),
            mustexist=True
        )
        if not output_folder:
            return

        try:

            self._begin_progress("Gerando a base final")

            self.log(
                f"Iniciando geração da base em {output_folder}..."
            )

            output_file = (
                self.merge_service.execute(
                    cultura_folder=
                        cultura_files[0].parent,

                    epimed_xlsx_folder=
                        epimed_files[0].parent,

                    output_folder=
                        Path(output_folder),

                    cultura_files=cultura_files,
                    cultura_column=cultura_column,
                    epimed_column=epimed_column,
                    epimed_files=epimed_files,
                    progress_callback=self._update_progress,
                    output_filename=output_filename
                )
            )

            self._finish_progress(True)

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
            self._finish_progress(False)

            messagebox.showerror(
                "Não foi possível gerar a base",
                str(ex)
            )

    def _validate_merge_files(
        self,
        data_name: str,
        files: list[Path]
    ):
        invalid_files = [
            file
            for file in files
            if not file.is_file()
            or file.suffix.lower() != ".xlsx"
        ]
        if not invalid_files:
            return True

        invalid_names = "\n".join(
            f"- {file.name}"
            for file in invalid_files
        )
        messagebox.showerror(
            "Formato de arquivo incorreto",
            (
                f"Selecione somente arquivos Excel (.xlsx) de "
                f"{data_name}:\n\n"
                f"{invalid_names}"
            )
        )
        return False

    def _ask_merge_files(
        self,
        data_name: str,
        initial_folder: Path
    ):
        files = filedialog.askopenfilenames(
            title=f"Selecione um ou mais arquivos Excel de {data_name}",
            initialdir=str(
                initial_folder
                if initial_folder.exists()
                else self.base_path
            ),
            filetypes=[
                ("Arquivos Excel (*.xlsx)", "*.xlsx"),
                ("Todos os arquivos", "*.*")
            ]
        )
        return [Path(file) for file in files]

    def _is_cultura_folder(
        self,
        folder: Path
    ):
        folder = folder.resolve()
        cultura_folder = self.folder_service.culturas.resolve()

        return (
            folder == cultura_folder
            or cultura_folder in folder.parents
        )

    def _select_crossing_column(
        self,
        data_name: str,
        columns: list[str]
    ):
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Campo de cruzamento — {data_name}")
        dialog.transient(self.root)
        dialog.resizable(False, False)

        ttk.Label(
            dialog,
            text=(
                f"Selecione a coluna de {data_name} que será usada no "
                "cruzamento:"
            ),
            wraplength=450
        ).pack(
            padx=20,
            pady=(20, 8)
        )

        ttk.Label(
            dialog,
            text=f"Coluna {data_name}"
        ).pack(
            padx=20,
            pady=(0, 8)
        )

        selected_column = tk.StringVar()
        column_picker = ttk.Combobox(
            dialog,
            textvariable=selected_column,
            values=columns,
            state="readonly",
            width=50
        )
        column_picker.pack(
            padx=20,
            pady=5
        )

        result = {"column": None}

        def confirm():
            column = selected_column.get()
            if not column:
                messagebox.showwarning(
                    "Campo obrigatório",
                    f"Selecione a coluna de {data_name}.",
                    parent=dialog
                )
                return

            result["column"] = column
            dialog.destroy()

        def cancel():
            dialog.destroy()

        buttons = ttk.Frame(dialog)
        buttons.pack(
            padx=20,
            pady=(12, 20)
        )
        ttk.Button(
            buttons,
            text="Continuar",
            command=confirm
        ).pack(
            side="left",
            padx=5
        )
        ttk.Button(
            buttons,
            text="Cancelar",
            command=cancel
        ).pack(
            side="left",
            padx=5
        )

        dialog.update_idletasks()
        dialog_x = self.root.winfo_rootx() + (
            self.root.winfo_width() - dialog.winfo_width()
        ) // 2
        dialog_y = self.root.winfo_rooty() + (
            self.root.winfo_height() - dialog.winfo_height()
        ) // 2
        dialog.geometry(
            f"+{max(0, dialog_x)}+{max(0, dialog_y)}"
        )

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.grab_set()
        self.root.wait_window(dialog)

        return result["column"]

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

        minimum_splash_duration_ms = 2000
        elapsed_ms = int(
            (time.monotonic() - self.splash_started_at) * 1000
        )
        self.root.after(
            max(0, minimum_splash_duration_ms - elapsed_ms),
            self._show_main_window
        )
        self.root.mainloop()