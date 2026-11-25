from pathlib import Path

from cultura_epimed.services.docx_to_excel_converter import (
    DocxToExcelConverter
)
from cultura_epimed.services.folder_service import FolderService
from cultura_epimed.services.merge_service import MergeService


class EpimedCultureSystem:

    def __init__(self, root_path: Path | None = None):

        self.folder_service = FolderService(
            root_path or Path.cwd()
        )

        self.converter = DocxToExcelConverter()

        self.merge_service = MergeService()

    def create_environment(self):

        self.folder_service.create_folders()

    def execute(self):

        self.create_environment()

        converted_count = self.converter.convert(
            source_folder=self.folder_service.dados_epimed,
            destination_folder=self.folder_service.xlsx
        )

        output_file = self.merge_service.execute(
            cultura_folder=self.folder_service.culturas,
            epimed_xlsx_folder=self.folder_service.xlsx,
            output_folder=self.folder_service.saida
        )

        return converted_count, output_file