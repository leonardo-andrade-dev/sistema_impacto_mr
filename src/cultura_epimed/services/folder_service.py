from pathlib import Path
import shutil


class FolderService:

    def __init__(self, root_path: Path):
        self.root_path = root_path

        self.culturas = root_path / "Culturas"
        self.dados_epimed = root_path / "Dados_Epimed"
        self.xlsx = self.dados_epimed / "xlsx"
        self.saida = root_path / "Saida"

    def create_folders(self):

        self.culturas.mkdir(
            parents=True,
            exist_ok=True
        )
        self.dados_epimed.mkdir(exist_ok=True)
        self.xlsx.mkdir(exist_ok=True)
        self.saida.mkdir(exist_ok=True)

    def clear_folders(self):

        folders = [
            self.culturas,
            self.dados_epimed,
            self.xlsx,
            self.saida
        ]

        for folder in folders:

            if not folder.exists():
                continue

            for item in folder.iterdir():

                if item.is_file():
                    item.unlink()

                elif item.is_dir():
                    shutil.rmtree(item)