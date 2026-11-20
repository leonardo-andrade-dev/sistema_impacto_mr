from pathlib import Path

import pandas as pd
from docx import Document


class DocxToExcelConverter:

    def convert(
        self,
        source_folder: Path,
        destination_folder: Path
    ):

        destination_folder.mkdir(
            parents=True,
            exist_ok=True
        )

        docx_files = list(
            source_folder.glob("*.docx")
        )

        if not docx_files:

            raise Exception(
                "Nenhum arquivo DOCX encontrado na pasta Dados_Epimed."
            )

        converted_count = 0

        for docx_file in docx_files:

            try:

                self._convert_file(
                    docx_file,
                    destination_folder
                )

                converted_count += 1

            except Exception as ex:

                print(
                    f"Erro ao converter "
                    f"{docx_file.name}: {ex}"
                )

        return converted_count

    def _convert_file(
        self,
        docx_file: Path,
        destination_folder: Path
    ):

        document = Document(
            docx_file
        )

        if not document.tables:

            print(
                f"Nenhuma tabela encontrada em "
                f"{docx_file.name}"
            )

            return

        table = document.tables[0]

        data = []

        for row in table.rows:

            row_data = []

            for cell in row.cells:

                row_data.append(
                    cell.text.strip()
                )

            if any(
                value.strip()
                for value in row_data
            ):

                data.append(
                    row_data
                )

        if not data:

            print(
                f"Tabela vazia em "
                f"{docx_file.name}"
            )

            return

        max_columns = max(
            len(row)
            for row in data
        )

        normalized_rows = []

        for row in data:

            row = row + (
                [""] *
                (max_columns - len(row))
            )

            normalized_rows.append(
                row
            )

        header = normalized_rows[0]

        rows = normalized_rows[1:]

        while header and not header[-1].strip():
            header.pop()

        leading_blank_columns = 0

        if len(header) < max_columns:

            for row in rows:

                if not any(row):
                    continue

                leading_blank_columns = next(
                    (
                        index
                        for index, value in enumerate(row)
                        if value.strip()
                    ),
                    0
                )

                break

        if leading_blank_columns > 0:
            rows = [
                row[
                    leading_blank_columns:
                    leading_blank_columns + len(header)
                ]
                for row in rows
            ]

        header = [
            value.strip()
            or f"Coluna_{index + 1}"
            for index, value in enumerate(header)
        ]

        df = pd.DataFrame(
            rows,
            columns=header
        )

        # Remove colunas totalmente vazias
        df = df.dropna(
            axis=1,
            how="all"
        )

        # Remove colunas "Unnamed"
        df = df.loc[
            :,
            ~df.columns.astype(str)
            .str.contains(
                "^Unnamed",
                na=False
            )
        ]

        # Remove colunas vazias
        df = df.loc[
            :,
            df.columns.astype(str)
            .str.strip()
            != ""
        ]

        output_file = (
            destination_folder /
            f"{docx_file.stem}.xlsx"
        )

        df.to_excel(
            output_file,
            index=False,
            engine="openpyxl"
        )

        print(
            f"Convertido: "
            f"{docx_file.name}"
        )