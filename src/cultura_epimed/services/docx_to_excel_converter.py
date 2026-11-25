from collections.abc import Callable
from pathlib import Path
import zipfile

import pandas as pd
from docx import Document


class DocxToExcelConverter:

    def convert(
        self,
        source_folder: Path,
        destination_folder: Path,
        file_type: str = "docx",
        selected_files: list[Path] | None = None,
        progress_callback: Callable[[int], None] | None = None
    ):

        file_type = file_type.lower().lstrip(".")

        if file_type == "pdf":
            raise ValueError(
                "A conversão de PDF para XLSX ainda está em construção."
            )

        converters = {
            "docx": self._convert_docx_file,
            "xls": self._convert_xls_file,
            "csv": self._convert_csv_file
        }

        if file_type not in converters:
            raise ValueError(
                f"Formato de origem não suportado: .{file_type}"
            )

        if selected_files is None:
            files = sorted(
                file
                for file in source_folder.iterdir()
                if file.is_file()
                and file.suffix.lower() == f".{file_type}"
            )
        else:
            files = sorted(
                {
                    Path(file)
                    for file in selected_files
                },
                key=lambda file: file.name.casefold()
            )
            invalid_files = [
                file
                for file in files
                if not file.is_file()
                or file.suffix.lower() != f".{file_type}"
            ]
            if invalid_files:
                invalid_names = "\n".join(
                    f"- {file.name}"
                    for file in invalid_files
                )
                raise ValueError(
                    f"Selecione somente arquivos .{file_type.upper()}:\n"
                    f"{invalid_names}"
                )

        if not files:
            if selected_files is not None:
                raise ValueError(
                    f"Nenhum arquivo .{file_type.upper()} foi selecionado."
                )
            raise ValueError(
                f"Nenhum arquivo .{file_type.upper()} encontrado na pasta "
                f"selecionada:\n{source_folder}\n\n"
                "Confira se selecionou a pasta correta e se os arquivos "
                "estão diretamente nela."
            )

        destination_folder.mkdir(
            parents=True,
            exist_ok=True
        )

        converted_count = 0
        errors = []

        for index, source_file in enumerate(files, start=1):

            try:

                converters[file_type](
                    source_file,
                    destination_folder
                )

                converted_count += 1

            except Exception as ex:
                errors.append(
                    f"- {source_file.name}: {ex}"
                )

            if progress_callback is not None:
                progress_callback(
                    round(index * 100 / len(files))
                )

        if errors:
            raise RuntimeError(
                f"{converted_count} de {len(files)} arquivo(s) "
                f"convertido(s) para XLSX em:\n{destination_folder}\n\n"
                "Falhas encontradas:\n"
                + "\n".join(errors)
            )

        return converted_count

    def _convert_docx_file(
        self,
        docx_file: Path,
        destination_folder: Path
    ):

        document = Document(
            docx_file
        )

        if not document.tables:
            raise ValueError(
                "o documento não contém tabelas para conversão"
            )

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
            raise ValueError(
                "a primeira tabela do documento está vazia"
            )

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

    def _convert_xls_file(
        self,
        xls_file: Path,
        destination_folder: Path
    ):
        engine = (
            "openpyxl"
            if zipfile.is_zipfile(xls_file)
            else "xlrd"
        )

        sheets = pd.read_excel(
            xls_file,
            sheet_name=None,
            engine=engine
        )

        for sheet_name in sheets:
            df = sheets[sheet_name]
            has_unnamed_columns = any(
                str(column).strip().lower().startswith("unnamed")
                for column in df.columns
            )
            if not has_unnamed_columns:
                continue

            corrected_df = pd.read_excel(
                xls_file,
                sheet_name=sheet_name,
                header=1,
                engine=engine
            )
            has_named_columns = any(
                not str(column).strip().lower().startswith("unnamed")
                for column in corrected_df.columns
            )
            if has_named_columns:
                sheets[sheet_name] = corrected_df

        output_file = destination_folder / f"{xls_file.stem}.xlsx"
        with pd.ExcelWriter(
            output_file,
            engine="openpyxl"
        ) as writer:
            for sheet_name, df in sheets.items():
                df.to_excel(
                    writer,
                    sheet_name=sheet_name[:31],
                    index=False
                )

    def _convert_csv_file(
        self,
        csv_file: Path,
        destination_folder: Path
    ):
        try:
            df = pd.read_csv(
                csv_file,
                sep=None,
                engine="python",
                encoding="utf-8-sig"
            )
        except UnicodeDecodeError:
            df = pd.read_csv(
                csv_file,
                sep=None,
                engine="python",
                encoding="cp1252"
            )

        self._write_excel(
            df,
            csv_file,
            destination_folder
        )

    def _write_excel(
        self,
        df: pd.DataFrame,
        source_file: Path,
        destination_folder: Path
    ):
        output_file = destination_folder / f"{source_file.stem}.xlsx"

        df.to_excel(
            output_file,
            index=False,
            engine="openpyxl"
        )