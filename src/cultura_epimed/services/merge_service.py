from collections.abc import Callable
from datetime import date
from pathlib import Path
import re
from unidecode import unidecode
import pandas as pd

from cultura_epimed.domain.execution_log import ExecutionLog
from cultura_epimed.domain.name_normalizer import NameNormalizer


class MergeService:

    MESES = [
        "JANEIRO",
        "FEVEREIRO",
        "MARCO",
        "ABRIL",
        "MAIO",
        "JUNHO",
        "JULHO",
        "AGOSTO",
        "SETEMBRO",
        "OUTUBRO",
        "NOVEMBRO",
        "DEZEMBRO"
    ]

    ABREVIACOES_MESES = {
        "JANEIRO": {"JAN"},
        "FEVEREIRO": {"FEV"},
        "MARCO": {"MAR"},
        "ABRIL": {"ABR"},
        "MAIO": {"MAI"},
        "JUNHO": {"JUN"},
        "JULHO": {"JUL"},
        "AGOSTO": {"AGO"},
        "SETEMBRO": {"SET"},
        "OUTUBRO": {"OUT"},
        "NOVEMBRO": {"NOV"},
        "DEZEMBRO": {"DEZ"}
    }

    def __init__(self):

        self.normalizer = NameNormalizer()

        self.logs = []

        self.nao_encontrados_geral = []
        self._epimed_sheet_cache = {}
        self._epimed_header_fallbacks = {}

    @staticmethod
    def get_output_filename() -> str:
        return f"Cultura_Epimed_{date.today():%d_%m_%Y}.xlsx"

    def _find_cultura_file(
        self,
        cultura_folder: Path
    ):

        arquivos = list(
            cultura_folder.glob("*.xlsx")
        )

        if not arquivos:

            raise Exception(
                "Nenhum arquivo XLSX encontrado na pasta Culturas."
            )

        return arquivos[0]

    def get_cultura_columns(
        self,
        cultura_files: Path | list[Path],
        progress_callback: Callable[[int], None] | None = None
    ):
        files = (
            [cultura_files]
            if isinstance(cultura_files, Path)
            else cultura_files
        )
        frames = self._read_cultura_frames(
            files,
            progress_callback=progress_callback
        )
        columns = self._common_columns(frames)
        if not columns:
            raise ValueError(
                "As planilhas de Cultura não têm colunas com cabeçalhos "
                "idênticos em comum."
            )
        return columns

    def _read_cultura_frames(
        self,
        cultura_files: list[Path],
        progress_callback: Callable[[int], None] | None = None
    ):
        frames = []
        for index, cultura_file in enumerate(cultura_files, start=1):
            for sheet_name, df in self._read_cultura_sheets(cultura_file):
                frames.append(
                    (
                        f"{cultura_file.name} / {sheet_name}",
                        df
                    )
                )
            if progress_callback is not None:
                progress_callback(
                    round(index * 90 / len(cultura_files))
                )

        if not frames:
            raise ValueError(
                "Não encontrei abas processáveis nas planilhas de Cultura."
            )

        return frames

    def _read_cultura_sheets(
        self,
        cultura_file: Path
    ):
        try:
            with pd.ExcelFile(
                cultura_file,
                engine="openpyxl"
            ) as workbook:
                sheets = []
                for sheet_name in workbook.sheet_names:
                    if self._is_export_sheet(sheet_name):
                        continue
                    df = self._read_sheet_with_header_detection(
                        workbook,
                        sheet_name
                    )
                    sheets.append(
                        (
                            sheet_name,
                            self._clean_columns(df)
                        )
                    )
        except Exception as ex:
            raise ValueError(
                f"Não foi possível ler a planilha de Cultura "
                f"'{cultura_file.name}': {ex}"
            ) from ex

        if not sheets:
            raise ValueError(
                f"Não encontrei abas processáveis na planilha de Cultura "
                f"'{cultura_file.name}'."
            )

        return sheets

    def _read_sheet_with_header_detection(
        self,
        workbook,
        sheet_name: str
    ):
        first_row_header = pd.read_excel(
            workbook,
            sheet_name=sheet_name,
            header=0
        )
        second_row_header = pd.read_excel(
            workbook,
            sheet_name=sheet_name,
            header=1
        )
        if len(self._named_columns(second_row_header)) > len(
            self._named_columns(first_row_header)
        ):
            return second_row_header
        return first_row_header

    def _read_epimed_sheet(
        self,
        epimed_file: Path,
        sheet_name,
        fallback_columns: list[str] | None = None
    ):
        selected_sheet = sheet_name or 0
        if fallback_columns is not None:
            df = pd.read_excel(
                epimed_file,
                sheet_name=selected_sheet,
                header=None,
                engine="openpyxl"
            )
            if len(df.columns) != len(fallback_columns):
                raise ValueError(
                    f"a estrutura sem cabeçalho possui {len(df.columns)} "
                    f"colunas, mas os outros meses possuem "
                    f"{len(fallback_columns)}"
                )
            df.columns = fallback_columns
        else:
            with pd.ExcelFile(
                epimed_file,
                engine="openpyxl"
            ) as workbook:
                df = self._read_sheet_with_header_detection(
                    workbook,
                    selected_sheet
                )

        return self._clean_columns(df)

    def _header_looks_like_data(
        self,
        df
    ):
        columns = self._named_columns(df)
        numeric_labels = sum(
            bool(
                re.fullmatch(
                    r"\d+(?:[.,]\d+)?",
                    column
                )
            )
            for column in columns
        )
        date_labels = sum(
            bool(
                re.search(
                    r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}",
                    column
                )
            )
            for column in columns
        )
        return numeric_labels >= 2 or date_labels > 0

    def _common_columns(
        self,
        frames
    ):
        first_frame = frames[0][1]
        common_columns = self._named_columns(first_frame)

        for frame in frames[1:]:
            df = frame[1]
            available_columns = set(self._named_columns(df))
            common_columns = [
                column
                for column in common_columns
                if column in available_columns
            ]

        return common_columns

    def _named_columns(
        self,
        df
    ):
        return [
            str(column).strip()
            for column in df.columns
            if str(column).strip()
            and not str(column).strip().lower().startswith("unnamed")
        ]

    def _is_export_sheet(
        self,
        sheet_name: str
    ):
        return "EXPORT" in unidecode(sheet_name).upper()

    def get_crossing_columns(
        self,
        cultura_files: list[Path],
        epimed_xlsx_folder: Path,
        epimed_files: list[Path] | None = None,
        progress_callback: Callable[[int], None] | None = None
    ):
        cultura_columns = None
        epimed_columns = None
        epimed_schemas = []
        headerless_sources = []
        matched_epimed_sources = set()
        missing_epimed_references = []
        self._epimed_sheet_cache.clear()
        self._epimed_header_fallbacks.clear()
        epimed_source_files = (
            epimed_files
            if epimed_files is not None
            else list(epimed_xlsx_folder.glob("*.xlsx"))
        )
        total_epimed_files = len(
            {
                Path(file).resolve()
                for file in epimed_source_files
            }
        )
        processed_epimed_files = set()

        for cultura_file in cultura_files:
            cultura_sheets = self._read_cultura_sheets(cultura_file)

            for sheet_name, df_cultura in cultura_sheets:
                columns = self._named_columns(df_cultura)
                if cultura_columns is None:
                    cultura_columns = columns
                else:
                    available = set(columns)
                    cultura_columns = [
                        column
                        for column in cultura_columns
                        if column in available
                    ]

                reference_name = (
                    sheet_name
                    if self._extract_month(sheet_name)
                    else cultura_file.stem
                )
                epimed_file, epimed_sheet = self._find_epimed_source(
                    epimed_xlsx_folder,
                    reference_name,
                    epimed_files
                )
                if epimed_file is None:
                    if epimed_files is not None:
                        missing_epimed_references.append(reference_name)
                    continue

                source_key = (epimed_file, epimed_sheet or 0)
                if source_key in matched_epimed_sources:
                    continue
                matched_epimed_sources.add(source_key)

                try:
                    df_epimed = self._read_epimed_sheet(
                        epimed_file,
                        epimed_sheet
                    )
                except Exception as ex:
                    raise ValueError(
                        f"Não foi possível ler os dados Epimed de "
                        f"'{epimed_file.name}': {ex}"
                    ) from ex

                resolved_epimed_file = epimed_file.resolve()
                if resolved_epimed_file not in processed_epimed_files:
                    processed_epimed_files.add(resolved_epimed_file)
                    if progress_callback is not None and total_epimed_files:
                        progress_callback(
                            round(
                                len(processed_epimed_files)
                                * 90
                                / total_epimed_files
                            )
                        )

                if self._header_looks_like_data(df_epimed):
                    headerless_sources.append(
                        (
                            source_key,
                            epimed_file,
                            epimed_sheet,
                            df_epimed
                        )
                    )
                    continue

                columns = self._named_columns(df_epimed)
                epimed_schemas.append(columns)
                if epimed_columns is None:
                    epimed_columns = columns
                else:
                    available = set(columns)
                    epimed_columns = [
                        column
                        for column in epimed_columns
                        if column in available
                    ]

        if not cultura_columns:
            raise ValueError(
                "As planilhas de Cultura não têm colunas comuns para "
                "selecionar. Confira os cabeçalhos de todas as abas."
            )

        if missing_epimed_references:
            missing_months = ", ".join(
                dict.fromkeys(missing_epimed_references)
            )
            raise ValueError(
                "Não encontrei arquivos ou abas Epimed selecionados para "
                f"todos os meses da Cultura. Sem correspondência para: "
                f"{missing_months}."
            )

        if not matched_epimed_sources:
            raise ValueError(
                "Não encontrei arquivos ou abas Epimed correspondentes "
                "aos meses das planilhas de Cultura."
            )

        if headerless_sources:
            if len(epimed_schemas) < 2:
                raise ValueError(
                    "Há um arquivo Epimed sem cabeçalho identificável, mas "
                    "não existem pelo menos dois outros meses com o mesmo "
                    "cabeçalho para validar uma inferência segura."
                )

            reference_schema = epimed_schemas[0]
            if any(
                schema != reference_schema
                for schema in epimed_schemas[1:]
            ):
                raise ValueError(
                    "Há um arquivo Epimed sem cabeçalho identificável, e "
                    "os cabeçalhos dos outros meses não têm a mesma ordem. "
                    "Não é seguro inferir as colunas."
                )

            for (
                source_key,
                epimed_file,
                epimed_sheet,
                df_epimed
            ) in headerless_sources:
                raw_df = pd.read_excel(
                    epimed_file,
                    sheet_name=epimed_sheet or 0,
                    header=None,
                    engine="openpyxl"
                )
                if len(raw_df.columns) != len(reference_schema):
                    raise ValueError(
                        f"O arquivo Epimed '{epimed_file.name}' parece "
                        f"estar sem cabeçalho e possui "
                        f"{len(raw_df.columns)} colunas, enquanto os outros "
                        f"meses possuem {len(reference_schema)}. Não é "
                        "seguro inferir os campos. Corrija o arquivo de "
                        "origem para incluir cabeçalhos compatíveis com os "
                        "outros meses e gere novamente o XLSX."
                    )
                self._epimed_header_fallbacks[source_key] = (
                    reference_schema
                )
                if epimed_columns is None:
                    epimed_columns = reference_schema
                else:
                    available = set(reference_schema)
                    epimed_columns = [
                        column
                        for column in epimed_columns
                        if column in available
                    ]

        if not epimed_columns:
            raise ValueError(
                "Os arquivos ou abas Epimed correspondentes não têm uma "
                "coluna comum. Confira os cabeçalhos."
            )

        return cultura_columns, epimed_columns

    def _clean_columns(
        self,
        df
    ):

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        return df

    def _create_key(
        self,
        df,
        column_name
    ):

        df["CHAVE"] = (
            df[column_name]
            .apply(
                self.normalizer.normalize
            )
        )

        return df

    def _extract_month(
        self,
        sheet_name
    ):

        normalized = unidecode(sheet_name).upper()
        tokens = set(
            re.findall(
                r"[A-Z]+",
                normalized
            )
        )

        for mes in self.MESES:

            if (
                mes in normalized
                or tokens.intersection(
                    self.ABREVIACOES_MESES[mes]
                )
            ):
                return mes

        return None

    def _epimed_sheet_names(
        self,
        epimed_file: Path
    ):
        if epimed_file not in self._epimed_sheet_cache:
            try:
                with pd.ExcelFile(
                    epimed_file,
                    engine="openpyxl"
                ) as workbook:
                    self._epimed_sheet_cache[epimed_file] = (
                        workbook.sheet_names
                    )
            except Exception as ex:
                raise ValueError(
                    f"Não foi possível ler o arquivo Epimed "
                    f"'{epimed_file.name}': {ex}"
                ) from ex

        return self._epimed_sheet_cache[epimed_file]

    def _find_epimed_sheet(
        self,
        epimed_file: Path,
        month: str
    ):
        for sheet_name in self._epimed_sheet_names(epimed_file):
            if self._extract_month(sheet_name) == month:
                return sheet_name

        return None

    def _find_epimed_source(
        self,
        xlsx_folder: Path,
        reference_name: str,
        epimed_files: list[Path] | None = None
    ):
        month = self._extract_month(reference_name)
        if not month:
            return None, None

        aliases = self.ABREVIACOES_MESES[month] | {month}
        files = sorted(
            (
                epimed_files
                if epimed_files is not None
                else xlsx_folder.glob("*.xlsx")
            ),
            key=lambda file: file.name.casefold()
        )

        named_candidates = []
        for epimed_file in files:
            tokens = set(
                re.findall(
                    r"[A-Z]+",
                    unidecode(epimed_file.stem).upper()
                )
            )
            if tokens.intersection(aliases):
                named_candidates.append(epimed_file)

        for epimed_file in files:
            sheet_name = self._find_epimed_sheet(epimed_file, month)
            if sheet_name:
                return epimed_file, sheet_name

        if named_candidates:
            return named_candidates[0], None

        return None, None

    def _find_epimed_file(
        self,
        xlsx_folder: Path,
        sheet_name: str
    ):
        epimed_file, _ = self._find_epimed_source(
            xlsx_folder,
            sheet_name
        )
        return epimed_file

    def _unique_sheet_name(
        self,
        requested_name: str,
        existing_names: set[str]
    ):
        base_name = re.sub(
            r"[\[\]:*?/\\]",
            "_",
            requested_name
        ).strip().strip("'")[:31]

        if not base_name:
            base_name = "Dados"

        candidate = base_name
        suffix_number = 2
        while candidate.casefold() in existing_names:
            suffix = f"_{suffix_number}"
            candidate = f"{base_name[:31 - len(suffix)]}{suffix}"
            suffix_number += 1

        existing_names.add(candidate.casefold())
        return candidate

    def execute(
        self,
        cultura_folder: Path,
        epimed_xlsx_folder: Path,
        output_folder: Path,
        cultura_file: Path | None = None,
        cultura_column: str = "INICIAL DO PACIENTE",
        cultura_files: list[Path] | None = None,
        epimed_column: str = "Paciente",
        epimed_files: list[Path] | None = None,
        progress_callback: Callable[[int], None] | None = None,
        output_filename: str | None = None
    ):

        if cultura_files is not None:
            files = [Path(file) for file in cultura_files]
        elif cultura_file is None:
            files = [self._find_cultura_file(cultura_folder)]
        else:
            files = [Path(cultura_file)]

        files = list(
            dict.fromkeys(file.resolve() for file in files)
        )
        if not files:
            raise ValueError(
                "Nenhum arquivo XLSX de Cultura foi informado."
            )

        def culture_file_sort_key(file):
            month = self._extract_month(file.stem)
            month_index = (
                self.MESES.index(month)
                if month
                else len(self.MESES)
            )
            return month_index, file.name.casefold()

        files.sort(key=culture_file_sort_key)
        cultura_abas = []

        for file_index, cultura_file in enumerate(files, start=1):
            abas = self._read_cultura_sheets(cultura_file)

            for sheet_name, df_hj in abas:
                if cultura_column not in df_hj.columns:
                    available_columns = ", ".join(
                        str(column)
                        for column in df_hj.columns
                    )
                    raise ValueError(
                        f"A coluna selecionada '{cultura_column}' não foi "
                        f"encontrada na planilha '{cultura_file.name}', "
                        f"aba '{sheet_name}'. Colunas disponíveis: "
                        f"{available_columns or 'nenhuma'}."
                    )

                sheet_month = self._extract_month(sheet_name)
                reference_name = (
                    sheet_name
                    if sheet_month
                    else cultura_file.stem
                )
                output_sheet_name = (
                    sheet_name
                    if sheet_month
                    else cultura_file.stem
                )
                cultura_abas.append(
                    (
                        reference_name,
                        output_sheet_name,
                        df_hj
                    )
                )
            if progress_callback is not None:
                progress_callback(
                    round(file_index * 25 / len(files))
                )

        if not cultura_abas:
            raise ValueError(
                "Não encontrei abas processáveis nas planilhas de Cultura."
            )

        self._epimed_sheet_cache.clear()
        prepared_abas = []
        total_abas = len(cultura_abas)
        for aba_index, (
            reference_name,
            output_sheet_name,
            df_hj
        ) in enumerate(cultura_abas, start=1):
            epimed_file, epimed_sheet = self._find_epimed_source(
                epimed_xlsx_folder,
                reference_name,
                epimed_files
            )

            if epimed_file is None:
                prepared_abas.append(
                    (
                        reference_name,
                        output_sheet_name,
                        df_hj,
                        None
                    )
                )
                if progress_callback is not None:
                    progress_callback(
                        25 + round(30 * aba_index / total_abas)
                    )
                continue

            try:
                source_key = (epimed_file, epimed_sheet or 0)
                fallback_columns = self._epimed_header_fallbacks.get(
                    source_key
                )
                df_epi = self._read_epimed_sheet(
                    epimed_file,
                    epimed_sheet,
                    fallback_columns=fallback_columns
                )
            except Exception as ex:
                raise ValueError(
                    f"Não foi possível ler os dados Epimed de "
                    f"'{epimed_file.name}': {ex}"
                ) from ex

            df_epi = self._clean_columns(df_epi)
            if (
                fallback_columns is None
                and self._header_looks_like_data(df_epi)
            ):
                raise ValueError(
                    f"O arquivo Epimed '{epimed_file.name}' parece estar "
                    "sem cabeçalho. Execute novamente a identificação dos "
                    "campos antes de gerar a base."
                )
            if epimed_column not in df_epi.columns:
                epimed_tab = (
                    f", aba '{epimed_sheet}'"
                    if epimed_sheet
                    else ""
                )
                raise ValueError(
                    f"A coluna Epimed selecionada '{epimed_column}' não foi "
                    f"encontrada no arquivo '{epimed_file.name}'{epimed_tab}."
                )

            prepared_abas.append(
                (
                    reference_name,
                    output_sheet_name,
                    df_hj,
                    df_epi
                )
            )
            if progress_callback is not None:
                progress_callback(
                    25 + round(30 * aba_index / total_abas)
                )

        if progress_callback is not None:
            progress_callback(55)

        output_file = output_folder / (
            output_filename or self.get_output_filename()
        )

        self.logs.clear()

        self.nao_encontrados_geral.clear()

        used_sheet_names = {
            "log",
            "pacientes_nao_encontrados"
        }
        quantidade_abas = 0

        with pd.ExcelWriter(
            output_file,
            engine="openpyxl"
        ) as writer:

            for aba_index, (
                _reference_name,
                requested_sheet_name,
                df_hj,
                df_epi
            ) in enumerate(prepared_abas, start=1):
                if df_epi is None:
                    self.logs.append(
                        ExecutionLog(
                            mes=requested_sheet_name,
                            total_epimed=0,
                            encontrados=0,
                            nao_encontrados=0
                        )
                    )

                    if progress_callback is not None:
                        progress_callback(
                            55 + round(40 * aba_index / len(prepared_abas))
                        )
                    continue

                sheet_name = self._unique_sheet_name(
                    requested_sheet_name,
                    used_sheet_names
                )

                df_hj = (
                    self._create_key(
                        df_hj,
                        cultura_column
                    )
                )

                df_epi = (
                    self._create_key(
                        df_epi,
                        epimed_column
                    )
                )

                match_column = "__CULTURA_MATCH__"
                while (
                    match_column in df_hj.columns
                    or match_column in df_epi.columns
                ):
                    match_column += "_"

                df_hj[match_column] = df_hj[cultura_column].notna()

                df_hj = df_hj[
                    df_hj["CHAVE"] != ""
                ]

                df_epi = df_epi[
                    df_epi["CHAVE"] != ""
                ]

                resultado = pd.merge(
                    df_epi,
                    df_hj,
                    how="left",
                    on="CHAVE",
                    suffixes=(
                        "",
                        "_CULTURA"
                    )
                )

                nao_encontrados = resultado[
                    ~resultado[match_column].fillna(False)
                ].copy()

                nao_encontrados[
                    "MES"
                ] = sheet_name

                if (
                    len(
                        nao_encontrados
                    ) > 0
                ):

                    self.nao_encontrados_geral.append(
                        nao_encontrados
                    )

                encontrados = (
                    resultado[
                        match_column
                    ]
                    .fillna(False)
                    .sum()
                )

                self.logs.append(

                    ExecutionLog(
                        mes=sheet_name,
                        total_epimed=len(
                            resultado
                        ),
                        encontrados=encontrados,
                        nao_encontrados=len(
                            nao_encontrados
                        )
                    )
                )

                resultado.drop(
                    columns=["CHAVE", match_column],
                    errors="ignore",
                    inplace=True
                )
                resultado.to_excel(
                    writer,
                    sheet_name=sheet_name[:31],
                    index=False
                )

                quantidade_abas += 1
                if progress_callback is not None:
                    progress_callback(
                        55 + round(40 * aba_index / len(prepared_abas))
                    )

            self._create_log_sheet(
                writer
            )

            self._create_not_found_sheet(
                writer
            )

            if quantidade_abas == 0:

                pd.DataFrame(
                    {
                        "Mensagem": [
                            "Nenhuma aba processada."
                        ]
                    }
                ).to_excel(
                    writer,
                    sheet_name="LOG"
                )

        if progress_callback is not None:
            progress_callback(100)

        return output_file

    def _create_log_sheet(
        self,
        writer
    ):

        if not self.logs:
            return

        data = []

        for item in self.logs:

            data.append(
                {
                    "Mes": item.mes,
                    "Total Epimed": item.total_epimed,
                    "Encontrados": item.encontrados,
                    "Nao Encontrados": item.nao_encontrados
                }
            )

        pd.DataFrame(
            data
        ).to_excel(
            writer,
            sheet_name="LOG",
            index=False
        )

    def _create_not_found_sheet(
        self,
        writer
    ):

        if not self.nao_encontrados_geral:
            return

        df = pd.concat(
            self.nao_encontrados_geral,
            ignore_index=True
        )

        df.to_excel(
            writer,
            sheet_name="PACIENTES_NAO_ENCONTRADOS",
            index=False
        )