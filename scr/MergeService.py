from pathlib import Path
from unidecode import unidecode
import pandas as pd

from NameNormalizer import NameNormalizer
from ExecutionLog import ExecutionLog


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

    def __init__(self):

        self.normalizer = NameNormalizer()

        self.logs = []

        self.nao_encontrados_geral = []

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

        normalized = (
            unidecode(sheet_name)
            .upper()
        )

        for mes in self.MESES:

            if mes in normalized:
                return mes

        return None

    def _find_epimed_file(
        self,
        xlsx_folder: Path,
        sheet_name: str
    ):

        mes = self._extract_month(
            sheet_name
        )

        if not mes:
            return None

        arquivos = list(
            xlsx_folder.glob(
                "*.xlsx"
            )
        )

        for arquivo in arquivos:

            nome = (
                unidecode(
                    arquivo.stem
                )
                .upper()
            )

            if mes in nome:
                return arquivo

        return None

    def execute(
        self,
        cultura_folder: Path,
        epimed_xlsx_folder: Path,
        output_folder: Path
    ):

        cultura_file = (
            self._find_cultura_file(
                cultura_folder
            )
        )

        abas = pd.read_excel(
            cultura_file,
            sheet_name=None,
            header=1,
            engine="openpyxl"
        )

        output_file = (
            output_folder /
            "Epimed_Cultura.xlsx"
        )

        self.logs.clear()

        self.nao_encontrados_geral.clear()

        quantidade_abas = 0

        with pd.ExcelWriter(
            output_file,
            engine="openpyxl"
        ) as writer:

            for sheet_name, df_hj in abas.items():

                nome_aba = (
                    unidecode(
                        sheet_name
                    )
                    .upper()
                )

                if (
                    "EXPORTED"
                    in nome_aba
                ):
                    continue

                epimed_file = (
                    self._find_epimed_file(
                        epimed_xlsx_folder,
                        sheet_name
                    )
                )

                if not epimed_file:

                    self.logs.append(

                        ExecutionLog(
                            mes=sheet_name,
                            total_epimed=0,
                            encontrados=0,
                            nao_encontrados=0
                        )
                    )

                    continue

                df_epi = pd.read_excel(
                    epimed_file,
                    engine="openpyxl"
                )

                df_hj = (
                    self._clean_columns(
                        df_hj
                    )
                )

                df_epi = (
                    self._clean_columns(
                        df_epi
                    )
                )

                if (
                    "INICIAL DO PACIENTE"
                    not in df_hj.columns
                ):
                    continue

                if (
                    "Paciente"
                    not in df_epi.columns
                ):
                    continue

                df_hj = (
                    self._create_key(
                        df_hj,
                        "INICIAL DO PACIENTE"
                    )
                )

                df_epi = (
                    self._create_key(
                        df_epi,
                        "Paciente"
                    )
                )

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
                    resultado[
                        "INICIAL DO PACIENTE"
                    ].isna()
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
                        "INICIAL DO PACIENTE"
                    ]
                    .notna()
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
                    columns=["CHAVE"],
                    errors="ignore",
                    inplace=True
                )

                resultado.to_excel(
                    writer,
                    sheet_name=sheet_name[:31],
                    index=False
                )

                quantidade_abas += 1

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