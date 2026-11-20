import re
from unidecode import unidecode
import pandas as pd


class NameNormalizer:

    @staticmethod
    def normalize(value):

        if pd.isna(value):
            return ""

        value = str(value)

        value = unidecode(value)

        value = value.upper()

        value = re.sub(
            r"[^A-Z0-9 ]",
            " ",
            value
        )

        value = re.sub(
            r"\s+",
            " ",
            value
        )

        return value.strip()