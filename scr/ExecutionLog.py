from dataclasses import dataclass


@dataclass
class ExecutionLog:

    mes: str

    total_epimed: int

    encontrados: int

    nao_encontrados: int