import sys
import warnings
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
warnings.filterwarnings("ignore", category=UserWarning)

ARQUIVO_EXEMPLO = RAIZ / "Dataset_Financas_Complementar.xlsx"


@pytest.fixture(scope="session")
def planilha():
    """Todas as abas do arquivo de exemplo, lidas como o app lê."""
    return pd.read_excel(ARQUIVO_EXEMPLO, sheet_name=None)


@pytest.fixture(scope="session")
def gabarito(planilha):
    return planilha["Gabarito"].set_index("Nº")["Resultado"].to_dict()
