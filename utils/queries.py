import pandas as pd
import numpy as np

def generate_sample_data(firma: str, waluta: str, start_date, end_date) -> pd.DataFrame:
    dates = pd.date_range(start=start_date, end=end_date, freq='D')
    saldo = np.cumsum(np.random.randn(len(dates)) * 10000 + (5000 if waluta == 'PLN' else 1000))
    df = pd.DataFrame({
        "Data": dates,
        "Saldo": saldo,
        "Firma": firma,
        "Waluta": waluta
    })
    return df
