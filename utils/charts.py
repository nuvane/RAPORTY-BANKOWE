import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO

def create_trend_chart(df, title, start_date=None, unit='mln'):
    df = df.copy()
    if start_date:
        df = df[df["Data"] >= start_date]
    if df.empty:
        return None

    df = df.sort_values("Data")
    x = df["Data"]
    y = df["Saldo"].values / (1_000_000 if unit == 'mln' else 1_000)

    z = np.polyfit(range(len(x)), y, 1)
    p = np.poly1d(z)
    trend = p(range(len(x)))

    plt.figure(figsize=(8, 4))
    plt.plot(x, y, label='Saldo')
    plt.plot(x, trend, '--', label='Trend')
    plt.title(title)
    plt.xlabel("Data")
    plt.ylabel(f"Saldo ({unit})")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.grid(True)
    plt.legend()

    buf = BytesIO()
    plt.savefig(buf, format='png')
    plt.close()
    buf.seek(0)
    return buf

def create_summary_chart(df, title, start_date, unit='mln'):
    df = df[df["Data"] >= start_date]
    if df.empty:
        return None

    divider = 1_000_000 if unit == 'mln' else 1_000

    plt.figure(figsize=(10, 5))
    for firma in df["Firma"].unique():
        df_firma = df[df["Firma"] == firma]
        plt.plot(df_firma["Data"], df_firma["Saldo"] / divider, label=firma)

    plt.title(title)
    plt.xlabel("Data")
    plt.ylabel(f"Saldo ({unit})")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.grid(True)
    plt.legend()

    buf = BytesIO()
    plt.savefig(buf, format='png')
    plt.close()
    buf.seek(0)
    return buf
