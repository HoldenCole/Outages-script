#!/usr/bin/env python3
"""
fetch_market_data.py
Pulls a monthly **gasoline crack spread** ($/bbl) from the EIA open data feed
(no API key required) and vendors it to `data/market_crack.csv` so the workbook's
Margin Context sheet can value offline capacity in dollar terms.

    crack ($/bbl) = NY-Harbor conventional regular gasoline spot ($/gal) x 42
                    - WTI crude spot ($/bbl)

This is the "refresh the feed" pattern for market data: re-run it to update the
CSV, or just overwrite the CSV with a Bloomberg pull (e.g. RBOB crack 321) using
the same columns. The engine reads the CSV and degrades gracefully if it's absent.

    python scripts/fetch_market_data.py            # -> data/market_crack.csv

Source: U.S. Energy Information Administration (EIA), monthly spot prices.
  WTI       https://www.eia.gov/dnav/pet/hist_xls/RWTCm.xls
  Gasoline  https://www.eia.gov/dnav/pet/hist_xls/EER_EPMRU_PF4_Y35NY_DPGm.xls
"""
import sys
import urllib.request
from pathlib import Path

import pandas as pd

_ROOT = Path(__file__).resolve().parent.parent
OUT = _ROOT / "data" / "market_crack.csv"
WTI_URL = "https://www.eia.gov/dnav/pet/hist_xls/RWTCm.xls"
GAS_URL = "https://www.eia.gov/dnav/pet/hist_xls/EER_EPMRU_PF4_Y35NY_DPGm.xls"
GAL_PER_BBL = 42.0


def _load_eia(url, name):
    """Download an EIA hist_xls series and return a tidy Date|value frame."""
    tmp = Path("/tmp") / Path(url).name
    urllib.request.urlretrieve(url, tmp)
    df = pd.read_excel(tmp, sheet_name="Data 1", skiprows=2)
    df.columns = ["Date", name]
    df["Date"] = pd.to_datetime(df["Date"])
    return df.dropna()


def main():
    wti = _load_eia(WTI_URL, "wti")
    gas = _load_eia(GAS_URL, "gasoline")          # $/gal
    m = pd.merge(wti, gas, on="Date").sort_values("Date")
    m["crack"] = (m["gasoline"] * GAL_PER_BBL - m["wti"]).round(3)
    m["year"] = m["Date"].dt.year
    m["month"] = m["Date"].dt.month
    out = m[["year", "month", "crack", "gasoline", "wti"]].copy()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("# Gasoline crack spread ($/bbl) = NYH conventional regular gasoline "
                "($/gal) x 42 - WTI ($/bbl).  Source: EIA monthly spot prices.\n")
        f.write("# Overwrite with a Bloomberg pull (same columns) to go live on RBOB 321, etc.\n")
        out.to_csv(f, index=False)
    print(f"Wrote {OUT}  ({len(out)} months, {out['year'].min()}-{out['year'].max()})")
    print(out[out["year"] >= 2024].to_string(index=False))


if __name__ == "__main__":
    sys.exit(main())
