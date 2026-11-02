"""
Build kilian_data.xlsx for the ECO 439 Kilian (2009) replication/extension.

Downloads:
  prod  : World crude oil incl. lease condensate production, thousand b/d
          (EIA International Energy Statistics, API v2; free key required)
  rac   : Refiner acquisition cost, imported, $/bbl  (EIA MER Table 9.1)
  igrea : Global real economic activity index        (FRED IGREA)
  cpi   : US CPI-U, seasonally adjusted              (FRED CPIAUCSL)

Output: kilian_data.xlsx, Sheet1, columns obsdate | prod | rac | igrea | cpi,
        monthly 1973-01 to 2025-12 (636 rows), rac blank for 1973.

Requires: pip install pandas requests openpyxl
          export EIA_API_KEY=your_key   (free: https://www.eia.gov/opendata/register.php)
"""
import io
import os
import sys
import pandas as pd
import requests

START, END = "1973-01", "2025-12"
HEADERS = {"User-Agent": "Mozilla/5.0 (ECO439 student data pull)"}


def get_csv(url):
    r = requests.get(url, headers=HEADERS, timeout=60)
    r.raise_for_status()
    return pd.read_csv(io.StringIO(r.text))


def fred(series_id, retries=3):
    """FRED series via the official API (needs FRED_API_KEY); falls back to fredgraph.csv."""
    import time
    key = os.environ.get("FRED_API_KEY")
    for attempt in range(1, retries + 1):
        try:
            if key:
                r = requests.get("https://api.stlouisfed.org/fred/series/observations",
                                 params={"series_id": series_id, "api_key": key,
                                         "file_type": "json"}, timeout=60)
                r.raise_for_status()
                df = pd.DataFrame(r.json()["observations"])[["date", "value"]]
            else:
                df = get_csv(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}")
                df.columns = ["date", "value"]
            break
        except requests.exceptions.RequestException as e:
            print(f"[FRED {series_id}] attempt {attempt} failed: {e.__class__.__name__}")
            if attempt == retries:
                raise
            time.sleep(5)
    df["date"] = pd.to_datetime(df["date"]).dt.to_period("M")
    s = pd.to_numeric(df["value"], errors="coerce")       # '.' -> NaN
    s.index = df["date"]
    print(f"[FRED] {series_id}: {s.index[0]} to {s.index[-1]}")
    return s.rename(series_id)


def eia_mer(table, must_contain):
    """Pull one monthly series from an EIA MER CSV table by matching its Description."""
    df = get_csv(f"https://www.eia.gov/totalenergy/data/browser/csv.php?tbl={table}")
    desc = df["Description"].astype(str)
    mask = pd.Series(True, index=df.index)
    for word in must_contain:
        mask &= desc.str.contains(word, case=False, regex=False)
    sel = df[mask]
    if sel["Description"].nunique() != 1:
        print(f"\n[{table}] Could not uniquely match {must_contain}. Descriptions available:")
        for d in sorted(desc.unique()):
            print("   ", d)
        sys.exit("Edit the must_contain keywords above and rerun.")
    print(f"[{table}] using: {sel['Description'].iloc[0]} ({sel['Unit'].iloc[0]})")
    sel = sel[sel["YYYYMM"].astype(str).str[-2:] != "13"]          # drop annual rows (MM = 13)
    idx = pd.PeriodIndex(pd.to_datetime(sel["YYYYMM"].astype(str), format="%Y%m"), freq="M")
    vals = pd.to_numeric(sel["Value"], errors="coerce").to_numpy()  # 'Not Available' -> NaN
    return pd.Series(vals, index=idx).sort_index()


def eia_world_crude():
    """World crude oil (incl. lease condensate) production, thousand b/d, monthly.
    EIA International Energy Statistics via API v2:
      activityId 1 = production, productId 57 = crude oil incl. lease condensate,
      countryRegionId WORL = world, unit TBPD = thousand barrels per day."""
    key = os.environ.get("EIA_API_KEY")
    if not key:
        sys.exit("Set your EIA key first:  export EIA_API_KEY=your_key")
    params = {
        "api_key": key,
        "frequency": "monthly",
        "data[0]": "value",
        "facets[activityId][]": "1",
        "facets[productId][]": "57",
        "facets[countryRegionId][]": "WORL",
        "facets[unit][]": "TBPD",
        "start": "1973-01",
        "end": END,
        "sort[0][column]": "period",
        "sort[0][direction]": "asc",
        "length": 5000,
    }
    r = requests.get("https://api.eia.gov/v2/international/data/",
                     params=params, headers=HEADERS, timeout=60)
    r.raise_for_status()
    rows = r.json().get("response", {}).get("data", [])
    if not rows:
        sys.exit("EIA API returned no rows: check the key and facet codes.")
    df = pd.DataFrame(rows)
    print(f"[EIA intl] using: {df['productName'].iloc[0]}, "
          f"{df['countryRegionName'].iloc[0]} ({df['unit'].iloc[0]}), "
          f"{df['period'].iloc[0]} to {df['period'].iloc[-1]}")
    idx = pd.PeriodIndex(df["period"], freq="M")
    return pd.Series(pd.to_numeric(df["value"], errors="coerce").to_numpy(),
                     index=idx).sort_index()


prod = eia_world_crude()
rac = eia_mer("T09.01", ["refiner acquisition cost", "imported"])
igrea = fred("IGREA")
cpi = fred("CPIAUCSL")

months = pd.period_range(START, END, freq="M")
data = pd.DataFrame({
    "obsdate": months.strftime("%Y-%m"),
    "prod": prod.reindex(months).values,
    "rac": rac.reindex(months).values,
    "igrea": igrea.reindex(months).values,
    "cpi": cpi.reindex(months).values,
})

# ---- Fill Oct-2025 CPI (not published: 2025 federal shutdown) ----------
# Log-linear interpolation of interior gaps = geometric mean of neighbours.
import numpy as np
gaps = data.loc[data["cpi"].isna(), "obsdate"].tolist()
data["cpi"] = np.exp(np.log(data["cpi"]).interpolate(limit_area="inside"))
for g in gaps:
    print(f"[CPI] interpolated {g}: {data.loc[data.obsdate == g, 'cpi'].iloc[0]:.3f}")

# ---- Checks -------------------------------------------------------------
print(f"\nRows: {len(data)} (expected 636)")
print("Missing values per column:")
print(data.isna().sum().to_string())
print("\nFirst/last observations:")
print(data.head(3).to_string(index=False))
print(data.tail(3).to_string(index=False))
print("\nExpected magnitudes: prod ~54,000-86,000 | rac ~$12 in 1974 | "
      "igrea roughly -160 to +190 | cpi ~42.6 in 1973-01")
miss_rac = data.loc[data["rac"].isna(), "obsdate"].tolist()
print(f"rac missing months: {miss_rac[:3]} ... {miss_rac[-3:] if miss_rac else ''} "
      "(should be 1973 only)")

data.to_excel("kilian_data.xlsx", sheet_name="Sheet1", index=False)
print("\nSaved kilian_data.xlsx")