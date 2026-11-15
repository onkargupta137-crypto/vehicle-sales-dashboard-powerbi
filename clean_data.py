"""Step 1 - Clean the raw vehicle sales dataset.
Input : vehicle_dataset.xlsx   Output: vehicle_sales_clean.csv
"""
import datetime as dt
import pandas as pd

RAW = "vehicle_dataset.xlsx"
OUT = "vehicle_sales_clean.csv"

# keep_default_na=False -> otherwise the territory code "NA" (North America) is read as missing
df = pd.read_excel(RAW, keep_default_na=False, na_values=[""])

def parse_date(v):
    """Excel auto-converted some text dates with day and month swapped.
    - real datetime cells  -> swap day/month back
    - text 'dd-mm-yyyy ...' -> parse normally
    """
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return pd.Timestamp(v.year, v.day, v.month, v.hour, v.minute, v.second)
    return pd.to_datetime(v, format="%d-%m-%Y %H:%M:%S")

df["ORDER_DATE"] = df["ORDERDATETIME"].map(parse_date)
df["DELIVERY_DATE"] = df["DELIVERYDATETIME"].map(parse_date)

# Validation: a delivery can never happen before the order
assert (df["DELIVERY_DATE"] >= df["ORDER_DATE"]).all(), "Date fix failed"

df["SALES"] = df["QUANTITYORDERED"] * df["PRICEEACH"]
df["ORDER_YEAR"] = df["ORDER_DATE"].dt.year
df["DELIVERY_DAYS"] = (df["DELIVERY_DATE"] - df["ORDER_DATE"]).dt.total_seconds() / 86400

df["TERRITORY"] = df["TERRITORY"].replace({"NA": "North America"})
df["STATE"] = df["STATE"].fillna("Not applicable")
df["POSTALCODE"] = df["POSTALCODE"].fillna("Unknown")

# Privacy: drop customer names before publishing on GitHub
df = df.drop(columns=["CONTACTLASTNAME", "CONTACTFIRSTNAME", "ORDERDATETIME", "DELIVERYDATETIME"])
df.to_csv(OUT, index=False)
print(df.shape, "rows/cols saved to", OUT)
