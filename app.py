import math
from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(page_title="FuelCheck AI", page_icon="⛽", layout="wide")
st.title("⛽ FuelCheck AI")
st.caption("HPCL transaction analyzer — every non-zero amount below the next full rupee is included.")

def num(x):
    if pd.isna(x): return None
    try: return float(str(x).replace(",", "").replace("₹", "").strip())
    except: return None

def col(df, names):
    for c in df.columns:
        n = str(c).strip().lower().replace("_"," ").replace("-"," ")
        if n in names or any(x in n for x in names):
            return c
    return None

f = st.file_uploader("Upload HPCL Excel or CSV", type=["xlsx","xls","csv"])
if not f: st.info("Upload a transaction file to begin."); st.stop()

df = pd.read_csv(f) if f.name.lower().endswith(".csv") else pd.read_excel(f)
amount = col(df, ["amount","sale amount","transaction amount","value"])
volume = col(df, ["volume","quantity","qty","litres","liters","litre","liter"])
tid = col(df, ["transaction id","transaction no","transaction number","id","sr.no","sr no"])
dt = col(df, ["date and time","date time","datetime","transaction date time"])
date = col(df, ["date"])
time = col(df, ["time"])
nozzle = col(df, ["nozzle","nozzle no","nozzle number"])
pump = col(df, ["pump","pump no","pump number"])
product = col(df, ["product"])
mop = col(df, ["mop","mode of payment","payment"])

if not amount:
    st.error("Amount column not found.")
    st.write("Columns detected:", list(df.columns))
    st.stop()

w = df.copy()
w["_amount"] = w[amount].map(num)
w = w[w["_amount"].notna() & (w["_amount"] > 0)].copy()

# Every decimal amount below the next whole rupee is flagged.
w["_full"] = w["_amount"].apply(
    lambda x: float(math.ceil(x)) if abs(x-math.ceil(x)) > 1e-9 else None
)
flag = w[w["_full"].notna()].copy()
flag["_shortfall"] = (flag["_full"] - flag["_amount"]).round(2)
if volume: flag["_litres"] = flag[volume].map(num)

out = pd.DataFrame({"S.No.": range(1,len(flag)+1)})
out["Transaction ID"] = flag[tid].values if tid else flag.index.values
if dt: out["Date & Time"] = flag[dt].values
else:
    if date: out["Date"] = flag[date].values
    if time: out["Time"] = flag[time].values
if pump: out["Pump"] = flag[pump].values
if nozzle: out["Nozzle"] = flag[nozzle].values
if product: out["Product"] = flag[product].values
if mop: out["MOP"] = flag[mop].values
out["Actual Amount (₹)"] = flag["_amount"].round(2).values
out["Full Value (₹)"] = flag["_full"].values
out["Shortfall (₹)"] = flag["_shortfall"].values
if volume: out["Equivalent Litres"] = flag["_litres"].round(2).values

litres = float(flag["_litres"].sum()) if volume else None
c1,c2,c3,c4 = st.columns(4)
c1.metric("Transactions analysed",f"{len(w):,}")
c2.metric("Below full value",f"{len(flag):,}")
c3.metric("Total shortfall",f"₹{flag['_shortfall'].sum():,.2f}")
c4.metric("Equivalent litres",f"{litres:,.2f} L" if litres is not None else "N/A")

st.subheader("Detailed transactions")
st.dataframe(out,use_container_width=True,hide_index=True)

fv=flag.groupby("_full").agg(Transactions=("_amount","size"),Total_Shortfall=("_shortfall","sum")).reset_index()
fv=fv.rename(columns={"_full":"Full Value (₹)"})
if volume:
    lv=flag.groupby("_full")["_litres"].sum().reset_index().rename(columns={"_full":"Full Value (₹)","_litres":"Equivalent Litres"})
    fv=fv.merge(lv,on="Full Value (₹)")
st.subheader("Full-value summary")
st.dataframe(fv,use_container_width=True,hide_index=True)

if nozzle:
    ns=flag.groupby(nozzle).agg(Transactions=("_amount","size"),Total_Shortfall=("_shortfall","sum")).reset_index()
    if volume:
        nl=flag.groupby(nozzle)["_litres"].sum().reset_index().rename(columns={"_litres":"Equivalent Litres"})
        ns=ns.merge(nl,on=nozzle)
    st.subheader("Nozzle-wise summary")
    st.dataframe(ns,use_container_width=True,hide_index=True)

buf=BytesIO()
with pd.ExcelWriter(buf,engine="openpyxl") as x:
    out.to_excel(x,sheet_name="Flagged Transactions",index=False)
    fv.to_excel(x,sheet_name="Full Value Summary",index=False)
    if nozzle: ns.to_excel(x,sheet_name="Nozzle Summary",index=False)
buf.seek(0)
st.download_button("⬇️ Download Excel Report",buf.getvalue(),"FuelCheck_AI_Analysis.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("Rule: ₹219.85 → ₹220; ₹349.81 → ₹350. Exact whole-rupee transactions are excluded.")
