
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

SUPPLIER_SUMMARY_PATH = "Barebells Supplier Summary 4.xlsx"
WEEKLY_TRACKER_PATH   = "Barebells Weekly Item Tracker 4.xlsx"

st.set_page_config(page_title="Barebells @ Walmart – COG Dashboard", layout="wide")

@st.cache_data(show_spinner=False)
def load_topline():
    xls = pd.ExcelFile(SUPPLIER_SUMMARY_PATH)
    df = xls.parse("Topline")
    df = df.rename(columns=lambda x: str(x).strip())
    def g(i, j):
        try:
            return pd.to_numeric(df.iloc[i, j], errors="coerce")
        except Exception:
            return np.nan
    return {
        "wk_sales_ty":  g(1, 2),
        "wk_sales_ly":  g(1, 3),
        "pw_sales_ty":  g(1, 5),
        "l4w_sales_ty": g(1, 7),
        "ytd_sales_ty": g(1, 9),
        "ytd_sales_ly": g(1,10),
        "wk_units_ty":  g(3, 2),
        "wk_units_ly":  g(3, 3),
        "wk_avg_price_ty": g(8, 2),
        "wk_avg_price_ly": g(8, 3),
        "wk_instock_ty": g(7, 2),
        "wk_instock_ly": g(7, 3),
    }

@st.cache_data(show_spinner=False)
def load_item_recap():
    df = pd.read_excel(WEEKLY_TRACKER_PATH, sheet_name="Item Recap", header=2)
    df = df.rename(columns=lambda x: str(x).strip())
    if {"Prime Item Nbr","Prime Item Desc"}.issubset(df.columns):
        df[["Prime Item Nbr","Prime Item Desc"]] = df[["Prime Item Nbr","Prime Item Desc"]].ffill()
    return df

@st.cache_data(show_spinner=False)
def load_weekly_agg():
    df = pd.read_excel(WEEKLY_TRACKER_PATH, sheet_name="Item Details", header=10)
    df = df.rename(columns=lambda x: str(x).strip())
    if "Unnamed: 0" in df.columns:
        df = df.rename(columns={"Unnamed: 0":"WM Week"})
    if "WM Week" in df.columns:
        df = df[pd.to_numeric(df["WM Week"], errors="coerce").notna()].copy()
        df["WM Week"] = df["WM Week"].astype(int)
    return df

def find_indicator_col(df: pd.DataFrame):
    for col in df.columns:
        try:
            s = pd.Series(df[col]).astype("string").str.strip().str.upper()
            if s.isin(["UNITS", "SALES"]).any():
                return col
        except Exception:
            continue
    return None

@st.cache_data(show_spinner=False)
def build_sku_tidy(df: pd.DataFrame):
    df = df.copy()
    if {"Prime Item Nbr","Prime Item Desc"}.issubset(df.columns):
        df[["Prime Item Nbr","Prime Item Desc"]] = df[["Prime Item Nbr","Prime Item Desc"]].ffill()
    if "Prime Item Nbr" in df.columns:
        df = df[pd.to_numeric(df["Prime Item Nbr"], errors="coerce").notna()].copy()
    indicator_col = find_indicator_col(df)
    lwk_series = df["LWK"].astype(str).str.strip() if "LWK" in df.columns else pd.Series([""]*len(df), index=df.index)
    pspw = df[lwk_series == "$PSPW"].copy()
    if not pspw.empty:
        pspw = pspw.rename(columns={
            "Traited Stores": "$PSPW (TY)",
            "POS Stores": "$PSPW (LY)"
        })
        pspw = pspw[["Prime Item Nbr","Prime Item Desc","$PSPW (TY)","$PSPW (LY)"]]
    else:
        pspw = pd.DataFrame(columns=["Prime Item Nbr","Prime Item Desc","$PSPW (TY)","$PSPW (LY)"])
    if indicator_col is not None:
        mark = pd.Series(df[indicator_col]).astype("string").str.strip().str.upper()
        units = df[mark == "UNITS"].copy()
        sales = df[mark == "SALES"].copy()
        if not units.empty:
            units = units.rename(columns={
                "LWK": "Units (Wk28)",
                "Prev WK": "Units (Prev)",
                "Traited Stores":"Traited Stores",
                "POS Stores":"POS Stores",
                "% Stores Selling LW":"% Stores Selling",
                "Avg Retail":"Avg Retail",
                "RL Prime Unit Retail":"RL Retail"
            })[["Prime Item Nbr","Prime Item Desc","Units (Wk28)","Units (Prev)",
                "Traited Stores","POS Stores","% Stores Selling","Avg Retail","RL Retail"]]
        else:
            units = pd.DataFrame(columns=["Prime Item Nbr","Prime Item Desc","Units (Wk28)","Units (Prev)",
                                          "Traited Stores","POS Stores","% Stores Selling","Avg Retail","RL Retail"])
        if not sales.empty:
            sales = sales.rename(columns={
                "LWK":"POS$ (Wk28)",
                "Prev WK":"POS$ (Prev)",
                "YTD":"POS$ (YTD)"
            })[["Prime Item Nbr","Prime Item Desc","POS$ (Wk28)","POS$ (Prev)","POS$ (YTD)"]]
        else:
            sales = pd.DataFrame(columns=["Prime Item Nbr","Prime Item Desc","POS$ (Wk28)","POS$ (Prev)","POS$ (YTD)"])
    else:
        df["_idx"] = np.arange(len(df))
        anchors = df[lwk_series == "$PSPW"][["_idx","Prime Item Nbr"]].copy()
        rows = []
        for _, r in anchors.iterrows():
            sku_id = r["Prime Item Nbr"]
            i = int(r["_idx"])
            if i-2 >= 0 and i-1 >= 0:
                units_row = df.iloc[i-2]
                sales_row = df.iloc[i-1]
                if (units_row.get("Prime Item Nbr") == sku_id) and (sales_row.get("Prime Item Nbr") == sku_id):
                    rows.append({
                        "Prime Item Nbr": sku_id,
                        "Prime Item Desc": sales_row.get("Prime Item Desc"),
                        "Units (Wk28)": pd.to_numeric(units_row.get("LWK"), errors="coerce"),
                        "Units (Prev)": pd.to_numeric(units_row.get("Prev WK"), errors="coerce"),
                        "Traited Stores": pd.to_numeric(units_row.get("Traited Stores"), errors="coerce"),
                        "POS Stores": pd.to_numeric(units_row.get("POS Stores"), errors="coerce"),
                        "% Stores Selling": pd.to_numeric(units_row.get("% Stores Selling LW"), errors="coerce"),
                        "Avg Retail": pd.to_numeric(units_row.get("Avg Retail"), errors="coerce"),
                        "RL Retail": pd.to_numeric(units_row.get("RL Prime Unit Retail"), errors="coerce"),
                        "POS$ (Wk28)": pd.to_numeric(sales_row.get("LWK"), errors="coerce"),
                        "POS$ (Prev)": pd.to_numeric(sales_row.get("Prev WK"), errors="coerce"),
                        "POS$ (YTD)": pd.to_numeric(sales_row.get("YTD"), errors="coerce"),
                    })
        rec_df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=[
            "Prime Item Nbr","Prime Item Desc","Units (Wk28)","Units (Prev)","Traited Stores",
            "POS Stores","% Stores Selling","Avg Retail","RL Retail","POS$ (Wk28)","POS$ (Prev)","POS$ (YTD)"
        ])
        units = rec_df[["Prime Item Nbr","Prime Item Desc","Units (Wk28)","Units (Prev)",
                        "Traited Stores","POS Stores","% Stores Selling","Avg Retail","RL Retail"]]
        sales = rec_df[["Prime Item Nbr","Prime Item Desc","POS$ (Wk28)","POS$ (Prev)","POS$ (YTD)"]]
    sku = pd.merge(sales, units, on=["Prime Item Nbr","Prime Item Desc"], how="outer")
    sku = pd.merge(sku, pspw, on=["Prime Item Nbr","Prime Item Desc"], how="left")
    for c in ["POS$ (Wk28)","POS$ (Prev)","POS$ (YTD)","Units (Wk28)","Units (Prev)",
              "Traited Stores","POS Stores","% Stores Selling","Avg Retail","RL Retail",
              "$PSPW (TY)","$PSPW (LY)"]:
        if c in sku.columns:
            sku[c] = pd.to_numeric(sku[c], errors="coerce")
    if {"POS Stores","Traited Stores"}.issubset(sku.columns):
        with np.errstate(divide='ignore', invalid='ignore'):
            sku["% Stores Selling (calc)"] = sku["POS Stores"] / sku["Traited Stores"]
        sku["% Stores Selling"] = sku["% Stores Selling"].fillna(sku["% Stores Selling (calc)"])
    sku["WoW $ Δ"] = sku["POS$ (Wk28)"] - sku["POS$ (Prev)"]
    sku["WoW Units Δ"] = sku["Units (Wk28)"] - sku["Units (Prev)"]
    sku = sku.dropna(subset=["Prime Item Nbr","Prime Item Desc"], how="any")
    return sku

# Load data
topline = load_topline()
recap = load_item_recap()
weekly = load_weekly_agg()
sku = build_sku_tidy(recap)

st.title("Barebells @ Walmart — COG Dashboard")
tab_overview, tab_rank, tab_mix, tab_movers, tab_insights, tab_trends = st.tabs(
    ["Overview", "SKU Rankings", "SKU Mix", "Top Movers", "Insights", "Trends"]
)

with tab_overview:
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Sales $ (Wk28, TY)", f"${topline['wk_sales_ty']:,.0f}",
              f"{(topline['wk_sales_ty']-topline['pw_sales_ty'])/topline['pw_sales_ty']:.1%} vs Wk27" if topline['pw_sales_ty'] else None)
    c2.metric("Units (Wk28, TY)", f"{int(topline['wk_units_ty']):,}",
              f"{(topline['wk_units_ty']-topline['wk_units_ly'])/topline['wk_units_ly']:.1%} YoY" if topline['wk_units_ly'] else None)
    c3.metric("Avg Price (Wk28, TY)", f"${topline['wk_avg_price_ty']:.2f}",
              f"{(topline['wk_avg_price_ty']-topline['wk_avg_price_ly'])/topline['wk_avg_price_ly']:.1%} YoY" if topline['wk_avg_price_ly'] else None)
    c4.metric("Instock % (Wk28, TY)", f"{topline['wk_instock_ty']:.1%}",
              f"{(topline['wk_instock_ty']-topline['wk_instock_ly']):.1%} YoY" if topline['wk_instock_ly'] else None)
    comp = pd.DataFrame({
        "Metric":["Prev Week (27)","Week 28","Last 4 Weeks"],
        "POS $":[topline["pw_sales_ty"], topline["wk_sales_ty"], topline["l4w_sales_ty"]],
        "Instock %":[np.nan, topline["wk_instock_ty"], np.nan]
    })
    colA, colB = st.columns(2)
    colA.plotly_chart(px.bar(comp, x="Metric", y="POS $", title="POS $ Comparison"), use_container_width=True)
    colB.plotly_chart(px.bar(comp, x="Metric", y="Instock %", title="Instock % Comparison"), use_container_width=True)

with tab_rank:
    st.subheader("SKU Rankings — Week 28")
    rank_mode = st.radio("Rank by:", ["POS $ (Wk28)", "Units (Wk28)"], horizontal=True, key="rankmode")
    rank_df = sku.copy()
    metric_col = "POS$ (Wk28)" if rank_mode.startswith("POS") else "Units (Wk28)"
    rank_df = rank_df.sort_values(metric_col, ascending=False)
    st.dataframe(rank_df[["Prime Item Nbr","Prime Item Desc","POS$ (Wk28)","Units (Wk28)","POS$ (Prev)","Units (Prev)"]],
                 use_container_width=True, height=520)
    st.download_button("Download Current Ranking (CSV)",
                       rank_df.to_csv(index=False).encode("utf-8"),
                       "sku_ranking_wk28.csv","text/csv")

with tab_mix:
    st.subheader("SKU Mix — Week 28 (POS $ Share)")
    top_n = st.slider("Show top N SKUs (rest grouped as 'Others')", 3, 15, 8, key="topn")
    mix = sku[["Prime Item Desc","POS$ (Wk28)"]].dropna().sort_values("POS$ (Wk28)", ascending=False)
    top = mix.head(top_n).copy()
    others_val = mix["POS$ (Wk28)"].iloc[top_n:].sum()
    if others_val > 0:
        top = pd.concat([top, pd.DataFrame({"Prime Item Desc":["Others"], "POS$ (Wk28)":[others_val]})], ignore_index=True)
    fig_mix = px.pie(top, names="Prime Item Desc", values="POS$ (Wk28)", hole=0.5, title="Week 28 POS $ Mix")
    fig_mix.update_traces(textposition="inside", textinfo="percent+label")
    st.plotly_chart(fig_mix, use_container_width=True)

with tab_movers:
    st.subheader("Top Movers — WoW (POS $ and Units)")
    movers = sku.copy()
    movers_d = movers.dropna(subset=["WoW $ Δ"]).sort_values("WoW $ Δ", ascending=False)
    moversu = movers.dropna(subset=["WoW Units Δ"]).sort_values("WoW Units Δ", ascending=False)
    c1,c2 = st.columns(2)
    with c1:
        st.markdown("**Top $ Gainers**")
        st.dataframe(movers_d.head(10)[["Prime Item Desc","POS$ (Wk28)","POS$ (Prev)","WoW $ Δ"]], use_container_width=True)
        st.markdown("**Top $ Decliners**")
        st.dataframe(movers_d.tail(10).sort_values("WoW $ Δ")[["Prime Item Desc","POS$ (Wk28)","POS$ (Prev)","WoW $ Δ"]], use_container_width=True)
    with c2:
        st.markdown("**Top Unit Gainers**")
        st.dataframe(moversu.head(10)[["Prime Item Desc","Units (Wk28)","Units (Prev)","WoW Units Δ"]], use_container_width=True)
        st.markdown("**Top Unit Decliners**")
        st.dataframe(moversu.tail(10).sort_values("WoW Units Δ")[["Prime Item Desc","Units (Wk28)","Units (Prev)","WoW Units Δ"]], use_container_width=True)

with tab_insights:
    st.subheader("Insights")
    vdf = sku.dropna(subset=["$PSPW (TY)"]).copy()
    pct_col = "% Stores Selling" if "% Stores Selling" in vdf.columns else ("% Stores Selling (calc)" if "% Stores Selling (calc)" in vdf.columns else None)
    if pct_col is not None and not vdf.empty:
        fig1 = px.scatter(vdf, x="$PSPW (TY)", y=pct_col,
                          size="POS$ (Wk28)", hover_name="Prime Item Desc",
                          labels={"$PSPW (TY)":"$ per Store per Week (TY)", pct_col:"% Stores Selling"},
                          title="Velocity vs Distribution (Bubble size = POS$ Wk28)")
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.info("Not enough data to plot Velocity vs Distribution.")
    pdf = sku.dropna(subset=["Avg Retail","$PSPW (TY)"]).copy()
    if not pdf.empty:
        fig2 = px.scatter(pdf, x="Avg Retail", y="$PSPW (TY)", hover_name="Prime Item Desc",
                          trendline="ols", title="Avg Retail vs $PSPW (TY)")
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Not enough data to plot Avg Retail vs $PSPW.")

with tab_trends:
    st.subheader("Topline Trends (Weekly Aggregate)")
    row1, row2 = st.columns(2)
    if not weekly.empty and "POS $" in weekly.columns:
        row1.plotly_chart(px.line(weekly, x="WM Week", y="POS $", title="POS $ by WM Week"), use_container_width=True)
    if not weekly.empty and "Instock %" in weekly.columns:
        row2.plotly_chart(px.line(weekly, x="WM Week", y="Instock %", title="Instock % by WM Week"), use_container_width=True)

st.caption("Use the rankings toggle to switch between $ and unit leaders. Download tables via the buttons above.")
