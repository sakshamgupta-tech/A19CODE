import streamlit as st
import pandas as pd
import numpy as np
import re
from collections import deque
import os
import shutil
import traceback
from datetime import datetime, date
import tempfile
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Algo19 Utils", layout="wide")

st.title("Algo19 Utils")

st.write("Welcome to Algo19 Utils - App loaded successfully!")

# ────────────────────────────────────────────────
#  Session state keys prefixed for Tab 2
# ────────────────────────────────────────────────
if 'tab2_processed' not in st.session_state:
    st.session_state.tab2_processed = False
if 'tab2_pnl_bytes' not in st.session_state:
    st.session_state.tab2_pnl_bytes = None
if 'tab2_leg_bytes' not in st.session_state:
    st.session_state.tab2_leg_bytes = None
if 'tab2_hedge_bytes' not in st.session_state:
    st.session_state.tab2_hedge_bytes = None
if 'tab2_user_hedge_bytes' not in st.session_state:
    st.session_state.tab2_user_hedge_bytes = None
if 'tab2_per_date_hedge' not in st.session_state:
    st.session_state.tab2_per_date_hedge = None
if 'tab2_hedge_exchange' not in st.session_state:
    st.session_state.tab2_hedge_exchange = None

# ────────────────────────────────────────────────
#   Jainam Calculation for Algo19
# ────────────────────────────────────────────────

st.header("Jainam Calculation for Algo19")
st.caption("Jainam Realized & Settlement Excel + SUMMARY.xlsx → Segment PNL − NF/SN PNL = Hedge Cost")
st.info("Upload (1) Jainam Realized & Settlement Excel  (2) VS20 SUMMARY.xlsx → pick segment & date range → Calculate")

col1, col2 = st.columns([5, 5])

with col1:
    jainam_rs_file = st.file_uploader(
        "Upload Jainam Realized & Settlement Excel (.xlsx)",
        type=["xlsx", "xls"],
        key="jainam_rs_uploader"
    )

with col2:
    excel_file_tab2 = st.file_uploader(
        "Upload VS20 ... SUMMARY.xlsx",
        type=["xlsx", "xls"],
        key="excel_uploader"
    )

col3, col4, col5 = st.columns(3)
with col3:
    start_date_tab2 = st.date_input("Start Date", key="start_date")
with col4:
    end_date_tab2   = st.date_input("End Date",   key="end_date")
with col5:
    exchange_tab2 = st.radio(
        "Segment",
        options=["NFO", "BFO"],
        horizontal=True,
        key="exchange_radio"
    )

if st.button("Calculate Hedge Cost", type="primary", use_container_width=True, key="process_btn"):
    if not jainam_rs_file:
        st.error("Please upload the Jainam Realized & Settlement Excel file.")
    elif not excel_file_tab2:
        st.error("Please upload the SUMMARY.xlsx file.")
    else:
        with st.spinner("Calculating hedge cost..."):
            try:
                s_date = pd.to_datetime(start_date_tab2)
                e_date = pd.to_datetime(end_date_tab2)

                # ───── SUMMARY.xlsx: compute NF/SN PNL from MultiLeg Orders ─────
                df = pd.read_excel(excel_file_tab2, sheet_name="MultiLeg Orders")
                df = df[df["Status"] == "COMPLETE"]

                underlying_filter = "NF" if exchange_tab2 == "NFO" else "SN"
                
                leg_results = []
                for (user, port), g in df.groupby(["User ID", "Portfolio Name"]):
                    for leg, leg_data in g.groupby("Leg ID"):
                        leg_data = leg_data.copy()
                        
                        leg_data["Order Time"] = pd.to_datetime(leg_data["Order Time"], errors="coerce")
                        entry_time = leg_data["Order Time"].min()
                        exit_time = leg_data["Order Time"].max()

                        entry_time = entry_time.date() if pd.notna(entry_time) else None
                        exit_time = exit_time.date() if pd.notna(exit_time) else None
                        
                        buy_df = leg_data[leg_data["Transaction"] == "BUY"]
                        sell_df = leg_data[leg_data["Transaction"] == "SELL"]

                        buy_price = buy_df["Avg Price"].mean()
                        sell_price = sell_df["Avg Price"].mean()

                        buy_qty = buy_df["Quantity"].sum()
                        sell_qty = sell_df["Quantity"].sum()

                        buy_price = 0 if pd.isna(buy_price) else buy_price
                        sell_price = 0 if pd.isna(sell_price) else sell_price

                        qty = sell_qty if sell_qty != 0 else buy_qty
                        pnl_leg = (sell_price - buy_price) * qty
                        
                        if "_NF_" in port:
                            under = "NF"
                        elif "_SN_" in port:
                            under = "SN"
                        else:
                            under = "OTHER"

                        m = re.search(r'(\d+DTE)', port)
                        dte = m.group(1) if m else "UNKNOWN"
                        
                        leg_results.append({
                            "User ID": user,
                            "Underlying": under,
                            "DTE": dte,
                            "Leg ID": leg,
                            "Entry Time": entry_time,
                            "Exit Time": exit_time,
                            "PNL": pnl_leg
                        })

                leg_df = pd.DataFrame(leg_results)
                
                # Apply user filter based on UI
                leg_df = leg_df[leg_df["Underlying"] == underlying_filter]
                
                # DATE-WISE PNL PER USER
                datewise_pnl = leg_df.groupby(["User ID", "Entry Time"])["PNL"].sum().reset_index()
                
                # Emulate agg_df for downstream compatibility (P/L Distribution Row-wise)
                agg_df = datewise_pnl.copy()
                agg_df = agg_df.rename(columns={"Entry Time": "Date"})
                
                # Aggregate for hedge calculation total NF_SN_PNL
                nf_sn_pnl = (
                    agg_df.groupby("User ID")["PNL"].sum()
                    .reset_index()
                    .rename(columns={"PNL": "NF_SN_PNL", "User ID": "User Id"})
                )

                # ───── MAIN FILE: Jainam Realized & Settlement ─────
                # Get all months/sheets in the selected range
                months_in_range = pd.period_range(start=s_date, end=e_date, freq='M').strftime("%b%y").unique().tolist()
                
                try:
                    xl = pd.ExcelFile(jainam_rs_file)
                    available_sheets = xl.sheet_names
                    all_rs_dfs = []
                    found_any_sheet = False
                    
                    for m_name in months_in_range:
                        if m_name in available_sheets:
                            temp_df = pd.read_excel(xl, sheet_name=m_name)
                            temp_df.columns = temp_df.columns.str.strip()
                            all_rs_dfs.append(temp_df)
                            found_any_sheet = True
                    
                    if not found_any_sheet:
                        st.warning("None of the month-specific sheets found (e.g., 'Jan26') – trying the first sheet.")
                        jainam_rs_file.seek(0)
                        rs_df = pd.read_excel(jainam_rs_file)
                    else:
                        rs_df = pd.concat(all_rs_dfs, ignore_index=True)
                except Exception as e_xl:
                    st.warning(f"Error reading sheets: {e_xl}. Trying first sheet.")
                    jainam_rs_file.seek(0)
                    rs_df = pd.read_excel(jainam_rs_file)

                rs_df.columns = rs_df.columns.str.strip()

                # Robust column name finder (case-insensitive)
                col_map = {c.lower(): c for c in rs_df.columns}
                def find_col(target):
                    return col_map.get(target.lower())

                # Normalise Date column
                date_col_actual = find_col("date")
                if date_col_actual is None:
                    st.error(f"No 'Date' column found in the Excel data. Available: {list(rs_df.columns)}")
                    st.stop()
                if date_col_actual != "Date":
                    rs_df = rs_df.rename(columns={date_col_actual: "Date"})

                rs_df["Date"] = pd.to_datetime(rs_df["Date"], errors="coerce")
                rs_df = rs_df[(rs_df["Date"] >= s_date) & (rs_df["Date"] <= e_date)]

                if rs_df.empty:
                    st.warning(f"No data found for the selected date range: {s_date.date()} to {e_date.date()}.")
                    st.stop()

                # Rebuild col_map after rename
                col_map = {c.lower(): c for c in rs_df.columns}

                # Segment_Result = Realized + Settlement Value
                if exchange_tab2 == "NFO":
                    nifty_r = find_col("nifty realized")  or find_col("nifty_realized")
                    nifty_s = find_col("nifty settlement value") or find_col("nifty_settlement value")
                    if not nifty_r or not nifty_s:
                        st.error(f"NFO columns not found. Available: {list(rs_df.columns)}"); st.stop()
                    rs_df["Segment_Result"] = rs_df[nifty_r] + rs_df[nifty_s]
                else:  # BFO
                    sensex_r = find_col("sensex realized") or find_col("sensex_realized")
                    sensex_s = find_col("sensex settlement value") or find_col("sensex_settlement value")
                    if not sensex_r or not sensex_s:
                        st.error(f"BFO columns not found. Available: {list(rs_df.columns)}"); st.stop()
                    rs_df["Segment_Result"] = rs_df[sensex_r] + rs_df[sensex_s]

                # User-wise sum from RS file
                user_col = (find_col("user id") or find_col("user_id") or
                            find_col("userid") or list(rs_df.columns)[1])
                result = rs_df.groupby(user_col)["Segment_Result"].sum().reset_index()
                result.columns = ["User Id", "Segment_Result"]

                # ───── MERGE: Segment_Result - NF_SN_PNL = Final_Result ─────
                result = result.merge(nf_sn_pnl, on="User Id", how="left")
                result["NF_SN_PNL"]    = result["NF_SN_PNL"].fillna(0)
                result["Final_Result"] = result["Segment_Result"] - result["NF_SN_PNL"]
                result["Segment"]      = exchange_tab2

                # Display
                st.subheader(f"Result ({exchange_tab2}) — {s_date.date()} to {e_date.date()}")
                st.dataframe(result, use_container_width=True)

                # Store for distribution (Final_PNL key for compatibility)
                user_summary = result.rename(columns={"User Id": "User ID", "Final_Result": "Final_PNL"})
                st.session_state.tab2_user_summary = user_summary.copy()

                # Per-date for hedge distribution
                per_date_df = rs_df[[user_col, "Date", "Segment_Result"]].copy()
                per_date_df = per_date_df.rename(columns={user_col: "User ID", "Segment_Result": "Hedge_PNL"})
                per_date_df["dt"]        = per_date_df["Date"]
                per_date_df["File_Date"]  = per_date_df["Date"].dt.strftime("%d %b %Y")
                st.session_state.tab2_per_date_hedge = per_date_df
                st.session_state.tab2_hedge_exchange = exchange_tab2

                # Save downloads
                user_buf = BytesIO()
                result.to_csv(user_buf, index=False)
                user_buf.seek(0)
                st.session_state.tab2_user_hedge_bytes = user_buf.getvalue()

                hedge_buf = BytesIO()
                rs_df.to_csv(hedge_buf, index=False)
                hedge_buf.seek(0)
                st.session_state.tab2_hedge_bytes = hedge_buf.getvalue()

                # Save agg_df for row-wise distribution
                st.session_state.tab2_agg_df = agg_df.copy()

                # NF_SN_PNL file (agg_df from SUMMARY.xlsx)
                pnl_buf = BytesIO()
                with pd.ExcelWriter(pnl_buf, engine='openpyxl') as w:
                    agg_df.to_excel(w, index=False, sheet_name='NF_SN_PNL')
                pnl_buf.seek(0)
                st.session_state.tab2_pnl_bytes = pnl_buf.getvalue()

                # Leg-wise PNL file
                leg_buf = BytesIO()
                with pd.ExcelWriter(leg_buf, engine='openpyxl') as w:
                    leg_df.to_excel(w, index=False, sheet_name='Leg_PNL')
                leg_buf.seek(0)
                st.session_state.tab2_leg_bytes = leg_buf.getvalue()

                st.session_state.tab2_processed = True
                st.success("Calculation complete!")

            except Exception as e:
                st.error(f"Error: {str(e)}")
                with st.expander("Show error details"):
                    st.exception(e)

# ────────────────────────────────────────────────
#  Download section – always visible after successful processing
# ────────────────────────────────────────────────
if st.session_state.tab2_processed:
    st.markdown("### Download Results")
    col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)

    with col_dl1:
        if st.session_state.tab2_pnl_bytes:
            st.download_button(
                label="📥 NF_SN_PNL.xlsx",
                data=st.session_state.tab2_pnl_bytes,
                file_name="NF_SN_PNL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_pnl"
            )

    with col_dl2:
        if st.session_state.tab2_hedge_bytes:
            st.download_button(
                label="📥 Jainam RS Filtered Data (.csv)",
                data=st.session_state.tab2_hedge_bytes,
                file_name="jainam_rs_filtered.csv",
                mime="text/csv",
                key="dl_hedge"
            )

    with col_dl3:
        if st.session_state.tab2_user_hedge_bytes:
            st.download_button(
                label="📥 User-wise Final Result (.csv)",
                data=st.session_state.tab2_user_hedge_bytes,
                file_name="user_wise_hedge_result.csv",
                mime="text/csv",
                key="dl_user_hedge"
            )

    with col_dl4:
        if st.session_state.get('tab2_leg_bytes'):
            st.download_button(
                label="📥 Leg_PNL.xlsx",
                data=st.session_state.tab2_leg_bytes,
                file_name="Leg_PNL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_leg_pnl"
            )

    st.info("Files are kept in memory until you refresh or restart the app.")

# ──────────────────────────────────────────────────────────────────
#  Hedge Cost Distribution by Jainam Daily Allocation
# ──────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Hedge Cost Distribution")
st.caption("Upload 'Jainam Daily Allocation.xlsx' → parse Allocation Record sheet → Divide Hedge Cost based on VT, GB, PS, RD, RM ratios per date")

with st.form("hedge_dist_form"):
    alloc_file = st.file_uploader("Upload Jainam Daily Allocation.xlsx", type=["xlsx"], key="alloc_hedge_upl")
    submit_hedge_dist = st.form_submit_button("Process Hedge Cost Distribution", type="primary")

if submit_hedge_dist:
    if st.session_state.get('tab2_per_date_hedge') is None or st.session_state.get('tab2_user_summary') is None:
        st.error("Please process the Jainam files first (click 'Calculate Hedge Cost' above).")
    elif not alloc_file:
        st.error("Please upload the Jainam Daily Allocation file.")
    else:
        with st.spinner("Calculating Hedge Cost Distribution..."):
            try:
                df_raw_alloc = pd.read_excel(alloc_file, sheet_name='Record', header=None)
                all_allocs = []
                curr_date_str = None
                for i in range(len(df_raw_alloc)):
                    cell = df_raw_alloc.iloc[i, 0]
                    dt_val = None
                    if isinstance(cell, (datetime, pd.Timestamp)):
                        dt_val = cell
                        if dt_val.year == 2026 and dt_val.day == 2 and 2 <= dt_val.month <= 12:
                            dt_val = datetime(2026, 2, dt_val.month)
                    elif isinstance(cell, str):
                        m = re.match(r'(\d{2}-\d{2}-\d{4})', cell)
                        if m:
                            try: dt_val = datetime.strptime(m.group(1), '%d-%m-%Y')
                            except: pass
                    if dt_val:
                        curr_date_str = dt_val.strftime('%d-%m-%Y')
                        continue
                    uid = df_raw_alloc.iloc[i, 0]
                    if curr_date_str and pd.notna(uid) and uid != 'UserID' and str(uid).strip() != '':
                        try:
                            vt = float(df_raw_alloc.iloc[i, 4]) if pd.notna(df_raw_alloc.iloc[i, 4]) else 0
                            gb = float(df_raw_alloc.iloc[i, 5]) if pd.notna(df_raw_alloc.iloc[i, 5]) else 0
                            ps = float(df_raw_alloc.iloc[i, 6]) if pd.notna(df_raw_alloc.iloc[i, 6]) else 0
                            rd = float(df_raw_alloc.iloc[i, 7]) if pd.notna(df_raw_alloc.iloc[i, 7]) else 0
                            rm = float(df_raw_alloc.iloc[i, 8]) if pd.notna(df_raw_alloc.iloc[i, 8]) else 0
                            all_allocs.append({
                                'Date_Alloc': curr_date_str, 'UserID': str(uid).strip(),
                                'VT': vt, 'GB': gb, 'PS': ps, 'RD': rd, 'RM': rm,
                                'dt': pd.to_datetime(curr_date_str, dayfirst=True)
                            })
                        except: continue
                df_alloc = pd.DataFrame(all_allocs).sort_values('dt')
                pd_hedge = st.session_state.tab2_per_date_hedge.copy()
                usr_summary = st.session_state.tab2_user_summary.copy()
                hedge_dist_rows = []
                for _, urow in usr_summary.iterrows():
                    u_id = urow["User ID"]
                    h_pnl = urow["Final_PNL"]
                    user_dates = pd_hedge[pd_hedge["User ID"] == u_id][["File_Date", "dt"]].drop_duplicates()
                    vt_sum = gb_sum = ps_sum = rd_sum = rm_sum = 0.0
                    date_detail = []
                    u_allocs = df_alloc[df_alloc['UserID'] == u_id]
                    for _, drow in user_dates.iterrows():
                        p_dt = drow["dt"]
                        v_allocs = u_allocs[u_allocs['dt'] <= p_dt]
                        if v_allocs.empty:
                            dm = p_dt.strftime('%d-%m') if pd.notna(p_dt) else ''
                            v_allocs = u_allocs[u_allocs['dt'].dt.strftime('%d-%m') == dm]
                        best_a = v_allocs.iloc[-1] if not v_allocs.empty else None
                        if best_a is not None:
                            vt_sum += best_a['VT']; gb_sum += best_a['GB']; ps_sum += best_a['PS']
                            rd_sum += best_a['RD']; rm_sum += best_a['RM']
                            date_detail.append(f"{drow['File_Date']}:VT={best_a['VT']},GB={best_a['GB']},PS={best_a['PS']},RD={best_a['RD']},RM={best_a['RM']}")
                    gs = vt_sum + gb_sum + ps_sum + rd_sum + rm_sum
                    det_str = " | ".join(date_detail) if date_detail else "NOT FOUND"
                    hedge_dist_rows.append({
                        "User ID": u_id, "Total_Hedge_Cost": h_pnl,
                        "VT_Hedge": (vt_sum/gs)*h_pnl if gs>0 else 0, "GB_Hedge": (gb_sum/gs)*h_pnl if gs>0 else 0,
                        "PS_Hedge": (ps_sum/gs)*h_pnl if gs>0 else 0, "RD_Hedge": (rd_sum/gs)*h_pnl if gs>0 else 0,
                        "RM_Hedge": (rm_sum/gs)*h_pnl if gs>0 else 0,
                        "Date_Contributions": det_str
                    })
                df_hedge_dist = pd.DataFrame(hedge_dist_rows)
                st.session_state.df_hedge_dist = df_hedge_dist
                st.success("Hedge Cost Distribution calculated!")
                st.dataframe(df_hedge_dist.head(10))
                buf = BytesIO(); df_hedge_dist.to_excel(buf, index=False); buf.seek(0)
                st.download_button("📥 Download Hedge Cost Distribution", buf, "Hedge_Cost_Distribution.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")

# ──────────────────────────────────────────────────────────────────
#  P/L Distribution by Jainam Daily Allocation (Row-wise)
# ──────────────────────────────────────────────────────────────────
st.divider()
st.subheader("P/L Distribution (Row-wise)")
st.caption("Divide PNL of each row (from SUMMARY) based on Jainam Daily Allocation ratios for that User & Date")

with st.form("pl_dist_form"):
    alloc_file_pl = st.file_uploader("Upload Jainam Daily Allocation.xlsx", type=["xlsx"], key="alloc_pl_upl")
    submit_pl_dist = st.form_submit_button("Process P/L Distribution", type="primary")

if submit_pl_dist:
    agg_df = st.session_state.get('tab2_agg_df')
    if agg_df is None:
        st.error("Please process the Jainam files first (click 'Calculate Hedge Cost' above).")
    elif not alloc_file_pl:
        st.error("Please upload the Jainam Daily Allocation file.")
    else:
        with st.spinner("Calculating Row-wise P/L Distribution..."):
            try:
                df_raw_alloc = pd.read_excel(alloc_file_pl, sheet_name='Record', header=None)
                all_allocs = []
                curr_date_str = None
                for i in range(len(df_raw_alloc)):
                    cell = df_raw_alloc.iloc[i, 0]
                    dt_val = None
                    if isinstance(cell, (datetime, pd.Timestamp)):
                        dt_val = cell
                        if dt_val.year == 2026 and dt_val.day == 2 and 2 <= dt_val.month <= 12:
                            dt_val = datetime(2026, 2, dt_val.month)
                    elif isinstance(cell, str):
                        m = re.match(r'(\d{2}-\d{2}-\d{4})', cell)
                        if m:
                            try: dt_val = datetime.strptime(m.group(1), '%d-%m-%Y')
                            except: pass
                    if dt_val:
                        curr_date_str = dt_val.strftime('%d-%m-%Y')
                        continue
                    uid = df_raw_alloc.iloc[i, 0]
                    if curr_date_str and pd.notna(uid) and uid != 'UserID' and str(uid).strip() != '':
                        try:
                            vt = float(df_raw_alloc.iloc[i, 4]) if pd.notna(df_raw_alloc.iloc[i, 4]) else 0
                            gb = float(df_raw_alloc.iloc[i, 5]) if pd.notna(df_raw_alloc.iloc[i, 5]) else 0
                            ps = float(df_raw_alloc.iloc[i, 6]) if pd.notna(df_raw_alloc.iloc[i, 6]) else 0
                            rd = float(df_raw_alloc.iloc[i, 7]) if pd.notna(df_raw_alloc.iloc[i, 7]) else 0
                            rm = float(df_raw_alloc.iloc[i, 8]) if pd.notna(df_raw_alloc.iloc[i, 8]) else 0
                            all_allocs.append({
                                'Date_Alloc': curr_date_str, 'UserID': str(uid).strip(),
                                'VT': vt, 'GB': gb, 'PS': ps, 'RD': rd, 'RM': rm,
                                'dt': pd.to_datetime(curr_date_str, dayfirst=True)
                            })
                        except: continue
                df_alloc = pd.DataFrame(all_allocs).sort_values('dt')

                rows = agg_df.copy()
                rows['dt'] = pd.to_datetime(rows['Date'])
                results = []
                for _, row in rows.iterrows():
                    uid = row['User ID']
                    pdt = row['dt']
                    u_allocs = df_alloc[df_alloc['UserID'] == uid]
                    v_allocs = u_allocs[u_allocs['dt'] <= pdt]
                    if v_allocs.empty:
                        dm = pdt.strftime('%d-%m') if pd.notna(pdt) else ''
                        v_allocs = u_allocs[u_allocs['dt'].dt.strftime('%d-%m') == dm]
                    best_a = v_allocs.iloc[-1] if not v_allocs.empty else None
                    
                    row_dict = row.to_dict()
                    if best_a is not None:
                        vt, gb, ps, rd, rm = best_a['VT'], best_a['GB'], best_a['PS'], best_a['RD'], best_a['RM']
                        tot = vt+gb+ps+rd+rm
                        pnl = row['PNL']
                        row_dict.update({
                            'Alloc_D': best_a['Date_Alloc'],
                            'VT_Rat': vt, 'GB_Rat': gb, 'PS_Rat': ps, 'RD_Rat': rd, 'RM_Rat': rm,
                            'VT_PN': (vt/tot)*pnl if tot>0 else 0, 'GB_PN': (gb/tot)*pnl if tot>0 else 0,
                            'PS_PN': (ps/tot)*pnl if tot>0 else 0, 'RD_PN': (rd/tot)*pnl if tot>0 else 0,
                            'RM_PN': (rm/tot)*pnl if tot>0 else 0
                        })
                    results.append(row_dict)
                df_res = pd.DataFrame(results)
                if not df_res.empty:
                    # Rename columns to match requirements
                    rename_map = {
                        'User ID': 'User I',
                        'Underlying': 'Underly'
                    }
                    df_res = df_res.rename(columns=rename_map)
                    cols_order = ['User I', 'Underly', 'DTE', 'PNL', 'Date', 'dt', 'Alloc_D', 
                                  'VT_Rat', 'GB_Rat', 'PS_Rat', 'RD_Rat', 'RM_Rat', 
                                  'VT_PN', 'GB_PN', 'PS_PN', 'RD_PN', 'RM_PN']
                    final_cols = [c for c in cols_order if c in df_res.columns]
                    df_res = df_res[final_cols]

                st.session_state.df_res = df_res
                st.success("Row-wise P/L Distribution calculated!")
                st.dataframe(df_res.head(10))
                buf = BytesIO(); df_res.to_excel(buf, index=False); buf.seek(0)
                st.download_button("📥 Download P/L Distribution (Row)", buf, "PL_Distribution_Row.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")

# ──────────────────────────────────────────────────────────────────
#  Net P/L Distribution
# ──────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Net P/L Distribution")
st.caption("Calculate Net P/L as the sum of Hedge Cost Distribution and Filtered Row-wise P/L (NF or SN depending on selection)")

submit_net_dist = st.button("Process Net P/L Distribution", type="primary", key="btn_net_dist")

if submit_net_dist:
    if 'df_hedge_dist' not in st.session_state or 'df_res' not in st.session_state or 'tab2_user_summary' not in st.session_state:
        st.error("Please process Hedge Cost Distribution and Row-wise P/L Distribution first.")
    else:
        with st.spinner("Calculating Net P/L Distribution..."):
            try:
                df_hedge = st.session_state.df_hedge_dist
                df_pl = st.session_state.df_res
                usr_summary = st.session_state.tab2_user_summary

                exchange_tab2 = st.session_state.get('tab2_hedge_exchange', 'NFO')
                underlying_filter = "NF" if exchange_tab2 == "NFO" else "SN"

                if 'Underly' in df_pl.columns:
                    df_pl_filtered = df_pl[df_pl['Underly'] == underlying_filter]
                else:
                    df_pl_filtered = df_pl

                net_dist_rows = []
                for _, urow in usr_summary.iterrows():
                    u_id = urow["User ID"]
                    n_pnl = urow["Segment_Result"]
                    
                    h_row = df_hedge[df_hedge["User ID"] == u_id]
                    if not h_row.empty:
                        vt_h = h_row.iloc[0]["VT_Hedge"]
                        gb_h = h_row.iloc[0]["GB_Hedge"]
                        ps_h = h_row.iloc[0]["PS_Hedge"]
                        rd_h = h_row.iloc[0]["RD_Hedge"]
                        rm_h = h_row.iloc[0]["RM_Hedge"]
                    else:
                        vt_h = gb_h = ps_h = rd_h = rm_h = 0.0

                    user_col = 'User I' if 'User I' in df_pl_filtered.columns else 'User ID'
                    p_rows = df_pl_filtered[df_pl_filtered[user_col] == u_id]
                    if not p_rows.empty:
                        vt_p = p_rows['VT_PN'].sum() if 'VT_PN' in p_rows.columns else 0
                        gb_p = p_rows['GB_PN'].sum() if 'GB_PN' in p_rows.columns else 0
                        ps_p = p_rows['PS_PN'].sum() if 'PS_PN' in p_rows.columns else 0
                        rd_p = p_rows['RD_PN'].sum() if 'RD_PN' in p_rows.columns else 0
                        rm_p = p_rows['RM_PN'].sum() if 'RM_PN' in p_rows.columns else 0
                    else:
                        vt_p = gb_p = ps_p = rd_p = rm_p = 0.0

                    vt_net = vt_h + vt_p
                    gb_net = gb_h + gb_p
                    ps_net = ps_h + ps_p
                    rd_net = rd_h + rd_p
                    rm_net = rm_h + rm_p
                    
                    net_dist_rows.append({
                        "User ID": u_id,
                        "Net_PNL": vt_net + gb_net + ps_net + rd_net + rm_net,
                        "VT_Net": vt_net,
                        "GB_Net": gb_net,
                        "PS_Net": ps_net,
                        "RD_Net": rd_net,
                        "RM_Net": rm_net,
                    })

                df_net_dist = pd.DataFrame(net_dist_rows)
                st.success("Net P/L Distribution calculated!")
                st.dataframe(df_net_dist.head(10))
                buf = BytesIO()
                df_net_dist.to_excel(buf, index=False)
                buf.seek(0)
                st.download_button("📥 Download Net P/L Distribution", buf, "Net_PNL_Distribution.xlsx")
            except Exception as e:
                st.error(f"Error: {e}")
