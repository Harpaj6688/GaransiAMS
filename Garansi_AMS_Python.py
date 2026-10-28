import datetime
from datetime import datetime
import numpy as np
import pandas as pd
import plotly.express as px
import pydeck as pdk
import streamlit as st

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Dashboard Tracking Garansi AMS", page_icon="📦", layout="wide"
)

st.title("📦 Dashboard Tracking & Monitoring Garansi AMS")
st.markdown(
    "Aplikasi pemantauan status garansi, approval, dan pengiriman material"
    " secara terpusat."
)

# File Uploader untuk file Excel utama
uploaded_file = st.file_uploader(
    "Upload file Excel `Garansi_AMS.xlsx`", type=["xlsx"]
)

if uploaded_file is not None:

  @st.cache_data
  def load_data(file):
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names

    df_claim = (
        pd.read_excel(xls, sheet_name="Claim")
        if "Claim" in sheet_names
        else pd.DataFrame()
    )
    df_manual = (
        pd.read_excel(xls, sheet_name="Manual")
        if "Manual" in sheet_names
        else pd.DataFrame()
    )
    df_engpar = (
        pd.read_excel(xls, sheet_name="Engpar")
        if "Engpar" in sheet_names
        else pd.DataFrame()
    )

    return df_claim, df_manual, df_engpar

  df_claim, df_manual, df_engpar = load_data(uploaded_file)
  st.success("File Excel berhasil dimuat dengan sukses!")

  # Preprocessing Tanggal & Koordinat untuk Manual & Engpar
  today = pd.to_datetime(datetime.today().date())

  # --- AMBIL TANGGAL CREATED DARI SHEET CLAIM & MERGE KE MANUAL ---
  if not df_claim.empty and not df_manual.empty:
    claim_date_col = None
    for col in df_claim.columns:
      col_lower = col.lower()
      if "date" in col_lower or "tanggal" in col_lower or "created" in col_lower:
        claim_date_col = col
        break

    if not claim_date_col and len(df_claim.columns) > 0:
      claim_date_col = df_claim.columns[0]

    join_key = None
    for key in ["Pre Approval Cert#", "Site Id", "Cert#"]:
      if key in df_claim.columns and key in df_manual.columns:
        join_key = key
        break

    if claim_date_col and join_key:
      df_claim_temp = df_claim[[join_key, claim_date_col]].drop_duplicates(
          subset=[join_key]
      )
      df_claim_temp["Date_Created_dt"] = pd.to_datetime(
          df_claim_temp[claim_date_col], format="%d/%m/%Y", errors="coerce"
      )

      if df_claim_temp["Date_Created_dt"].isnull().all():
        df_claim_temp["Date_Created_dt"] = pd.to_datetime(
            df_claim_temp[claim_date_col], errors="coerce"
        )

      df_manual = df_manual.merge(
          df_claim_temp[[join_key, "Date_Created_dt"]],
          on=join_key,
          how="left",
      )

  if not df_manual.empty:
    if "Date_Created_dt" not in df_manual.columns:
      df_manual["Date_Created_dt"] = pd.to_datetime(today)

    df_manual["Aging_Created_Days"] = (
        today - df_manual["Date_Created_dt"]
    ).dt.days

    if "Date Pengiriman MSS" in df_manual.columns:
      df_manual["Date_MSS_dt"] = pd.to_datetime(
          df_manual["Date Pengiriman MSS"], format="%d/%m/%Y", errors="coerce"
      )
      if df_manual["Date_MSS_dt"].isnull().all():
        df_manual["Date_MSS_dt"] = pd.to_datetime(
            df_manual["Date Pengiriman MSS"], errors="coerce"
        )
      df_manual["Aging_MSS_Days"] = (today - df_manual["Date_MSS_dt"]).dt.days
    else:
      df_manual["Aging_MSS_Days"] = np.nan

    if "Date Pengiriman KSO" in df_manual.columns:
      df_manual["Date_KSO_dt"] = pd.to_datetime(
          df_manual["Date Pengiriman KSO"], format="%d/%m/%Y", errors="coerce"
      )
      if df_manual["Date_KSO_dt"].isnull().all():
        df_manual["Date_KSO_dt"] = pd.to_datetime(
            df_manual["Date Pengiriman KSO"], errors="coerce"
        )
      df_manual["Aging_KSO_Days"] = (today - df_manual["Date_KSO_dt"]).dt.days
    else:
      df_manual["Aging_KSO_Days"] = np.nan

    # Merge Koordinat jika ada di Sheet Engpar
    if not df_engpar.empty and "Site Id" in df_manual.columns:
      lat_col = [c for c in df_engpar.columns if "lat" in c.lower()]
      lon_col = [
          c
          for c in df_engpar.columns
          if "lon" in c.lower() or "long" in c.lower()
      ]
      if lat_col and lon_col and "Site Id" in df_engpar.columns:
        df_manual = df_manual.merge(
            df_engpar[["Site Id", lat_col[0], lon_col[0]]].drop_duplicates(
                "Site Id"
            ),
            on="Site Id",
            how="left",
        )
        df_manual.rename(
            columns={lat_col[0]: "latitude", lon_col[0]: "longitude"},
            inplace=True,
        )

  # --- TAB NAVIGASI UTAMA ---
  tab1, tab2, tab3, tab4, tab5 = st.tabs([
      "📊 Dashboard Approval (Claim)",
      "🚚 Dashboard Pengiriman (Manual)",
      "📋 Data Sheet Manual",
      "📥 Data Sheet Claim",
      "🗺️ Data Sheet Engpar",
  ])

  # =========================================================================
  # --- TAB 1: DASHBOARD APPROVAL (Bersumber dari Sheet Claim + Pie & Bar Chart) ---
  # =========================================================================
  with tab1:
    st.subheader(
        "📊 Dashboard Proses Approval & Statistik Claim (Sheet Claim)"
    )

    if not df_claim.empty:
      total_tiket = len(df_claim)
      status_col = (
          "Overall Status"
          if "Overall Status" in df_claim.columns
          else ("Request Status" if "Request Status" in df_claim.columns else None)
      )

      rejected_cnt = 0
      canceled_cnt = 0
      in_progress_cnt = total_tiket

      if status_col:
        rejected_cnt = (
            df_claim[status_col]
            .astype(str)
            .str.contains("Reject", case=False, na=False)
            .sum()
        )
        canceled_cnt = (
            df_claim[status_col]
            .astype(str)
            .str.contains("Cancel", case=False, na=False)
            .sum()
        )
        in_progress_cnt = total_tiket - (rejected_cnt + canceled_cnt)

      m1, m2, m3, m4 = st.columns(4)
      m1.metric("Total Tiket", total_tiket)
      m2.metric("In Progress", int(in_progress_cnt))
      m3.metric("Rejected", int(rejected_cnt))
      m4.metric("Canceled", int(canceled_cnt))

      st.divider()

      # --- VISUALISASI PIE CHART & BAR CHART DENGAN LABELS ANGKAT ---
      st.markdown("### 📈 Grafik Analisis Approval & Claim")
      col_c1, col_c2 = st.columns(2)

      with col_c1:
        if "Current Phase" in df_claim.columns:
          phase_counts = df_claim["Current Phase"].value_counts().reset_index()
          phase_counts.columns = ["Phase", "Jumlah"]
          fig_pie_claim = px.pie(
              phase_counts,
              values="Jumlah",
              names="Phase",
              hole=0.4,
              title="<b>Distribusi Berdasarkan Current Phase</b>",
          )
          fig_pie_claim.update_traces(textinfo="label+value")
          fig_pie_claim.update_layout(margin=dict(t=40, b=20, l=20, r=20))
          st.plotly_chart(fig_pie_claim, use_container_width=True)

      with col_c2:
        if "Material" in df_claim.columns:
          mat_counts = (
              df_claim["Material"].value_counts().head(7).reset_index()
          )
          mat_counts.columns = ["Material", "Jumlah"]
          fig_bar_claim = px.bar(
              mat_counts,
              x="Jumlah",
              y="Material",
              orientation="h",
              title="<b>Top 7 Material Claim (Bar Chart)</b>",
              color="Jumlah",
              color_continuous_scale="Blues",
          )
          fig_bar_claim.update_traces(
              texttemplate="%{x}", textposition="outside"
          )
          fig_bar_claim.update_layout(
              yaxis=dict(autorange="reversed"),
              margin=dict(t=40, b=20, l=20, r=20),
          )
          st.plotly_chart(fig_bar_claim, use_container_width=True)

      st.divider()

      # Matriks Breakdown Mandiri per Fase
      st.markdown(
          "### 📋 Status Breakdown by Phase (Matriks Monitoring Approval)"
      )
      phase_col = (
          "Current Phase" if "Current Phase" in df_claim.columns else None
      )

      if phase_col:
        df_pa = df_claim[
            df_claim[phase_col].astype(str).str.strip().str.upper()
            == "PRE_APPROVAL"
        ]
        df_req = df_claim[
            df_claim[phase_col].astype(str).str.strip().str.upper()
            == "REQUEST_GARANSI"
        ]
        df_pbp = df_claim[
            df_claim[phase_col].astype(str).str.strip().str.upper()
            == "PERBAIKAN_ASSET"
        ]

        total_pa = len(df_pa)
        total_req = len(df_req)
        total_pbp = len(df_pbp)

        def get_count(sub_df, col_name, keyword):
          if col_name not in sub_df.columns:
            return 0
          return (
              sub_df[col_name]
              .astype(str)
              .str.strip()
              .str.lower()
              .str.contains(keyword.lower(), na=False)
              .sum()
          )

        pa_status_col = "Pre Approval Status"
        req_status_col = "Request Status"
        pbp_status_col = "Perbaikan Assets Status"

        pa_mss = get_count(df_pa, pa_status_col, "mss pending")
        pa_kso = get_count(df_pa, pa_status_col, "kso pending")
        pa_pmo = get_count(df_pa, pa_status_col, "pmo pending")
        pa_cancel = get_count(df_pa, pa_status_col, "cancel")
        pa_rej = get_count(df_pa, pa_status_col, "reject")
        pa_app = (
            max(0, total_pa - (pa_mss + pa_kso + pa_pmo + pa_cancel + pa_rej))
            if get_count(df_pa, pa_status_col, "approve") == 0
            else get_count(df_pa, pa_status_col, "approve")
        )

        pa_data = [
            {"Status / PIC": "Pre Approval (Total)", "Jumlah": total_pa},
            {"Status / PIC": "MSS Pending", "Jumlah": pa_mss},
            {"Status / PIC": "KSO Pending", "Jumlah": pa_kso},
            {"Status / PIC": "PMO Pending", "Jumlah": pa_pmo},
            {"Status / PIC": "Cancel", "Jumlah": pa_cancel},
            {"Status / PIC": "Reject", "Jumlah": pa_rej},
            {"Status / PIC": "Approved", "Jumlah": pa_app},
        ]

        req_mss_form = get_count(df_req, req_status_col, "mss pending form")
        req_mss_app = get_count(df_req, req_status_col, "mss pending approval")
        req_kso = get_count(df_req, req_status_col, "kso pending")
        req_pmo = get_count(df_req, req_status_col, "pmo pending")
        req_cancel = get_count(df_req, req_status_col, "cancel")
        req_rej = get_count(df_req, req_status_col, "reject")
        req_app = (
            max(
                0,
                total_req
                - (
                    req_mss_form
                    + req_mss_app
                    + req_kso
                    + req_pmo
                    + req_cancel
                    + req_rej
                ),
            )
            if get_count(df_req, req_status_col, "approve") == 0
            else get_count(df_req, req_status_col, "approve")
        )

        req_data = [
            {"Status / PIC": "Request Garansi (Total)", "Jumlah": total_req},
            {"Status / PIC": "MSS Pending Form", "Jumlah": req_mss_form},
            {"Status / PIC": "MSS Pending Approval", "Jumlah": req_mss_app},
            {"Status / PIC": "KSO Pending", "Jumlah": req_kso},
            {"Status / PIC": "PMO Pending", "Jumlah": req_pmo},
            {"Status / PIC": "Cancel", "Jumlah": req_cancel},
            {"Status / PIC": "Reject", "Jumlah": req_rej},
            {"Status / PIC": "Approved", "Jumlah": req_app},
        ]

        pbp_mss_form = get_count(df_pbp, pbp_status_col, "mss pending form")
        pbp_mss_app = get_count(df_pbp, pbp_status_col, "mss pending approval")
        pbp_kso = get_count(df_pbp, pbp_status_col, "kso pending")
        pbp_pmo = get_count(df_pbp, pbp_status_col, "pmo pending")
        pbp_pmo_head = get_count(df_pbp, pbp_status_col, "pmo head pending")
        pbp_bakti_head = get_count(df_pbp, pbp_status_col, "bakti head pending")
        pbp_bakti_ppk = get_count(df_pbp, pbp_status_col, "bakti ppk pending")
        pbp_cancel = get_count(df_pbp, pbp_status_col, "cancel")
        pbp_rej = get_count(df_pbp, pbp_status_col, "reject")
        pbp_app = (
            max(
                0,
                total_pbp
                - (
                    pbp_mss_form
                    + pbp_mss_app
                    + pbp_kso
                    + pbp_pmo
                    + pbp_pmo_head
                    + pbp_bakti_head
                    + pbp_bakti_ppk
                    + pbp_cancel
                    + pbp_rej
                ),
            )
            if get_count(df_pbp, pbp_status_col, "approve") == 0
            else get_count(df_pbp, pbp_status_col, "approve")
        )

        pbp_data = [
            {"Status / PIC": "Perbaikan Asset (Total)", "Jumlah": total_pbp},
            {"Status / PIC": "MSS Pending Form", "Jumlah": pbp_mss_form},
            {"Status / PIC": "MSS Pending Approval", "Jumlah": pbp_mss_app},
            {"Status / PIC": "KSO Pending", "Jumlah": pbp_kso},
            {"Status / PIC": "PMO Pending", "Jumlah": pbp_pmo},
            {"Status / PIC": "PMO Head Pending", "Jumlah": pbp_pmo_head},
            {"Status / PIC": "Bakti Head Pending", "Jumlah": pbp_bakti_head},
            {"Status / PIC": "Bakti PPK Pending", "Jumlah": pbp_bakti_ppk},
            {"Status / PIC": "Cancel", "Jumlah": pbp_cancel},
            {"Status / PIC": "Reject", "Jumlah": pbp_rej},
            {"Status / PIC": "Approved", "Jumlah": pbp_app},
        ]

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
          st.markdown("##### 📌 Pre Approval")
          st.dataframe(
              pd.DataFrame(pa_data), use_container_width=True, hide_index=True
          )
        with col_m2:
          st.markdown("##### 📌 Request Garansi")
          st.dataframe(
              pd.DataFrame(req_data), use_container_width=True, hide_index=True
          )
        with col_m3:
          st.markdown("##### 📌 Perbaikan Asset")
          st.dataframe(
              pd.DataFrame(pbp_data), use_container_width=True, hide_index=True
          )

      else:
        st.info("Kolom 'Current Phase' tidak terdeteksi.")

      st.divider()

      # Drill-down & Pencarian
      st.markdown("### 🔍 Drill-Down & Pencarian Detail Data")
      col_f1, col_f2, col_f3 = st.columns(3)
      list_phases = (
          sorted(df_claim[phase_col].dropna().unique().tolist())
          if phase_col in df_claim.columns
          else []
      )
      with col_f1:
        selected_phase = st.selectbox(
            "1️⃣ Pilih Current Phase:", ["Semua Phase"] + list_phases
        )

      if selected_phase == "PRE_APPROVAL":
        dynamic_statuses = [
            "MSS Pending",
            "KSO Pending",
            "PMO Pending",
            "Cancel",
            "Reject",
            "Approved",
        ]
        target_status_col = "Pre Approval Status"
      elif selected_phase == "REQUEST_GARANSI":
        dynamic_statuses = [
            "MSS Pending Form",
            "MSS Pending Approval",
            "KSO Pending",
            "PMO Pending",
            "Cancel",
            "Reject",
            "Approved",
        ]
        target_status_col = "Request Status"
      elif selected_phase == "PERBAIKAN_ASSET":
        dynamic_statuses = [
            "MSS Pending Form",
            "MSS Pending Approval",
            "KSO Pending",
            "PMO Pending",
            "PMO Head Pending",
            "Bakti Head Pending",
            "Bakti PPK Pending",
            "Cancel",
            "Reject",
            "Approved",
        ]
        target_status_col = "Perbaikan Assets Status"
      else:
        target_status_col = None
        dynamic_statuses = []
        for sc in [
            "Pre Approval Status",
            "Request Status",
            "Perbaikan Assets Status",
        ]:
          if sc in df_claim.columns:
            dynamic_statuses.extend(df_claim[sc].dropna().unique().tolist())
        dynamic_statuses = sorted(list(set([str(x) for x in dynamic_statuses])))

      with col_f2:
        selected_sub_status = st.selectbox(
            "2️⃣ Pilih Status Rinci:", ["Semua Status Rinci"] + dynamic_statuses
        )
      with col_f3:
        search_ticket = st.text_input(
            "3️⃣ Cari No Tiket / Site ID / Cert#:",
            placeholder="Ketik Site ID / Cert#...",
        )

      df_filtered_drill = df_claim.copy()
      if selected_phase != "Semua Phase":
        df_filtered_drill = df_filtered_drill[
            df_filtered_drill[phase_col] == selected_phase
        ]

      if selected_sub_status != "Semua Status Rinci":
        if target_status_col and target_status_col in df_filtered_drill.columns:
          df_filtered_drill = df_filtered_drill[
              df_filtered_drill[target_status_col]
              .astype(str)
              .str.strip()
              .str.lower()
              .str.contains(selected_sub_status.lower(), na=False)
          ]

      if search_ticket:
        mask_search = False
        for col in df_filtered_drill.columns:
          mask_search = mask_search | df_filtered_drill[col].astype(
              str
          ).str.contains(search_ticket, case=False, na=False)
        df_filtered_drill = df_filtered_drill[mask_search]

      st.info(f"Menampilkan **{len(df_filtered_drill)}** baris data.")
      st.dataframe(df_filtered_drill, use_container_width=True)
    else:
      st.warning("Sheet 'Claim' kosong atau tidak ditemukan.")

  # =========================================================================
  # --- TAB 2: DASHBOARD PENGIRIMAN & AGING (Bersumber dari Sheet Manual) ---
  # =========================================================================
  with tab2:
    st.subheader(
        "🚚 Dashboard Pengiriman Barang MSS & KSO (Sheet Manual & Analytics)"
    )

    if not df_manual.empty:
      df_manual_calc = df_manual.copy()

      # --- METRIK REVISI SCOREBOARD ---
      # 1. MSS Need Send
      df_mss_need = df_manual_calc[
          df_manual_calc["Status Pengiriman MSS"]
          .astype(str)
          .str.strip()
          .eq("MSS Need Send")
      ].copy()
      mss_need_send_cnt = len(df_mss_need)
      max_aging_mss_need = (
          df_mss_need["Aging_Created_Days"].max()
          if not df_mss_need.empty
          else 0
      )
      if pd.isna(max_aging_mss_need):
        max_aging_mss_need = 0

      # 2. MSS Done Send
      df_mss_done = df_manual_calc[
          df_manual_calc["Status Pengiriman MSS"]
          .astype(str)
          .str.strip()
          .eq("MSS Done Send")
      ].copy()
      mss_done_send_cnt = len(df_mss_done)
      max_aging_mss_done = (
          df_mss_done["Aging_MSS_Days"].max() if not df_mss_done.empty else 0
      )
      if pd.isna(max_aging_mss_done):
        max_aging_mss_done = 0

      # 3. KSO Done Send
      df_kso_done = df_manual_calc[
          df_manual_calc["Date_KSO_dt"].notnull()
      ].copy()
      total_kso_done = len(df_kso_done)

      mc1, mc2, mc3 = st.columns(3)
      with mc1:
        st.metric(
            label="🚨 MSS Need Send",
            value=int(mss_need_send_cnt),
            delta=(
                f"Max Aging: {int(max_aging_mss_need)} Hari"
                if max_aging_mss_need > 0
                else "0 Hari"
            ),
            delta_color="inverse",
            help="Total material berstatus MSS Need Send beserta Max Aging dari tanggal Created (Sheet Claim)",
        )
      with mc2:
        st.metric(
            label="✅ MSS Done Send",
            value=int(mss_done_send_cnt),
            delta=(
                f"Max Aging: {int(max_aging_mss_done)} Hari"
                if max_aging_mss_done > 0
                else "0 Hari"
            ),
            delta_color="inverse",
            help="Total material berstatus MSS Done Send beserta Max Aging dari tanggal MSS Done Send",
        )
      with mc3:
        st.metric(
            label="🏢 KSO Done Send",
            value=int(total_kso_done),
            help="Total barang pengganti yang sudah dikirim oleh KSO",
        )

      st.divider()

      # --- TAB 2: PETA SEBARAN LOKASI ---
      if (
          "latitude" in df_manual_calc.columns
          and "longitude" in df_manual_calc.columns
      ):
        df_map = df_manual_calc.dropna(
            subset=["latitude", "longitude"]
        ).copy()
        if not df_map.empty:
          st.markdown(
              "### 🗺️ OpenStreetMap / GIS Sebaran Lokasi Site Pengiriman"
          )
          st.pydeck_chart(
              pdk.Deck(
                  map_style="mapbox://styles/mapbox/light-v10",
                  initial_view_state=pdk.ViewState(
                      latitude=df_map["latitude"].mean(),
                      longitude=df_map["longitude"].mean(),
                      zoom=4,
                      pitch=0,
                  ),
                  layers=[
                      pdk.Layer(
                          "ScatterplotLayer",
                          data=df_map,
                          get_position="[longitude, latitude]",
                          get_color="[37, 99, 235, 180]",
                          get_radius=20000,
                          pickable=True,
                      ),
                  ],
                  tooltip={
                      "html": (
                          "<b>Site ID:</b> {Site Id}<br/><b>Sitename:</b>"
                          " {Sitename}<br/><b>Provinsi:</b> {Provinsi}"
                      ),
                      "style": {
                          "color": "white",
                          "backgroundColor": "#1E293B",
                      },
                  },
              )
          )
          st.divider()

      # --- DAFTAR KOLOM UTAMA + RESI PENGIRIMAN + AGING DAYS ---
      exact_columns = [
          "Pre Approval Cert#",
          "Site Id",
          "Sitename",
          "Provinsi",
          "Kota/Kab",
          "Material",
          "Quantity",
          "Current Phase",
          "Pre Approval Status",
          "Request Status",
          "Perbaikan Assets Status",
          "MSS Need Send?",
          "KSO Need Send?",
          "Date Pengiriman MSS",
          "Status Pengiriman MSS",
          "Resi Pengiriman",  # Kolom Resi Pengiriman yang baru ditambahkan
          "Date Pengiriman KSO",
          "Status Pengiriman KSO / On hand MSS",
      ]

      def get_unique_valid_cols(df, base_cols, aging_col_name):
        available = [c for c in base_cols if c in df.columns]
        if aging_col_name in df.columns and aging_col_name not in available:
          available.append(aging_col_name)
        return list(dict.fromkeys(available))

      # --- 1. DETAIL TABEL PENGIRIMAN MSS & STATUS APPROVAL ---
      st.markdown("### 📦 Detail Tabel Pengiriman MSS & Status Approval")

      df_mss_summary = df_manual_calc[
          df_manual_calc["Status Pengiriman MSS"]
          .astype(str)
          .str.strip()
          .isin(["MSS Need Send", "MSS Done Send"])
      ].copy()

      if not df_mss_summary.empty:
        # Menyesuaikan kolom aging berdasarkan status masing-masing baris di tabel detail
        df_mss_summary["Aging Days"] = np.where(
            df_mss_summary["Status Pengiriman MSS"].str.strip()
            == "MSS Need Send",
            df_mss_summary["Aging_Created_Days"],
            df_mss_summary["Aging_MSS_Days"],
        )

        # --- FITUR FILTER STATUS PENGIRIMAN MSS ---
        available_mss_statuses = (
            df_mss_summary["Status Pengiriman MSS"]
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )
        selected_mss_status = st.selectbox(
            "🔍 Filter Status Pengiriman MSS:",
            options=["Semua Status"] + available_mss_statuses,
            index=0,
            help="Filter data tabel pengiriman MSS berdasarkan statusnya",
        )

        # Apply Filter
        if selected_mss_status != "Semua Status":
          df_mss_filtered = df_mss_summary[
              df_mss_summary["Status Pengiriman MSS"].str.strip()
              == selected_mss_status
          ]
        else:
          df_mss_filtered = df_mss_summary.copy()

        df_mss_summary_sorted = df_mss_filtered.sort_values(
            by=["Status Pengiriman MSS", "Aging Days"], ascending=[False, False]
        )

        cols_mss = get_unique_valid_cols(
            df_mss_summary_sorted, exact_columns, "Aging Days"
        )
        st.dataframe(
            df_mss_summary_sorted[cols_mss],
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.info("Tidak ada data pengiriman MSS.")

      st.divider()

      # --- 2. DETAIL BARANG PENGGANTI DARI KSO ---
      st.markdown("### 🚛 Detail Barang Pengganti dari KSO")
      if not df_kso_done.empty:
        df_kso_done = df_kso_done.sort_values(
            by=["Aging_KSO_Days"], ascending=False
        )
        df_kso_done = df_kso_done.rename(
            columns={"Aging_KSO_Days": "Aging Days"}
        )

        cols_kso = get_unique_valid_cols(
            df_kso_done, exact_columns, "Aging Days"
        )
        st.dataframe(
            df_kso_done[cols_kso], use_container_width=True, hide_index=True
        )
      else:
        st.info("Belum ada barang pengganti yang dikirim oleh KSO.")

    else:
      st.warning("Sheet 'Manual' kosong atau tidak ditemukan.")

  # =========================================================================
  # --- TAB 3, 4, 5: DETAIL SHEET ---
  # =========================================================================
  with tab3:
    st.subheader("📋 Detail Sheet Manual (Tracking Pengiriman)")
    if not df_manual.empty:
      st.dataframe(df_manual, use_container_width=True)

  with tab4:
    st.subheader("📥 Detail Sheet Claim (Master Data Approval)")
    if not df_claim.empty:
      st.dataframe(df_claim, use_container_width=True)

  with tab5:
    st.subheader("🗺️ Detail Sheet Engpar (Referensi)")
    if not df_engpar.empty:
      st.dataframe(df_engpar, use_container_width=True)

else:
  st.info(
      "👋 Silakan upload file Excel `Garansi_AMS.xlsx` melalui tombol di atas"
      " untuk mulai melihat dashboard."
  )