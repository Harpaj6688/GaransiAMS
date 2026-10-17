import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Dashboard Tracking Garansi AMS",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Dashboard Tracking & Monitoring Garansi AMS")
st.markdown("Aplikasi pemantauan status garansi, approval, dan pengiriman material secara terpusat.")

# File Uploader untuk file Excel utama
uploaded_file = st.file_uploader("Upload file Excel `Garansi_AMS.xlsx`", type=["xlsx"])

if uploaded_file is not None:
    @st.cache_data
    def load_data(file):
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        
        df_claim = pd.read_excel(xls, sheet_name="Claim") if "Claim" in sheet_names else pd.DataFrame()
        df_manual = pd.read_excel(xls, sheet_name="Manual") if "Manual" in sheet_names else pd.DataFrame()
        df_engpar = pd.read_excel(xls, sheet_name="Engpar") if "Engpar" in sheet_names else pd.DataFrame()
        
        return df_claim, df_manual, df_engpar

    df_claim, df_manual, df_engpar = load_data(uploaded_file)
    st.success("File Excel berhasil dimuat dengan sukses!")

    # --- TAB NAVIGASI UTAMA ---
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Dashboard Approval (Claim)", 
        "🚚 Dashboard Pengiriman (Manual)", 
        "📋 Data Sheet Manual", 
        "📥 Data Sheet Claim", 
        "🗺️ Data Sheet Engpar"
    ])

    # =========================================================================
    # --- TAB 1: DASHBOARD APPROVAL (Bersumber dari Sheet Claim) ---
    # =========================================================================
    with tab1:
        st.subheader("📊 Dashboard Proses Approval (Sheet Claim)")
        
        if not df_claim.empty:
            # 1. Metrik Atas (Total Tiket, In Progress, Rejected, Canceled)
            total_tiket = len(df_claim)
            status_col = 'Overall Status' if 'Overall Status' in df_claim.columns else ('Request Status' if 'Request Status' in df_claim.columns else None)
            
            rejected_cnt = 0
            canceled_cnt = 0
            in_progress_cnt = total_tiket
            
            if status_col:
                rejected_cnt = df_claim[status_col].astype(str).str.contains("Reject", case=False, na=False).sum()
                canceled_cnt = df_claim[status_col].astype(str).str.contains("Cancel", case=False, na=False).sum()
                in_progress_cnt = total_tiket - (rejected_cnt + canceled_cnt)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Tiket", total_tiket)
            m2.metric("In Progress", int(in_progress_cnt))
            m3.metric("Rejected", int(rejected_cnt))
            m4.metric("Canceled", int(canceled_cnt))
            
            st.divider()
            
            # 2. Matriks Breakdown Mandiri per Fase
            st.markdown("### 📋 Status Breakdown by Phase (Matriks Monitoring Approval)")
            
            phase_col = 'Current Phase' if 'Current Phase' in df_claim.columns else None
            
            if phase_col:
                df_pa = df_claim[df_claim[phase_col].astype(str).str.strip().str.upper() == 'PRE_APPROVAL']
                df_req = df_claim[df_claim[phase_col].astype(str).str.strip().str.upper() == 'REQUEST_GARANSI']
                df_pbp = df_claim[df_claim[phase_col].astype(str).str.strip().str.upper() == 'PERBAIKAN_ASSET']
                
                total_pa = len(df_pa)
                total_req = len(df_req)
                total_pbp = len(df_pbp)
                
                def get_count(sub_df, col_name, keyword):
                    if col_name not in sub_df.columns:
                        return 0
                    return sub_df[col_name].astype(str).str.strip().str.lower().str.contains(keyword.lower(), na=False).sum()

                pa_status_col = 'Pre Approval Status'
                req_status_col = 'Request Status'
                pbp_status_col = 'Perbaikan Assets Status'

                # --- HITUNG PRE APPROVAL ---
                pa_mss = get_count(df_pa, pa_status_col, 'mss pending')
                pa_kso = get_count(df_pa, pa_status_col, 'kso pending')
                pa_pmo = get_count(df_pa, pa_status_col, 'pmo pending')
                pa_cancel = get_count(df_pa, pa_status_col, 'cancel')
                pa_rej = get_count(df_pa, pa_status_col, 'reject')
                pa_app = max(0, total_pa - (pa_mss + pa_kso + pa_pmo + pa_cancel + pa_rej)) if get_count(df_pa, pa_status_col, 'approve') == 0 else get_count(df_pa, pa_status_col, 'approve')

                pa_data = [
                    {"Status / PIC": "Pre Approval (Total)", "Jumlah": total_pa},
                    {"Status / PIC": "MSS Pending", "Jumlah": pa_mss},
                    {"Status / PIC": "KSO Pending", "Jumlah": pa_kso},
                    {"Status / PIC": "PMO Pending", "Jumlah": pa_pmo},
                    {"Status / PIC": "Cancel", "Jumlah": pa_cancel},
                    {"Status / PIC": "Reject", "Jumlah": pa_rej},
                    {"Status / PIC": "Approved", "Jumlah": pa_app}
                ]

                # --- HITUNG REQUEST GARANSI ---
                req_mss_form = get_count(df_req, req_status_col, 'mss pending form')
                req_mss_app = get_count(df_req, req_status_col, 'mss pending approval')
                req_kso = get_count(df_req, req_status_col, 'kso pending')
                req_pmo = get_count(df_req, req_status_col, 'pmo pending')
                req_cancel = get_count(df_req, req_status_col, 'cancel')
                req_rej = get_count(df_req, req_status_col, 'reject')
                req_app = max(0, total_req - (req_mss_form + req_mss_app + req_kso + req_pmo + req_cancel + req_rej)) if get_count(df_req, req_status_col, 'approve') == 0 else get_count(df_req, req_status_col, 'approve')

                req_data = [
                    {"Status / PIC": "Request Garansi (Total)", "Jumlah": total_req},
                    {"Status / PIC": "MSS Pending Form", "Jumlah": req_mss_form},
                    {"Status / PIC": "MSS Pending Approval", "Jumlah": req_mss_app},
                    {"Status / PIC": "KSO Pending", "Jumlah": req_kso},
                    {"Status / PIC": "PMO Pending", "Jumlah": req_pmo},
                    {"Status / PIC": "Cancel", "Jumlah": req_cancel},
                    {"Status / PIC": "Reject", "Jumlah": req_rej},
                    {"Status / PIC": "Approved", "Jumlah": req_app}
                ]

                # --- HITUNG PERBAIKAN ASSET ---
                pbp_mss_form = get_count(df_pbp, pbp_status_col, 'mss pending form')
                pbp_mss_app = get_count(df_pbp, pbp_status_col, 'mss pending approval')
                pbp_kso = get_count(df_pbp, pbp_status_col, 'kso pending')
                pbp_pmo = get_count(df_pbp, pbp_status_col, 'pmo pending')
                pbp_pmo_head = get_count(df_pbp, pbp_status_col, 'pmo head pending')
                pbp_bakti_head = get_count(df_pbp, pbp_status_col, 'bakti head pending')
                pbp_bakti_ppk = get_count(df_pbp, pbp_status_col, 'bakti ppk pending')
                pbp_cancel = get_count(df_pbp, pbp_status_col, 'cancel')
                pbp_rej = get_count(df_pbp, pbp_status_col, 'reject')
                pbp_app = max(0, total_pbp - (pbp_mss_form + pbp_mss_app + pbp_kso + pbp_pmo + pbp_pmo_head + pbp_bakti_head + pbp_bakti_ppk + pbp_cancel + pbp_rej)) if get_count(df_pbp, pbp_status_col, 'approve') == 0 else get_count(df_pbp, pbp_status_col, 'approve')

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
                    {"Status / PIC": "Approved", "Jumlah": pbp_app}
                ]

                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    st.markdown("##### 📌 Pre Approval")
                    st.dataframe(pd.DataFrame(pa_data), use_container_width=True, hide_index=True)
                with col_m2:
                    st.markdown("##### 📌 Request Garansi")
                    st.dataframe(pd.DataFrame(req_data), use_container_width=True, hide_index=True)
                with col_m3:
                    st.markdown("##### 📌 Perbaikan Asset")
                    st.dataframe(pd.DataFrame(pbp_data), use_container_width=True, hide_index=True)
                
            else:
                st.info("Kolom 'Current Phase' tidak terdeteksi.")
                
            st.divider()
            
            # 3. FITUR DRILL-DOWN & PENCARIAN
            st.markdown("### 🔍 Drill-Down & Pencarian Detail Data")
            st.markdown("Gunakan 3 kategori filter di bawah ini untuk menelusuri data secara spesifik:")
            
            col_f1, col_f2, col_f3 = st.columns(3)
            
            list_phases = sorted(df_claim[phase_col].dropna().unique().tolist()) if phase_col in df_claim.columns else []
            with col_f1:
                selected_phase = st.selectbox("1️⃣ Pilih Current Phase:", ["Semua Phase"] + list_phases)
            
            if selected_phase == "PRE_APPROVAL":
                dynamic_statuses = ["MSS Pending", "KSO Pending", "PMO Pending", "Cancel", "Reject", "Approved"]
                target_status_col = 'Pre Approval Status'
            elif selected_phase == "REQUEST_GARANSI":
                dynamic_statuses = ["MSS Pending Form", "MSS Pending Approval", "KSO Pending", "PMO Pending", "Cancel", "Reject", "Approved"]
                target_status_col = 'Request Status'
            elif selected_phase == "PERBAIKAN_ASSET":
                dynamic_statuses = ["MSS Pending Form", "MSS Pending Approval", "KSO Pending", "PMO Pending", "PMO Head Pending", "Bakti Head Pending", "Bakti PPK Pending", "Cancel", "Reject", "Approved"]
                target_status_col = 'Perbaikan Assets Status'
            else:
                target_status_col = None
                dynamic_statuses = []
                for sc in ['Pre Approval Status', 'Request Status', 'Perbaikan Assets Status']:
                    if sc in df_claim.columns:
                        dynamic_statuses.extend(df_claim[sc].dropna().unique().tolist())
                dynamic_statuses = sorted(list(set([str(x) for x in dynamic_statuses])))

            with col_f2:
                selected_sub_status = st.selectbox("2️⃣ Pilih Status Rinci (Sesuai Hirarki Phase):", ["Semua Status Rinci"] + dynamic_statuses)
            
            with col_f3:
                search_ticket = st.text_input("3️⃣ Cari No Tiket / Site ID / Cert#:", placeholder="Ketik Site ID / Cert#...")
            
            df_filtered_drill = df_claim.copy()
            if selected_phase != "Semua Phase":
                df_filtered_drill = df_filtered_drill[df_filtered_drill[phase_col] == selected_phase]
            
            if selected_sub_status != "Semua Status Rinci":
                if target_status_col and target_status_col in df_filtered_drill.columns:
                    df_filtered_drill = df_filtered_drill[df_filtered_drill[target_status_col].astype(str).str.strip().str.lower().str.contains(selected_sub_status.lower(), na=False)]
                else:
                    mask_status = False
                    for sc in ['Pre Approval Status', 'Request Status', 'Perbaikan Assets Status']:
                        if sc in df_filtered_drill.columns:
                            mask_status = mask_status | df_filtered_drill[sc].astype(str).str.strip().str.lower().str.contains(selected_sub_status.lower(), na=False)
                    df_filtered_drill = df_filtered_drill[mask_status]
            
            if search_ticket:
                mask_search = False
                for col in df_filtered_drill.columns:
                    mask_search = mask_search | df_filtered_drill[col].astype(str).str.contains(search_ticket, case=False, na=False)
                df_filtered_drill = df_filtered_drill[mask_search]
            
            st.info(f"Menampilkan **{len(df_filtered_drill)}** baris data sesuai kriteria filter.")
            st.dataframe(df_filtered_drill, use_container_width=True)
            
        else:
            st.warning("Sheet 'Claim' kosong atau tidak ditemukan.")


   # =========================================================================
    # --- TAB 2: DASHBOARD PENGIRIMAN & AGING (Bersumber dari Sheet Manual) ---
    # =========================================================================
    with tab2:
        st.subheader("🚚 Dashboard Pengiriman Barang MSS & KSO (Sheet Manual)")
        
        if not df_manual.empty:
            df_manual_calc = df_manual.copy()

            # Preprocessing Tanggal
            today = pd.to_datetime(datetime.today().date())
            df_manual_calc['Date_MSS_dt'] = pd.to_datetime(df_manual_calc['Date Pengiriman MSS'], format='%d/%m/%Y', errors='coerce')
            df_manual_calc['Date_KSO_dt'] = pd.to_datetime(df_manual_calc['Date Pengiriman KSO'], format='%d/%m/%Y', errors='coerce')

            # Kalkulasi Aging MSS (dari tanggal kirim MSS sampai hari ini)
            df_manual_calc['Aging_MSS_Days'] = (today - df_manual_calc['Date_MSS_dt']).dt.days
            df_manual_calc['Aging_KSO_Days'] = (today - df_manual_calc['Date_KSO_dt']).dt.days

            # --- 1. COUNTING DASHBOARD METRICS ---
            mss_need_send_cnt = df_manual_calc['Status Pengiriman MSS'].astype(str).str.strip().eq('MSS Need Send').sum()
            mss_done_send_cnt = df_manual_calc['Status Pengiriman MSS'].astype(str).str.strip().eq('MSS Done Send').sum()

            # Data MSS yang sudah dikirim (Done Send)
            df_mss_done = df_manual_calc[df_manual_calc['Status Pengiriman MSS'].astype(str).str.strip().eq('MSS Done Send')].copy()
            avg_aging_mss = df_mss_done['Aging_MSS_Days'].mean() if not df_mss_done.empty else 0

            # Data KSO Done
            df_kso_done = df_manual_calc[df_manual_calc['Date_KSO_dt'].notnull()].copy()
            total_kso_done = len(df_kso_done)
            avg_aging_kso = df_kso_done['Aging_KSO_Days'].mean() if not df_kso_done.empty else 0

            mc1, mc2, mc3 = st.columns(3)
            with mc1:
                st.metric(
                    label="🚨 MSS Need Send", 
                    value=int(mss_need_send_cnt),
                    help="Total material yang berstatus MSS Need Send (butuh dikirim)"
                )

            with mc2:
                st.metric(
                    label="✅ MSS Done Send", 
                    value=int(mss_done_send_cnt),
                    delta=f"Avg Aging: {avg_aging_mss:.0f} Hari" if avg_aging_mss > 0 else "0 Hari",
                    delta_color="inverse",
                    help="Total material yang berstatus MSS Done Send (sudah dikirim)"
                )

            with mc3:
                st.metric(
                    label="🏢 KSO Done Send", 
                    value=int(total_kso_done),
                    delta=f"Avg Aging: {avg_aging_kso:.0f} Hari" if avg_aging_kso > 0 else "0 Hari",
                    delta_color="normal",
                    help="Total barang pengganti yang sudah dikirim oleh KSO"
                )

            st.divider()

            # --- 2. TABEL DETAIL PENGIRIMAN MSS & STATUS APPROVAL ---
            st.markdown("### 📦 Detail Tabel Pengiriman MSS & Status Approval")
            
            # Filter khusus pengiriman MSS (MSS Need Send dan MSS Done Send)
            df_mss_summary = df_manual_calc[
                df_manual_calc['Status Pengiriman MSS'].astype(str).str.strip().isin(['MSS Need Send', 'MSS Done Send'])
            ].copy()

            if not df_mss_summary.empty:
                # Urutkan agar MSS Need Send di atas, lalu berdasarkan Aging
                df_mss_summary_sorted = df_mss_summary.sort_values(by=['Status Pengiriman MSS', 'Aging_MSS_Days'], ascending=[False, False])
                
                # Menampilkan kolom-kolom MSS + Status Approval
                st.dataframe(
                    df_mss_summary_sorted[[
                        'Site Id', 'Sitename', 'Provinsi', 'Material', 'Quantity',
                        'Current Phase', 'Pre Approval Status', 'Request Status', 'Perbaikan Assets Status',
                        'Status Pengiriman MSS', 'Date Pengiriman MSS', 'Aging_MSS_Days'
                    ]].rename(columns={'Aging_MSS_Days': 'Aging MSS (Hari)'}),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("Tidak ada data pengiriman MSS.")

            st.divider()

            # --- 3. TABEL DETAIL KSO DONE ---
            st.markdown("### ✅ Detail Barang Pengganti dari KSO")
            if not df_kso_done.empty:
                st.dataframe(
                    df_kso_done[[
                        'Site Id', 'Sitename', 'Provinsi', 'Material', 'Quantity',
                        'Date Pengiriman KSO', 'Aging_KSO_Days', 'Status Pengiriman KSO / On hand MSS'
                    ]].rename(columns={'Aging_KSO_Days': 'Aging KSO (Hari)'}),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("Belum ada barang pengganti yang dikirim oleh KSO.")

            st.divider()

            # --- 4. DATASET UTAMA SHEET MANUAL ---
            st.markdown("### 📋 Tabel Data Sheet Manual Lengkap")
            st.dataframe(df_manual, use_container_width=True)

        else:
            st.warning("Sheet 'Manual' kosong atau tidak ditemukan.")

    # =========================================================================
    # --- TAB 3: DETAIL SHEET MANUAL ---
    # =========================================================================
    with tab3:
        st.subheader("📋 Detail Sheet Manual (Tracking Pengiriman)")
        if not df_manual.empty:
            st.dataframe(df_manual, use_container_width=True)
        else:
            st.warning("Sheet 'Manual' tidak ditemukan.")

    # =========================================================================
    # --- TAB 4: DETAIL SHEET CLAIM ---
    # =========================================================================
    with tab4:
        st.subheader("📥 Detail Sheet Claim (Master Data Approval)")
        if not df_claim.empty:
            st.dataframe(df_claim, use_container_width=True)
        else:
            st.warning("Sheet 'Claim' tidak ditemukan.")

    # =========================================================================
    # --- TAB 5: DETAIL SHEET ENGPAR ---
    # =========================================================================
    with tab5:
        st.subheader("🗺️ Detail Sheet Engpar (Referensi)")
        if not df_engpar.empty:
            st.dataframe(df_engpar, use_container_width=True)
        else:
            st.warning("Sheet 'Engpar' tidak ditemukan.")

else:
    st.info("👋 Silakan upload file Excel `Garansi_AMS.xlsx` melalui tombol di atas untuk mulai melihat dashboard.")