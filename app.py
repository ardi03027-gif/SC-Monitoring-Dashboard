import os
from datetime import date, datetime
import pandas as pd
import streamlit as st
import plotly.express as px
import urllib.request
import urllib.parse
import re

st.set_page_config(page_title="SC Monitoring Dashboard", page_icon="📊", layout="wide")

NAVY = "#14294B"
BLUE = "#0877C9"
ORANGE = "#F58220"

st.markdown(f"""
<style>
.stApp {{ background: linear-gradient(135deg,#f3f6fb 0%,#e9eef6 100%); }}
section[data-testid="stSidebar"] {{ background: linear-gradient(180deg,{NAVY} 0%,#203e68 100%); }}
section[data-testid="stSidebar"] * {{ color: white !important; }}
div[data-testid="stMetric"] {{ background: white; border: 1px solid #dce4ef; padding: 14px; border-radius: 12px; box-shadow: 0 3px 12px rgba(20,41,75,.06); }}
h1,h2,h3 {{ color: {NAVY}; }}
.jn-hero {{ background: linear-gradient(105deg,{NAVY},#1b5a91 72%,{ORANGE}); color:white; padding:22px 26px; border-radius:16px; margin: 4px 0 18px 0; box-shadow:0 8px 24px rgba(20,41,75,.16); }}
.jn-kicker {{ letter-spacing:2px; font-size:11px; font-weight:800; color:#ffd3a6; }}
.jn-title {{ font-size:27px; font-weight:800; margin-top:4px; }}
.jn-sub {{ color:#eaf3ff; font-size:13px; margin-top:5px; }}
div[data-testid="stDataFrame"] {{ background:white; border-radius:10px; }}
</style>
<div class="jn-hero">
  <div class="jn-kicker">JAKARTANOTEBOOK • SUPPLY CHAIN OPERATIONS</div>
  <div class="jn-title">📦 SC Productivity & Case Monitoring</div>
  <div class="jn-sub">Warehouse performance • PIC productivity • Complaint control • Data quality</div>
</div>
""", unsafe_allow_html=True)

st.caption("Dashboard monitoring produktivitas PIC dan complaint — harian, mingguan, bulanan, dan tahunan.")

DEFAULT_FILE = "Salinan dari Database ID Summary Dec 2025-2026 .xlsx"

@st.cache_data(show_spinner="Membaca database Excel...")
def load_workbook(file_bytes):
    from io import BytesIO
    xl = pd.ExcelFile(BytesIO(file_bytes))
    masters = []
    for sheet, year in [("Master 2025", 2025), ("Master 2026", 2026), ("Master 2027", 2027)]:
        if sheet not in xl.sheet_names:
            continue
        raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet)
        if raw.empty:
            continue
        raw.columns = [str(c).strip() for c in raw.columns]
        date_col = "Tanggal Completed"
        if date_col not in raw.columns:
            continue
        for c in ["ID", "SKU", "QTY", "Nama PIC", date_col]:
            if c not in raw.columns:
                raw[c] = None
        d = raw[["ID", "SKU", "QTY", "Nama PIC", date_col]].copy()
        d["Tahun Sumber"] = year
        d["Tanggal Completed"] = pd.to_datetime(d["Tanggal Completed"], errors="coerce")
        d["Nama PIC"] = d["Nama PIC"].astype("string").str.strip()
        d["ID"] = d["ID"].astype("string").str.strip()
        d["SKU"] = pd.to_numeric(d["SKU"], errors="coerce").fillna(0)
        d["QTY"] = pd.to_numeric(d["QTY"], errors="coerce").fillna(0)
        d = d[d["Tanggal Completed"].notna() & d["Nama PIC"].notna() & (d["Nama PIC"] != "")]
        d["Tanggal"] = d["Tanggal Completed"].dt.date
        d["Tahun"] = d["Tanggal Completed"].dt.year
        d["Bulan"] = d["Tanggal Completed"].dt.to_period("M").astype(str)
        d["Minggu"] = d["Tanggal Completed"].dt.strftime("%G-W%V")
        masters.append(d)
    master = pd.concat(masters, ignore_index=True) if masters else pd.DataFrame(
        columns=["ID","SKU","QTY","Nama PIC","Tanggal Completed","Tahun","Bulan","Minggu","Tanggal"])
    cases = pd.read_excel(BytesIO(file_bytes), sheet_name="Case Pic") if "Case Pic" in xl.sheet_names else pd.DataFrame()
    if not cases.empty:
        cases.columns = [str(c).strip() for c in cases.columns]
        for c in ["Tanggal Diterima","Tanggal Complaint","ID","SKU","Qty Request","Qty Complaint","Kategori","Keterangan Tambahan","Cabang","Keterangan Case","Penyelesaian Case","Nama PIC"]:
            if c not in cases.columns: cases[c] = None
        cases["Tanggal Diterima"] = pd.to_datetime(cases["Tanggal Diterima"], errors="coerce")
        cases["Tanggal Complaint"] = pd.to_datetime(cases["Tanggal Complaint"], errors="coerce")
        cases["Penyelesaian Case"] = cases["Penyelesaian Case"].astype("string").str.strip()
        cases["Kategori"] = cases["Kategori"].astype("string").str.strip()
        cases["Nama PIC"] = cases["Nama PIC"].astype("string").str.strip()
        # Periode operasional Monitoring Case tetap menggunakan Tanggal Diterima.
        cases["Tahun"] = cases["Tanggal Diterima"].dt.year
        cases["Bulan"] = cases["Tanggal Diterima"].dt.to_period("M").astype(str)
        cases["Minggu"] = cases["Tanggal Diterima"].dt.strftime("%G-W%V")
        # Rekap bulanan/tahunan khusus menggunakan Tanggal Complaint.
        cases["Tahun Complaint"] = cases["Tanggal Complaint"].dt.year
        cases["Bulan Complaint"] = cases["Tanggal Complaint"].dt.to_period("M").astype(str)
    return master, cases, xl.sheet_names

@st.cache_data(ttl=300, show_spinner="Mengambil data terbaru dari Google Sheets...")
def fetch_google_sheet_xlsx(sheet_url):
    """Download seluruh workbook Google Sheets sebagai XLSX (sheet harus dapat dibaca akun aplikasi)."""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise ValueError("URL Google Sheets tidak dikenali. Gunakan link berbentuk docs.google.com/spreadsheets/d/...")
    spreadsheet_id = match.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
    request = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=45) as response:
        content = response.read()
    if not content.startswith(b"PK"):
        raise ValueError("Google Sheets tidak mengirim file XLSX. Pastikan akses tautan sebagai Viewer atau siapkan autentikasi untuk sheet privat.")
    return content

st.sidebar.markdown("### SUMBER DATA")
source_mode = st.sidebar.radio("Pilih sumber", ["Google Sheets (otomatis)", "Upload Excel (manual)"])
uploaded = st.sidebar.file_uploader("Upload database Excel (.xlsx)", type=["xlsx"])
local_path = os.path.join(os.path.dirname(__file__), DEFAULT_FILE)
file_bytes = None
source_label = "Excel lokal"

if source_mode == "Google Sheets (otomatis)":
    sheet_url = st.sidebar.text_input("URL Google Sheets", placeholder="https://docs.google.com/spreadsheets/d/...")
    auto_refresh = st.sidebar.checkbox("Refresh otomatis setiap 5 menit", value=True)
    if st.sidebar.button("↻ Refresh data sekarang", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    if sheet_url.strip():
        if auto_refresh:
            try:
                from streamlit_autorefresh import st_autorefresh
                st_autorefresh(interval=300000, key="google_sheet_auto_refresh")
            except ImportError:
                st.sidebar.warning("Auto-refresh belum aktif. Instal streamlit-autorefresh dari requirements.txt.")
        try:
            file_bytes = fetch_google_sheet_xlsx(sheet_url.strip())
            source_label = "Google Sheets"
        except Exception as e:
            st.error(f"Gagal mengambil Google Sheets: {e}")
            st.info("Periksa link dan izin akses. Jika spreadsheet privat, integrasi perlu autentikasi Google.")
            st.stop()
    elif uploaded is not None:
        file_bytes = uploaded.getvalue()
        source_label = "Excel upload (fallback)"
    elif os.path.exists(local_path):
        with open(local_path, "rb") as f: file_bytes = f.read()
        st.sidebar.info("URL Google Sheets belum diisi. Saat ini memakai Excel bawaan.")
    else:
        st.info("Masukkan URL Google Sheets atau upload Excel untuk mulai.")
        st.stop()
else:
    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        source_label = "Excel upload"
    elif os.path.exists(local_path):
        with open(local_path, "rb") as f: file_bytes = f.read()
    else:
        st.info("Upload file Excel database pada panel kiri untuk mulai.")
        st.stop()

st.sidebar.caption(f"Sumber aktif: {source_label}")

try:
    master, cases, sheet_names = load_workbook(file_bytes)
except Exception as e:
    st.error(f"Gagal membaca Excel: {e}")
    st.stop()

st.sidebar.markdown("---")
page = st.sidebar.radio("MENU", [
    "Executive Dashboard", "Produktivitas PIC", "Monitoring Case",
    "Rekap Case Tahunan", "Daftar Case SC", "Daftar Case Bukan SC",
    "Kualitas Data Case", "Data & Panduan"
])

# Defensive references: the loader normally returns both dataframes above.
_master_df = locals().get("master", pd.DataFrame())
_cases_df = locals().get("cases", pd.DataFrame())
available_year_values = []
if isinstance(_master_df, pd.DataFrame) and "Tahun" in _master_df.columns:
    available_year_values.extend(_master_df["Tahun"].dropna().astype(int).tolist())
if isinstance(_cases_df, pd.DataFrame):
    for _year_col in ["Tahun", "Tahun Complaint"]:
        if _year_col in _cases_df.columns:
            available_year_values.extend(_cases_df[_year_col].dropna().astype(int).tolist())
available_years = sorted(set(available_year_values))
if not available_years: available_years = [date.today().year]
default_year = 2026 if 2026 in available_years else available_years[-1]
year = st.sidebar.selectbox("Tahun", available_years, index=available_years.index(default_year))
period_mode = st.sidebar.selectbox("Periode", ["Tahunan", "Bulanan", "Mingguan", "Rentang tanggal"])
month_choices = [f"{year}-{m:02d}" for m in range(1,13)]
selected_month = st.sidebar.selectbox("Bulan", month_choices, index=min(date.today().month-1,11)) if period_mode=="Bulanan" else None
selected_week = st.sidebar.text_input("Minggu ISO (contoh 2026-W04)", f"{year}-W{date.today().isocalendar().week:02d}") if period_mode=="Mingguan" else None
date_range = st.sidebar.date_input("Tanggal mulai & akhir", value=(date(year,1,1), date(year,12,31))) if period_mode=="Rentang tanggal" else None

def filter_period(df, date_col="Tanggal"):
    if df.empty: return df.copy()
    out = df[df["Tahun"] == year].copy()
    if period_mode == "Bulanan" and selected_month:
        out = out[out["Bulan"] == selected_month]
    elif period_mode == "Mingguan" and selected_week:
        out = out[out["Minggu"] == selected_week.strip()]
    elif period_mode == "Rentang tanggal" and date_range:
        if len(date_range) == 2:
            out = out[(out[date_col].dt.date >= date_range[0]) & (out[date_col].dt.date <= date_range[1])]
    return out

def fmt(n): return f"{n:,.0f}".replace(",", ".")

m = filter_period(master, "Tanggal Completed") if not master.empty else master
c = filter_period(cases, "Tanggal Diterima") if not cases.empty else cases
if not cases.empty:
    c = c[c["Tanggal Diterima"].notna()]

def metric_row(df, case_df):
    total_id = df["ID"].nunique() if not df.empty else 0
    sku = df["SKU"].sum() if not df.empty else 0
    qty = df["QTY"].sum() if not df.empty else 0
    active_pic = df["Nama PIC"].nunique() if not df.empty else 0
    total_case = len(case_df)
    sc = (case_df["Penyelesaian Case"].str.casefold() == "tanggungan sc").sum() if not case_df.empty else 0
    non_sc = (case_df["Penyelesaian Case"].str.casefold() == "bukan tanggungan sc").sum() if not case_df.empty else 0
    cols = st.columns(6)
    for col, label, value in zip(cols, ["Total ID","Total SKU","Total Qty","PIC Aktif","Case Masuk","Tanggungan SC"],
                                  [fmt(total_id),fmt(sku),fmt(qty),fmt(active_pic),fmt(total_case),fmt(sc)]):
        col.metric(label, value)
    st.caption(f"Bukan tanggungan SC: {fmt(non_sc)} case")

def productivity_table(df):
    if df.empty:
        st.warning("Tidak ada data produktivitas pada periode yang dipilih.")
        return pd.DataFrame()
    agg = df.groupby("Nama PIC").agg(
        Total_ID=("ID","nunique"), Total_SKU=("SKU","sum"), Total_Qty=("QTY","sum"),
        Hari_Aktif=("Tanggal","nunique")
    ).reset_index()
    agg["Rata_ID_per_Hari"] = agg["Total_ID"] / agg["Hari_Aktif"].clip(lower=1)
    agg["Rata_SKU_per_Hari"] = agg["Total_SKU"] / agg["Hari_Aktif"].clip(lower=1)
    agg["Rata_Qty_per_Hari"] = agg["Total_Qty"] / agg["Hari_Aktif"].clip(lower=1)
    return agg.sort_values("Total_Qty", ascending=False)

if page == "Executive Dashboard":
    st.subheader("Executive Overview")
    metric_row(m,c)
    left,right = st.columns(2)
    with left:
        st.markdown("### Tren produktivitas")
        if not m.empty:
            trend = m.groupby("Tanggal").agg(ID=("ID","nunique"), SKU=("SKU","sum"), Qty=("QTY","sum")).reset_index()
            fig = px.line(trend, x="Tanggal", y=["ID","SKU","Qty"], markers=True, title="Output berdasarkan tanggal completed")
            st.plotly_chart(fig, use_container_width=True)
    with right:
        st.markdown("### Komposisi tanggungan case")
        if not c.empty:
            cc = c["Penyelesaian Case"].fillna("Belum diklasifikasikan").value_counts().rename_axis("Klasifikasi").reset_index(name="Jumlah")
            st.plotly_chart(px.pie(cc, names="Klasifikasi", values="Jumlah", hole=.45), use_container_width=True)
    st.markdown("### Ranking PIC — berdasarkan Qty")
    p = productivity_table(m)
    if not p.empty:
        p.insert(0,"Ranking",range(1,len(p)+1))
        st.dataframe(p, use_container_width=True, hide_index=True)

elif page == "Produktivitas PIC":
    st.subheader("Produktivitas setiap PIC")
    p = productivity_table(m)
    if not p.empty:
        metric = st.selectbox("Urutkan ranking berdasarkan", ["Total_Qty","Total_ID","Total_SKU","Rata_Qty_per_Hari","Rata_ID_per_Hari","Rata_SKU_per_Hari"])
        p = p.sort_values(metric, ascending=False).reset_index(drop=True)
        p.insert(0,"Ranking",range(1,len(p)+1))
        st.dataframe(p, use_container_width=True, hide_index=True)
        chosen = st.selectbox("Detail PIC", sorted(p["Nama PIC"].dropna().unique()))
        detail = m[m["Nama PIC"] == chosen]
        daily = detail.groupby("Tanggal").agg(ID=("ID","nunique"),SKU=("SKU","sum"),Qty=("QTY","sum")).reset_index()
        st.markdown(f"### Tren harian: {chosen}")
        st.plotly_chart(px.bar(daily, x="Tanggal", y=["ID","SKU","Qty"], barmode="group"), use_container_width=True)
        st.dataframe(daily, use_container_width=True, hide_index=True)
        st.download_button("Download rekap produktivitas CSV", p.to_csv(index=False).encode("utf-8-sig"), "rekap_produktivitas_pic.csv", "text/csv")

elif page == "Monitoring Case":
    st.subheader("Monitoring case masuk")
    if c.empty:
        st.warning("Tidak ada case pada periode yang dipilih.")
    else:
        cat = c.groupby("Kategori", dropna=False).size().reset_index(name="Jumlah").sort_values("Jumlah",ascending=False)
        st.plotly_chart(px.bar(cat, x="Kategori", y="Jumlah", title="Case berdasarkan kategori"), use_container_width=True)
        trend = c.groupby("Tanggal Diterima").size().reset_index(name="Case")
        st.plotly_chart(px.line(trend, x="Tanggal Diterima", y="Case", markers=True, title="Tren case masuk"), use_container_width=True)
        st.dataframe(c[["Tanggal Diterima","Tanggal Complaint","ID","SKU","Qty Request","Qty Complaint","Kategori","Cabang","Keterangan Case","Penyelesaian Case","Nama PIC"]], use_container_width=True, hide_index=True)
        st.download_button("Download data case CSV", c.to_csv(index=False).encode("utf-8-sig"), "monitoring_case.csv", "text/csv")

elif page == "Rekap Case Tahunan":
    st.subheader(f"Rekap case bulanan tahun {year}")
    cy = cases[
        (cases["Tahun Complaint"] == year) & cases["Tanggal Complaint"].notna()
    ].copy() if not cases.empty else cases
    if cy.empty:
        st.warning("Belum ada data case berdasarkan Tanggal Complaint pada tahun ini.")
    else:
        cy["Klasifikasi"] = cy["Penyelesaian Case"].fillna("Belum diklasifikasikan").astype(str).str.strip()
        # Hitung jumlah baris case (bukan jumlah Kategori yang terisi) agar tidak ada baris terlewat.
        cy["Jumlah Baris"] = 1
        monthly = cy.pivot_table(
            index="Bulan Complaint", columns="Klasifikasi",
            values="Jumlah Baris", aggfunc="sum", fill_value=0
        ).reset_index().rename(columns={"Bulan Complaint": "Bulan"})
        for col in ["Tanggungan SC", "Bukan Tanggungan SC", "Belum diklasifikasikan"]:
            if col not in monthly:
                monthly[col] = 0
        # Tampilkan seluruh bulan Jan-Des, termasuk bulan tanpa case.
        all_months = pd.DataFrame({"Bulan": [f"{year}-{m:02d}" for m in range(1, 13)]})
        monthly = all_months.merge(monthly, on="Bulan", how="left").fillna(0)
        count_cols = [col for col in monthly.columns if col != "Bulan"]
        monthly[count_cols] = monthly[count_cols].astype(int)
        monthly["Total Case"] = monthly[count_cols].sum(axis=1)
        st.caption("Dasar rekap: Tanggal Complaint. Setiap baris pada sheet Case Pic dihitung sebagai satu case.")
        st.dataframe(monthly, use_container_width=True, hide_index=True)
        st.plotly_chart(px.bar(monthly, x="Bulan", y=["Tanggungan SC","Bukan Tanggungan SC"], barmode="group",
                               title="Akumulasi case per bulan (Tanggal Complaint)"), use_container_width=True)
        st.metric("Total case setahun", fmt(len(cy)))
        st.download_button("Download rekap tahunan CSV", monthly.to_csv(index=False).encode("utf-8-sig"), f"rekap_case_{year}.csv", "text/csv")
    st.markdown("### Pemeriksaan case di luar rekap tahun terpilih")
    no_complaint_date = cases[cases["Tanggal Complaint"].isna()].copy() if not cases.empty else cases
    prior_year = cases[
        cases["Tanggal Complaint"].notna() & (cases["Tanggal Complaint"].dt.year < year)
    ].copy() if not cases.empty else cases
    q1, q2 = st.columns(2)
    q1.metric("Case tanpa Tanggal Complaint", fmt(len(no_complaint_date)))
    q2.metric(f"Case sebelum {year}", fmt(len(prior_year)))
    with st.expander("Lihat case tanpa Tanggal Complaint"):
        if not no_complaint_date.empty:
            st.dataframe(no_complaint_date, use_container_width=True, hide_index=True)
            st.download_button("Download case tanpa tanggal", no_complaint_date.to_csv(index=False).encode("utf-8-sig"), "case_tanpa_tanggal_complaint.csv", "text/csv")
        else:
            st.success("Tidak ada case tanpa Tanggal Complaint.")
    with st.expander(f"Lihat case sebelum tahun {year}"):
        if not prior_year.empty:
            st.dataframe(prior_year, use_container_width=True, hide_index=True)
            st.download_button("Download case tahun sebelumnya", prior_year.to_csv(index=False).encode("utf-8-sig"), "case_tahun_sebelumnya.csv", "text/csv")
        else:
            st.success("Tidak ada case dari tahun sebelumnya.")

elif page in ["Daftar Case SC","Daftar Case Bukan SC"]:
    is_sc = page == "Daftar Case SC"
    st.subheader(page)
    if c.empty:
        st.warning("Tidak ada case pada periode yang dipilih.")
    else:
        target = "Tanggungan SC" if is_sc else "Bukan Tanggungan SC"
        out = c[c["Penyelesaian Case"].str.casefold() == target.casefold()].copy()
        if out.empty: st.info(f"Tidak ada data {target.lower()} pada periode ini.")
        else:
            st.metric("Jumlah case", fmt(len(out)))
            categories = ["Semua"] + sorted(out["Kategori"].dropna().unique().tolist())
            chosen_cat = st.selectbox("Filter kategori", categories)
            if chosen_cat != "Semua": out = out[out["Kategori"] == chosen_cat]
            st.dataframe(out[["Tanggal Diterima","ID","SKU","Qty Request","Qty Complaint","Kategori","Keterangan Tambahan","Cabang","Keterangan Case","Penyelesaian Case","Nama PIC"]], use_container_width=True, hide_index=True)
            st.download_button("Download daftar case CSV", out.to_csv(index=False).encode("utf-8-sig"), f"daftar_case_{'sc' if is_sc else 'non_sc'}.csv", "text/csv")

elif page == "Kualitas Data Case":
    st.subheader("Kualitas & kelengkapan data case")
    if cases.empty:
        st.warning("Sheet Case Pic belum memiliki data.")
    else:
        missing_date = cases[cases["Tanggal Complaint"].isna()].copy()
        dated = cases[cases["Tanggal Complaint"].notna()].copy()
        dated["Tahun Complaint"] = dated["Tanggal Complaint"].dt.year
        previous = dated[dated["Tahun Complaint"] < year].copy()
        future = dated[dated["Tahun Complaint"] > year].copy()
        unknown_class = cases[
            ~cases["Penyelesaian Case"].fillna("").str.strip().str.casefold().isin(
                ["tanggungan sc", "bukan tanggungan sc"]
            )
        ].copy()
        qa, qb, qc, qd = st.columns(4)
        qa.metric("Total baris Case Pic", fmt(len(cases)))
        qb.metric("Tanpa Tanggal Complaint", fmt(len(missing_date)))
        qc.metric(f"Tahun sebelum {year}", fmt(len(previous)))
        qd.metric("Belum/di luar klasifikasi", fmt(len(unknown_class)))
        st.markdown("### Distribusi tahun Tanggal Complaint")
        year_dist = dated.groupby("Tahun Complaint").size().reset_index(name="Jumlah Case").sort_values("Tahun Complaint")
        st.plotly_chart(px.bar(year_dist, x="Tahun Complaint", y="Jumlah Case", title="Case berdasarkan tahun complaint"), use_container_width=True)
        for label, frame, filename in [
            ("Case tanpa Tanggal Complaint", missing_date, "case_tanpa_tanggal_complaint.csv"),
            (f"Case sebelum tahun {year}", previous, "case_tahun_sebelumnya.csv"),
            (f"Case setelah tahun {year}", future, "case_tahun_setelahnya.csv"),
            ("Case belum/di luar klasifikasi", unknown_class, "case_belum_terklasifikasi.csv")
        ]:
            with st.expander(f"{label} ({len(frame):,} baris)"):
                if frame.empty:
                    st.success("Tidak ada data.")
                else:
                    st.dataframe(frame, use_container_width=True, hide_index=True)
                    st.download_button(f"Download {label}", frame.to_csv(index=False).encode("utf-8-sig"), filename, "text/csv", key=f"dl_{filename}")
        st.caption("Perbaiki tanggal dan klasifikasi di sumber spreadsheet. Dashboard akan membaca perubahan setelah refresh berikutnya.")

else:
    st.subheader("Data & Panduan")
    st.markdown("""
    **Sumber data:** sheet `Master 2025`, `Master 2026`, `Master 2027`, dan `Case Pic`.

    **Aturan perhitungan**
    - Total ID = jumlah ID unik pada periode terpilih.
    - Total SKU = penjumlahan kolom SKU.
    - Total Qty = penjumlahan kolom QTY.
    - Hari aktif PIC = jumlah tanggal berbeda saat PIC memiliki data completed.
    - Rata-rata per hari = total indikator dibagi hari aktif PIC.
    - Monitoring Case = baris pada `Case Pic` berdasarkan `Tanggal Diterima`.
    - Rekap Case Tahunan = baris pada `Case Pic` berdasarkan `Tanggal Complaint`.
    - Klasifikasi SC mengikuti kolom `Penyelesaian Case`.

    **Catatan kualitas data**
    - Baris produktivitas dengan tanggal completed kosong/tidak valid tidak dihitung.
    - Case dengan tanggal diterima kosong tidak masuk tren Monitoring Case.
    - Rekap tahunan hanya menghitung case yang memiliki Tanggal Complaint valid.
    - Klasifikasi case yang kosong atau berbeda ejaannya akan terlihat sebagai belum diklasifikasikan/tidak masuk daftar dua klasifikasi utama.
    - Hari aktif adalah hari kalender dengan aktivitas tercatat, bukan asumsi hari kerja Senin–Jumat.
    """)
    st.write("Sheet terdeteksi:", ", ".join(sheet_names))
    st.write(f"Baris produktivitas valid: {len(master):,} | Baris case: {len(cases):,}")
    st.caption(f"Sumber data aktif: {source_label}. Jika Google Sheets dipakai, dashboard mengambil ulang data setiap 5 menit (jika auto-refresh aktif), atau melalui tombol Refresh data sekarang.")
