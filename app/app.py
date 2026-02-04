import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="CSV Iklan → Excel Berwarna",
    layout="centered"
)

st.title("📊 CSV Iklan → Excel Berwarna")
st.caption("Upload CSV iklan Shopee → otomatis rapi → download Excel laporan")

# =========================
# UPLOAD CSV
# =========================
uploaded_file = st.file_uploader(
    "Upload file CSV iklan Shopee",
    type=["csv"]
)

csv_mode = st.radio(
    "Jenis CSV yang di-upload",
    options=[
        "CSV Keseluruhan (Normal)",
        "CSV Grup Iklan (hanya iklan produk)"
    ],
    horizontal=True
)

# =========================
# LOAD CSV
# =========================
@st.cache_data
def load_uploaded_csv(file):
    file.seek(0)
    raw = file.read().decode("utf-8", errors="ignore")
    lines = raw.splitlines()

    HEADER_KEYS = ["Nama Iklan", "Nama Iklan/Produk"]
    header_idx = None

    for i, line in enumerate(lines[:30]):
        if any(k in line for k in HEADER_KEYS):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Header Nama Iklan tidak ditemukan")

    delimiter = ";" if lines[header_idx].count(";") > lines[header_idx].count(",") else ","
    clean_csv = "\n".join(lines[header_idx:])

    df = pd.read_csv(
        io.StringIO(clean_csv),
        sep=delimiter,
        engine="python",
        on_bad_lines="skip"
    )

    df.columns = df.columns.str.strip()
    return df

# =========================
# NORMALISASI NAMA IKLAN
# =========================
def normalize_nama_iklan_column(df):
    for col in ["Nama Iklan", "Nama Iklan/Produk"]:
        if col in df.columns:
            return df.rename(columns={col: "Nama Iklan"})
    raise ValueError("Kolom Nama Iklan tidak ditemukan")

# =========================
# PEMENDEK NAMA IKLAN (RINGKASAN SAJA)
# =========================
def short_nama_iklan(nama):
    if pd.isna(nama):
        return nama

    text = str(nama).strip()

    # 1️⃣ KUNCI GRUP IKLAN (STOP DI SINI)
    if text.lower().startswith("grup iklan"):
        return text.split(" - ")[0]

    # 2️⃣ buang tag [SB], [TEST], dll
    text = re.sub(r"\[.*?\]", "", text).strip()
    
    feature_blacklist = {
        "busui","friendly","bahan","soft","ultimate","ultimates",
        "motif","size","ukuran","promo","diskon","broad","testing",
        "rayon","katun","cotton","silk","sustra","viscose",
        "linen","polyester","jersey","crepe","chiffon",
        "woolpeach","baloteli","babyterry",
        "pink","hitam","black","putih","white","navy","biru","blue",
        "merah","red","hijau","green","coklat","brown",
        "abu","abu-abu","grey","gray","cream","krem","beige",
        "maroon","ungu","purple","tosca","olive","sage"
    }

    store_blacklist = {
        "official","shop","store","boutique","fashion",
        "my","zahir","myzahir","by","original","premium"
    }

    category_keywords = {
        "gamis","dress","tunik","abaya","set",
        "blouse","khimar","rok","pashmina","hijab","outer",
    }

    context_blacklist = {
        "terbaru","new","update","launch","launching",
        "viral","hits","best","seller","bestseller",
        "kondangan","lebaran","ramadhan","ramadan",
        "harian","pesta","formal","casual",
        "trend","trending","populer",
        "2024","2025","2026","2027", "2028", "2029", "2030"
    }

    parts = re.split(r"\s*[-|]\s*", text)

    # Prioritise product-like parts
    product_keywords = {"dress", "gamis", "set"}
    product_candidates = []

    for part in parts:
        words = part.split()
        words_lower = [w.lower() for w in words]

        if not any(w in product_keywords for w in words_lower):
            continue

        # drop store prefix
        while words_lower and words_lower[0] in store_blacklist:
            words_lower.pop(0)
            words.pop(0)

        unique_words = [
            w for w in words_lower
            if w not in store_blacklist
            and w not in feature_blacklist
            and w not in context_blacklist
            and w not in category_keywords
        ]

        if unique_words:
            product_candidates.append(words)

    if product_candidates:
        best_words = product_candidates[-1]
        return " ".join(best_words[:3])

    # fallback scoring
    def score(part):
        s = 0
        for w in part.lower().split():
            if w in store_blacklist:
                s -= 3
            elif w in feature_blacklist:
                s -= 1
            elif w in context_blacklist:
                s -= 2
            elif w in category_keywords:
                s += 1
            else:
                s += 3
        return s

    best = max(parts, key=score)
    return " ".join(best.split()[:3])

# =========================
# STYLING DATA_IKLAN (for Data sheet)
# =========================
def highlight_row(row):
    styles = [''] * len(row)

    roas = row.get('Efektifitas Iklan')
    sales = row.get('Produk Terjual')
    gmv = row.get('Penjualan Langsung (GMV Langsung)')
    cost = row.get('Biaya')

    # safety: if sales or cost missing -> can't decide special cases
    if pd.isna(sales) or pd.isna(cost):
        return styles

    # 🟢 HIJAU TIPE A — cost == 0 & sales > 0 -> dark green text (apply even if ROAS NaN)
    if (cost == 0) and (sales > 0):
        return ['color: #006400'] * len(row)

    # 🔴 MERAH TIPE A — rugi keras (text red)
    if sales == 0 and cost >= 10000:
        return ['color: #FF0000'] * len(row)

    # ⚪ NETRAL — pemanasan
    if sales == 0 and cost < 10000:
        return styles

    # 🟥🟨🟩 WARNA ROAS (background) — only if ROAS is numeric
    if pd.notna(roas):
        try:
            if roas < 8:
                styles = ['background-color: red'] * len(row)
            elif roas < 10:
                styles = ['background-color: yellow'] * len(row)
            else:
                styles = ['background-color: lightgreen'] * len(row)
        except Exception:
            pass

    # 🔵 OVERLAY BIRU — assist only (only highlight Nama Iklan & GMV col)
    try:
        nama_idx = row.index.get_loc('Nama Iklan')
    except Exception:
        nama_idx = None
    try:
        gmv_idx = row.index.get_loc('Penjualan Langsung (GMV Langsung)')
    except Exception:
        gmv_idx = None

    if sales > 0 and (pd.isna(gmv) or gmv == 0):
        if nama_idx is not None:
            styles[nama_idx] = 'background-color: lightblue'
        if gmv_idx is not None:
            styles[gmv_idx] = 'background-color: lightblue'

    return styles

# =========================
# KATEGORI IKLAN (for ringkasan)
# =========================
def get_iklan_color(row, csv_mode):
    roas = row.get('Efektifitas Iklan')
    sales = row.get('Produk Terjual')
    cost = row.get('Biaya')

    # safety
    if pd.isna(sales) or pd.isna(cost):
        return None

    # HIJAU TIPE A => exclude from ringkasan (go to its own sheet)
    if (cost == 0) and (sales > 0):
        return None

    # TANPA KONVERSI BESAR → exclude from ringkasan
    if sales == 0 and cost >= 10000:
        return None

    # pemanasan
    if sales == 0 and cost < 10000:
        return None

    # CSV Grup mode: if ROAS not present then fallback to sales>0 => HIJAU
    if csv_mode == "CSV Grup Iklan (hanya iklan produk)":
        if pd.isna(roas):
            return "HIJAU" if sales > 0 else None

    # ROAS-based
    if pd.isna(roas) or roas < 8:
        return "MERAH"
    elif roas < 10:
        return "KUNING"
    else:
        return "HIJAU"

# =========================
# PROCESS & EXPORT
# =========================
if uploaded_file:
    if st.button("🚀 Proses & Download Excel"):
        with st.spinner("Memproses data..."):
            # load + normalize
            df = load_uploaded_csv(uploaded_file)
            df = normalize_nama_iklan_column(df)

            # mark aggregate/group rows (baris yang mulai dengan "grup")
            df["IS_AGGREGATE"] = df["Nama Iklan"].astype(str).str.lower().str.match(r'^\s*grup\b')

            # convert numerik aman
            for col in [
                "Efektifitas Iklan",
                "Produk Terjual",
                "Penjualan Langsung (GMV Langsung)",
                "Biaya"
            ]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

            # ===== compute HIJAU_TIPE_A robustly from FULL df (not from df_nonagg) =====
            # require cost not-NaN and exactly zero, and sales > 0
            df["IS_HIJAU_TIPE_A"] = (
                df.get("Biaya").notna() &
                (df.get("Biaya") == 0) &
                (df.get("Produk Terjual") > 0)
            )

            # IS_BIRU (assist rows)
            df["IS_BIRU"] = (
                (df.get("Produk Terjual", 0) > 0) &
                (df.get("Penjualan Langsung (GMV Langsung)", 0) == 0)
            )

            # Nama Ringkasan: jangan rubah Nama Iklan asli for aggregate rows
            df["Nama Ringkasan"] = df["Nama Iklan"].where(
                df["IS_AGGREGATE"],
                df["Nama Iklan"].apply(short_nama_iklan)
            )

            # Kategori (apply with csv_mode) — this returns None for rows excluded from ringkasan
            df["Kategori"] = df.apply(lambda row: get_iklan_color(row, csv_mode), axis=1)

            # BUILD df_nonagg used for RINGKASAN and >10K:
            if csv_mode == "CSV Grup Iklan (hanya iklan produk)":
                df_nonagg = df[~df["IS_AGGREGATE"]].copy()
            else:
                df_nonagg = df.copy()

            # ALWAYS exclude HIJAU_TIPE_A from ringkasan/tanpa_konversi logic
            df_nonagg = df_nonagg[~df_nonagg["IS_HIJAU_TIPE_A"]].copy()

            # build ordered list for numbering (MERAH->KUNING->HIJAU->BIRU) from df_nonagg
            ordered_for_numbering = []
            for kat in ["MERAH", "KUNING", "HIJAU"]:
                for name in df_nonagg[df_nonagg["Kategori"] == kat]["Nama Ringkasan"]:
                    ordered_for_numbering.append({"nama": name, "kategori": kat})
            for name in df_nonagg[df_nonagg["IS_BIRU"]]["Nama Ringkasan"]:
                ordered_for_numbering.append({"nama": name, "kategori": "BIRU"})

            # Build per-col content depending on csv_mode
            per_col = {"MERAH": [], "KUNING": [], "HIJAU": [], "BIRU": []}

            if csv_mode == "CSV Keseluruhan (Normal)":
                # numbered vertical lists (global numbering across all categories)
                for idx, item in enumerate(ordered_for_numbering, start=1):
                    numbered = f"{idx}. {item['nama']}"
                    per_col[item["kategori"]].append(numbered)
            else:
                # CSV Grup Iklan: comma-separated with trailing commas (single-cell-per-color)
                for kat in ["MERAH", "KUNING", "HIJAU"]:
                    names = df_nonagg[df_nonagg["Kategori"] == kat]["Nama Ringkasan"].tolist()
                    per_col[kat] = [f"{n}," for n in names]
                names_biru = df_nonagg[df_nonagg["IS_BIRU"]]["Nama Ringkasan"].tolist()
                per_col["BIRU"] = [f"{n}," for n in names_biru]

            # >10K tanpa konversi (use df_nonagg) — consistent with mode
            tanpa_konversi_df = (
                df_nonagg[(df_nonagg.get("Produk Terjual", 0) == 0) & (df_nonagg.get("Biaya", 0) >= 10000)]
                [["Nama Ringkasan", "Biaya"]]
                .rename(columns={"Nama Ringkasan": "Nama Iklan"})
                .sort_values("Biaya", ascending=False)
            )

            # HIJAU TIPE A sheet: ALWAYS computed from FULL df (so it shows up for both modes)
            hijau_cols = ["Nama Ringkasan", "Produk Terjual", "Efektifitas Iklan", "Biaya"]
            available_cols = [c for c in hijau_cols if c in df.columns]
            hijau_tipe_a_df = df[(df.get("Biaya").notna()) & (df.get("Biaya") == 0) & (df.get("Produk Terjual", 0) > 0)][available_cols].copy()
            if "Nama Ringkasan" in hijau_tipe_a_df.columns:
                hijau_tipe_a_df = hijau_tipe_a_df.rename(columns={"Nama Ringkasan": "Nama Iklan"})

            # =========================
            # EXPORT EXCEL
            # =========================
            buffer = io.BytesIO()
            original_name = uploaded_file.name
            base_name = original_name.rsplit(".", 1)[0]
            filename = f"{base_name}.xlsx"

            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                # Sheet 1 — DATA (keep original Nama Iklan column; includes aggregates)
                df.style.apply(highlight_row, axis=1).to_excel(
                    writer, sheet_name="DATA_IKLAN", index=False
                )

                # create RINGKASAN_IKLAN sheet manually
                wb = writer.book
                from openpyxl.styles import Font, Alignment

                if "RINGKASAN_IKLAN" in wb.sheetnames:
                    wb.remove(wb["RINGKASAN_IKLAN"])
                ws_ring = wb.create_sheet("RINGKASAN_IKLAN")

                headers = ["MERAH", "KUNING", "HIJAU", "BIRU"]
                color_map = {
                    "MERAH": "FF0000",
                    "KUNING": "000000",
                    "HIJAU": "00AA00",
                    "BIRU": "0066CC"
                }

                # write headers
                for c_idx, h in enumerate(headers, start=1):
                    cell = ws_ring.cell(row=1, column=c_idx, value=h)
                    cell.font = Font(bold=True)

                # write content depending on mode
                if csv_mode == "CSV Keseluruhan (Normal)":
                    # put vertical numbered list in each color's single cell using newline
                    for c_idx, key in enumerate(headers, start=1):
                        items = per_col.get(key, [])
                        if items:
                            text = "\n".join(items)
                            cell = ws_ring.cell(row=2, column=c_idx, value=text)
                            cell.font = Font(color=color_map[key])
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
                        else:
                            ws_ring.cell(row=2, column=c_idx, value="")
                else:
                    # CSV Grup Iklan: comma-separated in single cell (row 2), each item with trailing comma
                    for c_idx, key in enumerate(headers, start=1):
                        items = per_col.get(key, [])
                        if items:
                            joined = " ".join(items)
                            if not joined.strip().endswith(","):
                                joined = joined + ","
                            cell = ws_ring.cell(row=2, column=c_idx, value=joined)
                            cell.font = Font(color=color_map[key])
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
                        else:
                            ws_ring.cell(row=2, column=c_idx, value="")

                # adjust column widths
                for i in range(1, 5):
                    ws_ring.column_dimensions[chr(64 + i)].width = 40

                # Sheet: >10K_TANPA_KONVERSI (from df_nonagg)
                tanpa_konversi_df.to_excel(
                    writer,
                    sheet_name=">10K_TANPA_KONVERSI",
                    index=False
                )
                ws_tc = writer.book[">10K_TANPA_KONVERSI"]
                for r in range(2, ws_tc.max_row + 1):
                    for c in range(1, ws_tc.max_column + 1):
                        cell = ws_tc.cell(row=r, column=c)
                        cell.font = Font(color="FF0000")

                # NEW sheet: SALES_0_BIAYA (HIJAU TIPE A) — always from full df
                # --> write even if empty so sheet exists
                hijau_tipe_a_df.to_excel(writer, sheet_name="SALES_0_BIAYA", index=False)
                ws_hi = writer.book["SALES_0_BIAYA"]
                # color dark green text if there are rows (if header-only, loop won't run)
                from openpyxl.styles import Font as _Font
                for r in range(2, ws_hi.max_row + 1):
                    for c in range(1, ws_hi.max_column + 1):
                        cell = ws_hi.cell(row=r, column=c)
                        cell.font = _Font(color="006400")  # dark green text

            buffer.seek(0)

        st.success("Excel laporan siap di-download 👇")

        st.download_button(
            "Download Excel Laporan",
            buffer,
            filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
