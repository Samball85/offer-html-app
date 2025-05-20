import io
import requests
import pandas as pd
import openpyxl
import streamlit as st
import premailer
import matplotlib.pyplot as plt
from openpyxl.utils import get_column_letter

# ─── Helpers ────────────────────────────────────────────────────────────────────
def format_value(val, fmt):
    """
    Format a cell value according to its Excel number format:
    - Pounds (£)
    - Euros (€)
    - Dollars ($)
    - Plain floats
    """
    if val is None:
        return ""
    try:
        if "£" in fmt or "\u00a3" in fmt:
            return f"£{float(val):,.2f}"
        elif "€" in fmt or "\u20ac" in fmt:
            return f"€{float(val):,.2f}"
        elif "$" in fmt or "\u0024" in fmt:
            return f"${float(val):,.2f}"
        if isinstance(val, float):
            return f"{val:,.2f}"
    except:
        pass
    return str(val)

def url_ok(url):
    try:
        r = requests.head(url, allow_redirects=True, timeout=2)
        return r.status_code < 400
    except:
        return False

# ─── 1. Page config & mapping load ─────────────────────────────────────────────
st.set_page_config(layout="wide", page_title="Offer → Email-HTML")
st.title("Intamarques Offer to Email HTML Converter")

mapping_path = "mapping.csv"
try:
    mapping_df = pd.read_csv(mapping_path)
except Exception as e:
    st.error(f"⚠️ Could not read {mapping_path}: {e}")
    st.stop()
if not {"Code","Image URL"}.issubset(mapping_df.columns):
    st.error("⚠️ mapping.csv must have 'Code' and 'Image URL' columns")
    st.stop()
mapping_df = mapping_df.drop_duplicates(subset=["Code"], keep="first")

# ─── 2. Offer upload ───────────────────────────────────────────────────────────
uploaded = st.file_uploader("1) Upload your Offer .xlsx", type="xlsx")
if not uploaded:
    st.info("Please upload your offer to proceed to the next step.")
    st.stop()
wb = openpyxl.load_workbook(io.BytesIO(uploaded.read()), data_only=True)
sheet_name = st.selectbox("2) Select sheet", wb.sheetnames)
ws = wb[sheet_name]

# ─── 3. Header row detection ────────────────────────────────────────────────────
auto_header = 1
for i in range(1,21):
    vals = [c.value for c in ws[i]]
    if sum(1 for v in vals if v not in (None, "")) > len(vals)/2:
        auto_header = i
        break
header_row = st.number_input(
    "3) Header row (detected), use +/− to adjust (e.g. usually 6):",
    min_value=1, max_value=20, value=auto_header
)

# ─── 4. Read & clean data ──────────────────────────────────────────────────────
buf = io.BytesIO(uploaded.getvalue())
df = pd.read_excel(buf, sheet_name=sheet_name, header=header_row-1, engine="openpyxl")
df.dropna(how="all", axis=0, inplace=True)
df.dropna(how="all", axis=1, inplace=True)
df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]
st.subheader("4) Raw Offer Data")
st.dataframe(df, use_container_width=True)

# ─── 5. Column selection ───────────────────────────────────────────────────────
cols = st.multiselect(
    "5) Columns to include in email — remove those you don't need:",
    options=list(df.columns),
    default=list(df.columns)
)
if not cols:
    st.warning("Select at least one column.")
    st.stop()
df_view = df[cols]

# ─── 6. Image toggle ───────────────────────────────────────────────────────────
use_images = st.checkbox("Would you like to include product images?")
if use_images:
    merged = pd.merge(
        df_view,
        mapping_df[["Code","Image URL"]],
        how="left",
        on="Code",
        validate="many_to_one"
    )
else:
    merged = df_view.copy()
    merged["Image URL"] = ""

# ─── 7. Preview rows ──────────────────────────────────────────────────────────
excel_headers = [c.value for c in ws[header_row]]
preview_rows = []
for idx, row in enumerate(merged.itertuples(index=False, name=None)):
    excel_r = header_row + 1 + idx
    img_html = ""
    if use_images:
        url = row[len(cols)]
        if isinstance(url,str) and url_ok(url):
            img_html = f'<img src="{url}" style="height:60px; width:auto;" />'
    cells = [img_html] if use_images else []
    for j, col in enumerate(cols):
        val = ws.cell(row=excel_r, column=excel_headers.index(col)+1).value
        fmt = ws.cell(row=excel_r, column=excel_headers.index(col)+1).number_format
        cells.append(format_value(val, fmt))
    preview_rows.append(cells)
preview_cols = (['Image'] if use_images else []) + cols
preview_df = pd.DataFrame(preview_rows, columns=preview_cols)
sub_header = "7) Preview with Images" if use_images else "7) Preview"
st.subheader(sub_header)
st.write(preview_df.to_html(escape=False, index=False), unsafe_allow_html=True)

# ─── 8. Header colors ──────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("8) Pick header colours")
hdr = st.multiselect("Select which headers to color:", preview_cols)
col_mapping = {c: st.color_picker(f"Color for {c}", "#f0f0f0", key=c) for c in hdr}

# ─── 9. Generate HTML ──────────────────────────────────────────────────────────
if st.button("9) Generate Brevo-Ready HTML"):
    def build_html():
        widths = []
        if use_images:
            widths.append(80)
        for col in cols:
            letter = get_column_letter(excel_headers.index(col)+1)
            w = ws.column_dimensions[letter].width or 8.43
            widths.append(int(w*7))
        total = sum(widths)
        html = f'<table width="{total}px" style="border-collapse:collapse;font-family:Arial;table-layout:fixed;">'
        html += "<colgroup>" + "".join(f'<col style="width:{w}px"/>' for w in widths) + "</colgroup>"
        html += "<tr>"
        for c in preview_cols:
            bg = col_mapping.get(c, "#f0f0f0")
            html += f'<th style="border:1px solid #ccc;padding:6px;background:{bg};font-weight:bold;text-align:center;white-space:normal;word-break:keep-all;">{c}</th>'
        html += "</tr>"
        for row in preview_rows:
            html += "<tr>"
            for cell in row:
                html += f'<td style="border:1px solid #ccc;padding:6px;background:#fff;text-align:center;white-space:normal;word-break:keep-all;">{cell}</td>'
            html += "</tr>"
        html += "</table>"
        return premailer.transform(html)

    email_html = build_html()
    st.subheader("📧 Your Brevo-Ready HTML")
    st.components.v1.html(email_html, height=600, scrolling=True)
    st.text_area("Copy this HTML:", email_html, height=300)

    # ─── 10. Downloads ──────────────────────────────────────────────────────────
    dl_cols = ["Image URL"] + cols if use_images else cols
    dl_df = merged[dl_cols].copy()
    if use_images:
        dl_df.rename(columns={"Image URL":"Image URL (CDN)"}, inplace=True)

    buf = io.BytesIO()
    dl_df.to_excel(buf, index=False, sheet_name="Offer")
    buf.seek(0)
    st.download_button("⬇️ Download as Excel", data=buf, file_name="offer.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    jpeg_df = dl_df.drop(columns=["Image URL (CDN)"]) if use_images else dl_df
    fig, ax = plt.subplots(figsize=(len(jpeg_df.columns)*1.2, max(2,len(jpeg_df)*0.5)), dpi=150)
    ax.axis('off')
    tbl = ax.table(cellText=jpeg_df.values.tolist(), colLabels=jpeg_df.columns.tolist(), cellLoc='center', loc='center')
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(10)
    fig.tight_layout(pad=1)
    imgb = io.BytesIO(); fig.savefig(imgb, format='jpeg')
    st.download_button("⬇️ Download as JPEG", data=imgb.getvalue(), file_name="offer.jpg", mime="image/jpeg")

