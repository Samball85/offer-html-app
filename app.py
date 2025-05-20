import io
import pandas as pd
import openpyxl
import streamlit as st
import premailer
from openpyxl.utils import get_column_letter

# ─── Helpers ────────────────────────────────────────────────────────────────────
def format_value(val, fmt):
    if val is None:
        return ""
    try:
        if "£" in fmt: return f"£{float(val):,.2f}"
        if "$" in fmt: return f"${float(val):,.2f}"
        if "€" in fmt: return f"€{float(val):,.2f}"
        if isinstance(val, float): return f"{val:,.2f}"
    except:
        pass
    return str(val)

# ─── 1. Streamlit setup ─────────────────────────────────────────────────────────
st.set_page_config(layout="wide", page_title="Offer → Email-HTML")
st.title("Offer-to-Email-HTML Converter")

uploaded = st.file_uploader("Upload .xlsx", type="xlsx")
if not uploaded:
    st.info("Upload an Excel file to get started.")
    st.stop()
st.success("✅ File received. Parsing…")

# ─── 2. Load workbook & pick sheet ───────────────────────────────────────────────
file_bytes = uploaded.read()
wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
sheet_name = st.selectbox("Select sheet", wb.sheetnames)
ws = wb[sheet_name]

# ─── 3. Auto-detect header row ──────────────────────────────────────────────────
auto_header = 1
for i in range(1, 21):
    vals = [c.value for c in ws[i]]
    if sum(1 for v in vals if v not in (None, "")) > len(vals)/2:
        auto_header = i
        break

header_row = st.number_input(
    "Header row (detected):",
    min_value=1, max_value=20, value=auto_header
)

# ─── 4. Read sheet into pandas & clean ──────────────────────────────────────────
buf = io.BytesIO(file_bytes)
df = pd.read_excel(buf, sheet_name=sheet_name,
                   header=header_row-1, engine="openpyxl")
df.dropna(how="all", axis=0, inplace=True)
df.dropna(how="all", axis=1, inplace=True)
df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]

st.subheader("📊 Raw preview")
st.dataframe(df, use_container_width=True)

# ─── 5. Column picker ───────────────────────────────────────────────────────────
cols = st.multiselect(
    "Columns to include",
    options=list(df.columns),
    default=list(df.columns)
)
df_view = df[cols]
st.write(f"Using columns: {cols}")

# ─── 6. Currency-formatted preview ──────────────────────────────────────────────
header_cells = ws[header_row]
excel_headers = [c.value for c in header_cells]

display = []
for i, row in enumerate(df_view.itertuples(index=False, name=None)):
    excel_r = header_row + 1 + i
    line = []
    for j, _ in enumerate(row):
        name = df_view.columns[j]
        idx = excel_headers.index(name) + 1
        cell = ws.cell(row=excel_r, column=idx)
        line.append(format_value(cell.value, cell.number_format or ""))
    display.append(line)

preview_df = pd.DataFrame(display, columns=df_view.columns)
st.subheader("📊 Preview with currencies")
st.dataframe(preview_df, use_container_width=True)

# ─── 7. Column-colour overrides ─────────────────────────────────────────────────
st.markdown("---")
st.subheader("🎨 Pick your column colours")
override_cols = st.multiselect("Which columns to colour?", options=list(df_view.columns))
col_cols = {}
for c in override_cols:
    col_cols[c] = st.color_picker(f"Colour for '{c}'", "#dddddd", key=f"col_{c}")

# ─── 8. Generate & inline HTML ──────────────────────────────────────────────────
if st.button("Generate Brevo-Ready HTML"):
    st.info("Building HTML…")

    def build_html(ws, df, header_row):
        # compute widths
        widths = []
        for col in df.columns:
            idx = excel_headers.index(col) + 1
            letter = get_column_letter(idx)
            dim = ws.column_dimensions.get(letter)
            w = dim.width if dim and dim.width else 8.43
            widths.append(int(w * 7))

        total_w = sum(widths)
        html = (
            f'<table width="{total_w}px" '
            'style="border-collapse:collapse;'
                  'font-family:Arial,sans-serif;'
                  'font-size:12px;'
                  'table-layout:fixed;">'
        )

        # colgroup
        html += "<colgroup>"
        for px in widths:
            html += f'<col style="width:{px}px;" />'
        html += "</colgroup>"

        # header row (wrap only at spaces)
        html += "<tr>"
        for col in df.columns:
            bg = col_cols.get(col, "#f0f0f0")
            html += (
                '<th '
                'style="border:1px solid #ccc;'
                      'padding:6px;'
                      f'background:{bg};'
                      'font-weight:bold;'
                      'text-align:center;'
                      'white-space:normal;'
                      'overflow-wrap:normal;'
                      'word-break:normal;">'
                f"{col}</th>"
            )
        html += "</tr>"

        # data rows
        for i, row in enumerate(df.itertuples(index=False, name=None)):
            excel_r = header_row + 1 + i
            html += "<tr>"
            for j, _ in enumerate(row):
                col = df.columns[j]
                name = col
                idx = excel_headers.index(name) + 1
                cell = ws.cell(row=excel_r, column=idx)
                txt = format_value(cell.value, cell.number_format or "")
                bg = col_cols.get(col, "#ffffff")
                style = (
                    f"border:1px solid #ccc;"
                    f"padding:6px;"
                    f"background:{bg};"
                    "text-align:center;"
                )
                html += f'<td style="{style}">{txt}</td>'
            html += "</tr>"

        html += "</table>"
        return html

    raw = build_html(ws, preview_df, header_row)
    email = premailer.transform(raw)

    st.subheader("📧 Email-ready HTML Preview")
    st.components.v1.html(email, height=600, scrolling=True)

    st.subheader("🔗 Copy & paste the code")
    st.text_area("HTML code", email, height=300)

