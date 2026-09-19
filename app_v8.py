# Library Sticker Generator - Streamlit

import re
from io import BytesIO, StringIO
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st
import fitz  # PyMuPDF - renders PDF pages for the in-app preview
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code128

st.set_page_config(page_title="Library Sticker Generator", layout="wide")

# -----------------------------
# File / data helpers
# -----------------------------

def google_sheet_id_and_gid(url):
    parsed = urlparse(url.strip())
    if parsed.netloc not in {"docs.google.com", "sheets.google.com"}:
        raise ValueError("Please enter a Google Sheets link.")
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", parsed.path)
    if not match:
        raise ValueError("Invalid Google Sheets link. Expected a Google Sheets URL.")
    sheet_id = match.group(1)
    gid = parse_qs(parsed.query).get("gid", [None])[0]
    if not gid and parsed.fragment:
        m = re.search(r"(?:^|&)gid=([0-9]+)", parsed.fragment)
        if m:
            gid = m.group(1)
    return sheet_id, gid


def download_google_sheet(url):
    sheet_id, gid = google_sheet_id_and_gid(url)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    request = Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urlopen(request, timeout=30) as response:
            data = response.read()
    except Exception as e:
        raise ValueError(
            "Google Sheet could not be downloaded. Make sure it is accessible "
            "to the app (for example Anyone with the link → Viewer). "
            f"Details: {e}"
        )
    if not data.startswith(b"PK"):
        raise ValueError("Google did not return an Excel file. Check sharing/link settings.")
    return BytesIO(data), f"google_sheet_{sheet_id}.xlsx", gid


def get_sheet_names_from_google_link(url):
    workbook, _, _ = download_google_sheet(url)
    return pd.ExcelFile(workbook, engine="openpyxl").sheet_names


def read_google_sheet(url, sheet_name, has_header=True):
    workbook, _, _ = download_google_sheet(url)
    return pd.read_excel(
        workbook,
        sheet_name=sheet_name,
        dtype=str,
        keep_default_na=False,
        header=0 if has_header else None,
        engine="openpyxl",
    )


def get_sheet_names(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return ["CSV"]
    raw = uploaded.getvalue()
    if name.endswith(".xlsx"):
        return pd.ExcelFile(BytesIO(raw), engine="openpyxl").sheet_names
    if name.endswith(".ods"):
        return pd.ExcelFile(BytesIO(raw), engine="odf").sheet_names
    raise ValueError("Unsupported file type.")


def read_uploaded_table(uploaded, sheet_name=None, has_header=True):
    name = uploaded.name.lower()
    header = 0 if has_header else None
    raw = uploaded.getvalue()
    if name.endswith(".csv"):
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                decoded = raw.decode(encoding)
                return pd.read_csv(
                    StringIO(decoded), dtype=str, keep_default_na=False, header=header
                )
            except (UnicodeDecodeError, pd.errors.ParserError) as e:
                last_error = e
        raise ValueError(f"Could not decode/read the CSV file. Last error: {last_error}")
    if name.endswith(".xlsx"):
        return pd.read_excel(
            BytesIO(raw), sheet_name=sheet_name, dtype=str,
            keep_default_na=False, header=header, engine="openpyxl"
        )
    if name.endswith(".ods"):
        return pd.read_excel(
            BytesIO(raw), sheet_name=sheet_name, dtype=str,
            keep_default_na=False, header=header, engine="odf"
        )
    raise ValueError("Unsupported file type.")


def get_column_options(df):
    options = []
    for i, col in enumerate(df.columns):
        label = f"Column {i + 1}" if isinstance(col, int) else str(col).strip() or f"Column {i + 1}"
        options.append((i, f"{i + 1}: {label}"))
    return options


def safe_base_name(filename):
    base = filename.rsplit(".", 1)[0]
    base = re.sub(r'[<>:"/\\|?*]+', "_", base).strip()
    return base or "stickers"


def source_base_name(uploaded=None, manual_name="manual_data"):
    if uploaded:
        return safe_base_name(uploaded.name)
    return safe_base_name(manual_name)


def output_filename(base, suffix):
    return f"{source_base_name(None, base)} ({suffix}).pdf"


def clean_dataframe(df):
    df = df.copy()
    if df.empty:
        return df
    df = df.fillna("").astype(str)
    # Keep rows, including completely empty rows, because they may be reserved.
    return df


def show_data_preview(df, key_prefix):
    st.subheader("Data Preview")
    if df is None or df.empty:
        st.warning("No data rows found.")
        return
    preview_rows = st.number_input(
        "Preview rows", min_value=1, max_value=min(100, max(1, len(df))),
        value=min(10, len(df)), key=f"{key_prefix}_preview_rows"
    )
    st.dataframe(df.head(int(preview_rows)), use_container_width=True, height=260)
    st.caption(f"{len(df)} data row(s) • {len(df.columns)} column(s)")


def show_google_link_settings(url, key_prefix, default_header=True):
    sheet_names = get_sheet_names_from_google_link(url)
    sheet_name = st.selectbox("Select Sheet", sheet_names, key=f"{key_prefix}_google_sheet")
    has_header = st.checkbox(
        "First row contains column headings", value=default_header,
        key=f"{key_prefix}_google_header",
        help="Turn this off when the first row is actual library data."
    )
    df = read_google_sheet(url, sheet_name, has_header)
    return clean_dataframe(df), has_header, sheet_name


def show_file_settings(uploaded, key_prefix, default_header=True):
    sheet_names = get_sheet_names(uploaded)
    if len(sheet_names) > 1:
        sheet_name = st.selectbox("Select Sheet", sheet_names, key=f"{key_prefix}_sheet")
    else:
        sheet_name = sheet_names[0]
        if sheet_name != "CSV":
            st.caption(f"Sheet: {sheet_name}")
    has_header = st.checkbox(
        "First row contains column headings", value=default_header,
        key=f"{key_prefix}_header",
        help="Turn this off when the first row is actual library data."
    )
    df = read_uploaded_table(
        uploaded, sheet_name=None if sheet_name == "CSV" else sheet_name, has_header=has_header
    )
    return clean_dataframe(df), has_header, sheet_name


def show_manual_data_editor(key_prefix):
    st.subheader("Enter Data Manually")
    c1, c2 = st.columns(2)
    with c1:
        manual_rows = st.number_input(
            "Rows", min_value=1, max_value=500, value=10, step=1,
            key=f"{key_prefix}_manual_rows"
        )
    with c2:
        manual_cols = st.number_input(
            "Columns", min_value=1, max_value=30, value=5, step=1,
            key=f"{key_prefix}_manual_cols"
        )

    use_headers = st.checkbox(
        "Use column headings", value=True, key=f"{key_prefix}_manual_headers",
        help="When enabled, enter/edit the column names in the heading fields below."
    )

    if use_headers:
        names = []
        name_cols = st.columns(min(int(manual_cols), 4))
        for i in range(int(manual_cols)):
            with name_cols[i % len(name_cols)]:
                names.append(st.text_input(
                    f"Column {i + 1} name", value=f"Column {i + 1}",
                    key=f"{key_prefix}_manual_colname_{i}"
                ))
        columns = [n.strip() or f"Column {i + 1}" for i, n in enumerate(names)]
    else:
        columns = [f"Column {i + 1}" for i in range(int(manual_cols))]

    old = st.session_state.get(f"{key_prefix}_manual_grid")
    if isinstance(old, pd.DataFrame) and len(old) == int(manual_rows) and len(old.columns) == int(manual_cols):
        data = old.copy()
        data.columns = columns
    else:
        data = pd.DataFrame("", index=range(int(manual_rows)), columns=columns)

    edited = st.data_editor(
        data,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        key=f"{key_prefix}_manual_grid_editor",
        column_config={col: st.column_config.TextColumn(col) for col in columns},
    )
    st.session_state[f"{key_prefix}_manual_grid"] = edited
    st.caption("Tip: paste a block copied from Excel directly into the grid.")
    return clean_dataframe(edited), use_headers, "Manual Data"


def render_source_selector(key_prefix, allowed_upload_types=("csv", "xlsx", "ods")):
    source_type = st.radio(
        "Data Source",
        ["Upload File", "Google Sheets Link", "Enter Data Manually"],
        horizontal=True,
        key=f"{key_prefix}_source_type",
    )
    uploaded = None
    google_url = ""
    df = None
    has_header = True
    sheet_name = None
    manual_name = "manual_data"

    try:
        if source_type == "Upload File":
            uploaded = st.file_uploader(
                "Upload CSV / Excel / ODS", type=list(allowed_upload_types),
                key=f"{key_prefix}_upload"
            )
            if uploaded:
                df, has_header, sheet_name = show_file_settings(uploaded, key_prefix)
        elif source_type == "Google Sheets Link":
            google_url = st.text_input(
                "Google Sheets Link",
                placeholder="https://docs.google.com/spreadsheets/d/...",
                key=f"{key_prefix}_google_url",
                help="The sheet must be accessible to the app, for example Anyone with the link → Viewer."
            )
            if google_url.strip():
                df, has_header, sheet_name = show_google_link_settings(google_url, key_prefix)
        else:
            manual_name = st.text_input(
                "Output file base name", "manual_data", key=f"{key_prefix}_manual_filename",
                help="For example: library_books"
            )
            df, has_header, sheet_name = show_manual_data_editor(key_prefix)
    except Exception as e:
        st.error(f"Could not read the data source: {e}")
        return source_type, uploaded, google_url, None, has_header, sheet_name, manual_name

    return source_type, uploaded, google_url, df, has_header, sheet_name, manual_name

# -----------------------------
# PDF helpers
# -----------------------------

def grid_position(index, cols, rows, start_col, start_row):
    per_page = cols * rows
    absolute = index + start_row * cols + start_col
    page = absolute // per_page
    slot = absolute % per_page
    col = slot % cols
    row = slot // cols
    return page, col, row


def draw_barcode_label(c, x, y, label_width, label_height, library_name, library_place, number):
    c.setFont("Helvetica-Bold", 7)
    c.drawCentredString(x + label_width / 2, y + 17 * mm, library_name)
    c.drawCentredString(x + label_width / 2, y + 14 * mm, library_place)
    barcode = code128.Code128(str(number), barHeight=8 * mm, barWidth=0.53 * mm, humanReadable=False)
    barcode_x = x + (label_width - barcode.width) / 2
    barcode.drawOn(c, barcode_x, y + 4.5 * mm)
    c.drawCentredString(x + label_width / 2, y + 2 * mm, str(number))


def generate_barcode_pdf(df, library_name, library_place, cols, rows, label_width_mm,
                         label_height_mm, left_margin_mm, top_margin_mm, col_spacing_mm,
                         row_spacing_mm, start_col, start_row, barcode_column,
                         preserve_empty_rows=True):
    if barcode_column < 0 or barcode_column >= len(df.columns):
        raise ValueError("The selected barcode column does not exist.")
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_width, page_height = A4
    lw, lh = label_width_mm * mm, label_height_mm * mm
    lm, tm = left_margin_mm * mm, top_margin_mm * mm
    cg, rg = col_spacing_mm * mm, row_spacing_mm * mm
    per_page = cols * rows
    values = df.iloc[:, barcode_column].tolist()

    last_page = -1
    if preserve_empty_rows:
        drawable_values = list(enumerate(values))
    else:
        drawable_values = [(i, v) for i, v in enumerate(values) if str(v).strip()]
    for sequence_i, (source_i, value) in enumerate(drawable_values):
        position_i = source_i if preserve_empty_rows else sequence_i
        page, col, row = grid_position(position_i, cols, rows, start_col, start_row)
        if page != last_page:
            if last_page >= 0:
                c.showPage()
            last_page = page
        text = "" if value is None else str(value).strip()
        if not text:
            continue  # reserved blank cell: position is consumed, sticker remains empty
        x = lm + col * (lw + cg)
        y_top = page_height - tm - row * (lh + rg)
        y = y_top - lh
        draw_barcode_label(c, x, y, lw, lh, library_name, library_place, text)

    c.save()
    buffer.seek(0)
    return buffer


def draw_side_label(pdf, x, y, label_width, data, font_size=12):
    pdf.setFont("Times-Bold", font_size)
    positions = [17.5 * mm, 13 * mm, 8.5 * mm, 4 * mm]
    for i in range(4):
        text = data[i] if i < len(data) else ""
        pdf.drawCentredString(x + label_width / 2, y + positions[i], text)


def generate_side_pdf(df, selected_columns, cols, rows, label_width_mm, label_height_mm,
                      left_margin_mm, top_margin_mm, col_spacing_mm, row_spacing_mm,
                      start_col, start_row, font_size, preserve_empty_rows=True):
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    page_width, page_height = A4
    lw, lh = label_width_mm * mm, label_height_mm * mm
    lm, tm = left_margin_mm * mm, top_margin_mm * mm
    cg, rg = col_spacing_mm * mm, row_spacing_mm * mm
    last_page = -1
    source_rows = list(df.iterrows())
    if not preserve_empty_rows:
        filtered = []
        for source_i, row_data in source_rows:
            vals = ["" if pd.isna(row_data.iloc[ci]) else str(row_data.iloc[ci]) for ci in selected_columns]
            if any(v.strip() for v in vals):
                filtered.append((source_i, row_data))
        source_rows = filtered

    for sequence_i, (source_i, row_data) in enumerate(source_rows):
        position_i = sequence_i if not preserve_empty_rows else source_i
        page, col, row = grid_position(position_i, cols, rows, start_col, start_row)
        if page != last_page:
            if last_page >= 0:
                pdf.showPage()
            last_page = page
        data = []
        for column_index in selected_columns:
            value = row_data.iloc[column_index] if column_index is not None else ""
            data.append("" if pd.isna(value) else str(value))
        if not any(v.strip() for v in data):
            continue
        x = lm + col * (lw + cg)
        y_top = page_height - tm - row * (lh + rg)
        y = y_top - lh
        draw_side_label(pdf, x, y, lw, data, font_size)

    pdf.save()
    buffer.seek(0)
    return buffer


def make_combined_pdf(df, barcode_column, side_columns, library_name, library_place,
                      barcode_copies, side_copies, cols, rows, label_width_mm,
                      label_height_mm, left_margin_mm, top_margin_mm, col_spacing_mm,
                      row_spacing_mm, start_col, start_row, preserve_empty_rows=True):
    items = []
    source_rows = list(df.iterrows())
    if not preserve_empty_rows:
        source_rows = [
            (i, row) for i, row in source_rows
            if str(row.iloc[barcode_column]).strip() or any(str(row.iloc[ci]).strip() for ci in side_columns)
        ]
    for _, row in source_rows:
        barcode = "" if pd.isna(row.iloc[barcode_column]) else str(row.iloc[barcode_column]).strip()
        side_data = []
        for ci in side_columns:
            value = row.iloc[ci]
            side_data.append("" if pd.isna(value) else str(value))
        for _ in range(int(barcode_copies)):
            items.append({"type": "barcode", "barcode": barcode})
        for _ in range(int(side_copies)):
            items.append({"type": "side", "data": side_data})

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    page_width, page_height = A4
    lw, lh = label_width_mm * mm, label_height_mm * mm
    lm, tm = left_margin_mm * mm, top_margin_mm * mm
    cg, rg = col_spacing_mm * mm, row_spacing_mm * mm
    last_page = -1

    for i, item in enumerate(items):
        page, col, row = grid_position(i, cols, rows, start_col, start_row)
        if page != last_page:
            if last_page >= 0:
                pdf.showPage()
            last_page = page
        x = lm + col * (lw + cg)
        y_top = page_height - tm - row * (lh + rg)
        y = y_top - lh
        if item["type"] == "barcode":
            if not item["barcode"]:
                continue
            draw_barcode_label(pdf, x, y, lw, lh, library_name, library_place, item["barcode"])
        else:
            if not any(str(v).strip() for v in item["data"]):
                continue
            draw_side_label(pdf, x, y, lw, item["data"], 12)

    pdf.save()
    buffer.seek(0)
    return buffer

# -----------------------------
# Preview helpers
# -----------------------------

def show_pdf_preview(pdf_bytes):
    """Render the generated PDF as images so Chrome/Edge PDF iframe blocking
    cannot affect the preview."""
    try:
        pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_count = len(pdf)

        if page_count == 0:
            st.warning("The generated PDF has no pages.")
            pdf.close()
            return

        if page_count == 1:
            page_no = 0
        else:
            page_no = st.number_input(
                "Preview page",
                min_value=1,
                max_value=page_count,
                value=1,
                step=1,
                key="pdf_preview_page",
            ) - 1

        page = pdf.load_page(page_no)

        # 1.5x gives a clear preview while keeping the image reasonably sized.
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
        png_bytes = pix.tobytes("png")

        st.image(
            png_bytes,
            caption=f"Page {page_no + 1} of {page_count}",
            use_container_width=True,
        )
        pdf.close()

    except Exception as e:
        st.error(
            "PDF preview could not be rendered. "
            "The PDF can still be downloaded below."
        )
        st.caption(f"Preview error: {e}")


def show_generated_pdf(pdf_buffer, filename, button_label):
    pdf_bytes = pdf_buffer.getvalue()
    st.subheader("Generated PDF Preview")
    show_pdf_preview(pdf_bytes)
    st.download_button(
        button_label,
        pdf_bytes,
        file_name=filename,
        mime="application/pdf",
    )


def column_selectors(df, prefix, include_barcode=False):
    options = get_column_options(df)
    indices = [i for i, _ in options]
    labels = dict(options)
    result = {}
    if not indices:
        return result
    if include_barcode:
        result["barcode"] = st.selectbox(
            "Barcode Number Column", indices,
            index=3 if len(indices) >= 4 else 0,
            format_func=lambda i: labels[i], key=f"{prefix}_barcode_column"
        )
    st.subheader("Side Sticker Data Columns")
    a, b = st.columns(2)
    c, d = st.columns(2)
    with a:
        result["side1"] = st.selectbox("Line 1", indices, index=0,
            format_func=lambda i: labels[i], key=f"{prefix}_side_col_1")
    with b:
        result["side2"] = st.selectbox("Line 2", indices, index=min(1, len(indices)-1),
            format_func=lambda i: labels[i], key=f"{prefix}_side_col_2")
    with c:
        result["side3"] = st.selectbox("Line 3", indices, index=min(2, len(indices)-1),
            format_func=lambda i: labels[i], key=f"{prefix}_side_col_3")
    with d:
        result["side4"] = st.selectbox("Line 4", indices, index=min(3, len(indices)-1),
            format_func=lambda i: labels[i], key=f"{prefix}_side_col_4")
    return result

# -----------------------------
# UI
# -----------------------------

st.title("📚 Library Sticker Generator")
st.caption("DARUL IRFAN LIBRARY • PANDIKKAD")
st.info("Google Sheets links are supported. The sheet must be accessible to the app (for example, Share → Anyone with the link → Viewer).")

preset = st.selectbox("Sticker Preset", ["38x21 mm (5 x 13)", "Custom"])
if preset == "38x21 mm (5 x 13)":
    default_cols, default_rows, default_w, default_h = 5, 13, 38.0, 21.0
else:
    default_cols, default_rows, default_w, default_h = 5, 13, 38.0, 21.0

st.subheader("Paper & Label Settings")
c1, c2 = st.columns(2)
with c1:
    cols = int(st.number_input("Columns", 1, 20, default_cols))
    label_width_mm = st.number_input("Label Width (mm)", value=default_w)
    left_margin_mm = st.number_input("Left Margin (mm)", value=6.0)
    col_spacing_mm = st.number_input("Column Gap (mm)", value=2.0)
with c2:
    rows = int(st.number_input("Rows", 1, 30, default_rows))
    label_height_mm = st.number_input("Label Height (mm)", value=default_h)
    top_margin_mm = st.number_input("Top Margin (mm)", value=13.0)
    row_spacing_mm = st.number_input("Row Gap (mm)", value=0.0)

st.subheader("Start Position")
sc1, sc2 = st.columns(2)
with sc1:
    start_col = int(st.number_input("Start Column", 0, max(0, cols - 1), 0))
with sc2:
    start_row = int(st.number_input("Start Row", 0, max(0, rows - 1), 0))

st.subheader("Empty / Reserved Cells")
preserve_empty_rows = st.checkbox(
    "Keep empty data rows as reserved blank sticker cells", value=True,
    help="An empty row consumes its normal grid position but no sticker is printed there. This is useful when you want to preserve positions on a partially used label sheet."
)

st.divider()
sticker_type = st.radio(
    "Sticker Type",
    ["Barcode Sticker", "Side Sticker", "Combined (2 Barcode + 1 Side)"],
    horizontal=True,
)

# -----------------------------
# Barcode Sticker
# -----------------------------
if sticker_type == "Barcode Sticker":
    library_name = st.text_input("Library Name", "DARUL IRFAN LIBRARY")
    library_place = st.text_input("Library Place", "PANDIKKAD")
    source_type, uploaded, google_url, df, has_header, sheet_name, manual_name = render_source_selector("barcode")

    if df is not None:
        show_data_preview(df, "barcode")
        options = get_column_options(df)
        if options:
            labels = dict(options)
            indices = [i for i, _ in options]
            barcode_column = st.selectbox(
                "Barcode Number Column", indices,
                index=3 if len(indices) >= 4 else 0,
                format_func=lambda i: labels[i], key="barcode_column"
            )
            st.caption("Blank barcode cells are reserved as blank sticker positions.")
            if st.button("Generate Barcode Sticker PDF", type="primary"):
                try:
                    pdf = generate_barcode_pdf(
                        df, library_name, library_place, cols, rows,
                        label_width_mm, label_height_mm, left_margin_mm,
                        top_margin_mm, col_spacing_mm, row_spacing_mm,
                        start_col, start_row, barcode_column, preserve_empty_rows
                    )
                    name = output_filename(uploaded.name if uploaded else manual_name, "barcode") if uploaded else output_filename(manual_name, "barcode")
                    st.success("Barcode sticker PDF generated successfully.")
                    show_generated_pdf(pdf, name, "📥 Download Barcode Sticker PDF")
                except Exception as e:
                    st.error(f"Could not generate the barcode sticker PDF: {e}")

# -----------------------------
# Side Sticker
# -----------------------------
elif sticker_type == "Side Sticker":
    font_size = st.number_input("Side Sticker Font Size", min_value=6, max_value=24, value=12)
    source_type, uploaded, google_url, df, has_header, sheet_name, manual_name = render_source_selector("side")

    if df is not None:
        show_data_preview(df, "side")
        options = get_column_options(df)
        if options:
            selected = column_selectors(df, "side")
            if st.button("Generate Side Sticker PDF", type="primary"):
                try:
                    pdf = generate_side_pdf(
                        df,
                        [selected["side1"], selected["side2"], selected["side3"], selected["side4"]],
                        cols, rows, label_width_mm, label_height_mm, left_margin_mm,
                        top_margin_mm, col_spacing_mm, row_spacing_mm, start_col,
                        start_row, font_size, preserve_empty_rows
                    )
                    name = output_filename(uploaded.name if uploaded else manual_name, "side sticker") if uploaded else output_filename(manual_name, "side sticker")
                    st.success("Side sticker PDF generated successfully.")
                    show_generated_pdf(pdf, name, "📥 Download Side Sticker PDF")
                except Exception as e:
                    st.error(f"Could not generate the side sticker PDF: {e}")

# -----------------------------
# Combined
# -----------------------------
elif sticker_type == "Combined (2 Barcode + 1 Side)":
    barcode_copies = int(st.number_input("Barcode Copies", min_value=1, max_value=10, value=2))
    side_copies = int(st.number_input("Side Sticker Copies", min_value=1, max_value=10, value=1))
    library_name = st.text_input("Library Name", "DARUL IRFAN LIBRARY", key="combined_library_name")
    library_place = st.text_input("Library Place", "PANDIKKAD", key="combined_library_place")
    source_type, uploaded, google_url, df, has_header, sheet_name, manual_name = render_source_selector("combined")

    if df is not None:
        show_data_preview(df, "combined")
        options = get_column_options(df)
        if options:
            selected = column_selectors(df, "combined", include_barcode=True)
            if st.button("Generate Combined PDF", type="primary"):
                try:
                    pdf = make_combined_pdf(
                        df, selected["barcode"],
                        [selected["side1"], selected["side2"], selected["side3"], selected["side4"]],
                        library_name, library_place, barcode_copies, side_copies,
                        cols, rows, label_width_mm, label_height_mm, left_margin_mm,
                        top_margin_mm, col_spacing_mm, row_spacing_mm, start_col, start_row,
                        preserve_empty_rows
                    )
                    name = output_filename(uploaded.name if uploaded else manual_name, "combined") if uploaded else output_filename(manual_name, "combined")
                    st.success("Combined PDF generated successfully.")
                    show_generated_pdf(pdf, name, "📥 Download Combined PDF")
                except Exception as e:
                    st.error(f"Could not generate the combined PDF: {e}")
