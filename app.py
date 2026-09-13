
# Library Sticker Generator - Streamlit
import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code128

st.set_page_config(page_title="Library Sticker Generator", layout="wide")


def read_uploaded_table(uploaded):
    """Read CSV/XLSX/ODS while preserving barcode values as text."""
    name = uploaded.name.lower()

    if name.endswith(".csv"):
        uploaded.seek(0)
        return pd.read_csv(uploaded, dtype=str, keep_default_na=False)
    elif name.endswith(".xlsx"):
        uploaded.seek(0)
        return pd.read_excel(uploaded, dtype=str, keep_default_na=False)
    elif name.endswith(".ods"):
        uploaded.seek(0)
        return pd.read_excel(
            uploaded,
            engine="odf",
            dtype=str,
            keep_default_na=False
        )

    raise ValueError("Unsupported file type.")

def generate_barcode_pdf(df, library_name, library_place,
                         cols, rows, label_width_mm, label_height_mm,
                         left_margin_mm, top_margin_mm,
                         col_spacing_mm, row_spacing_mm,
                         start_col, start_row,
                         barcode_column):

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_width, page_height = A4

    label_width = label_width_mm * mm
    label_height = label_height_mm * mm
    left_margin = left_margin_mm * mm
    top_margin = top_margin_mm * mm
    col_spacing = col_spacing_mm * mm
    row_spacing = row_spacing_mm * mm

    if barcode_column < 0 or barcode_column >= len(df.columns):
        raise ValueError("The selected barcode column does not exist.")

    numbers = df.iloc[:, barcode_column].astype(str).tolist()

    for i, number in enumerate(numbers):

        position = i + (start_row * cols) + start_col

        col = position % cols
        row = (position // cols) % rows

        x = left_margin + col * (label_width + col_spacing)
        y_top = page_height - top_margin - row * (label_height + row_spacing)
        y = y_top - label_height

        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(x + label_width/2, y + 17*mm, library_name)
        c.drawCentredString(x + label_width/2, y + 14*mm, library_place)

        barcode = code128.Code128(
            str(number),
            barHeight=8*mm,
            barWidth=0.53*mm,
            humanReadable=False
        )

        barcode_x = x + (label_width - barcode.width) / 2
        barcode_y = y + 4.5*mm

        barcode.drawOn(c, barcode_x, barcode_y)

        c.drawCentredString(
            x + label_width/2,
            y + 2*mm,
            str(number)
        )

        if (i + 1) % (cols * rows) == 0:
            c.showPage()

    c.save()
    buffer.seek(0)
    return buffer


def generate_side_pdf(df,
                      cols, rows, label_width_mm, label_height_mm,
                      left_margin_mm, top_margin_mm,
                      col_spacing_mm, row_spacing_mm,
                      start_col, start_row,
                      font_size):

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)

    page_width, page_height = A4

    label_width = label_width_mm * mm
    label_height = label_height_mm * mm
    left_margin = left_margin_mm * mm
    top_margin = top_margin_mm * mm
    col_spacing = col_spacing_mm * mm
    row_spacing = row_spacing_mm * mm

    line_positions = [4*mm, 8.5*mm, 13*mm, 17.5*mm]

    current_col = start_col
    current_row = start_row

    pdf.setFont("Times-Bold", font_size)

    labels = df.values.tolist()

    for label in labels:

        x = left_margin + current_col * (label_width + col_spacing)

        y_top = page_height - top_margin - current_row * (
            label_height + row_spacing
        )

        center_x = x + label_width / 2

        for i in range(4):

            text = ""

            if i < len(label):
                value = label[i]

                if pd.notna(value):
                    text = str(value)

            y = y_top - line_positions[i]

            pdf.drawCentredString(center_x, y, text)

        current_col += 1

        if current_col >= cols:
            current_col = 0
            current_row += 1

        if current_row >= rows:
            pdf.showPage()
            pdf.setFont("Times-Bold", font_size)

            current_row = start_row
            current_col = start_col

    pdf.save()
    buffer.seek(0)
    return buffer


st.title("📚 Library Sticker Generator")

preset = st.selectbox(
    "Sticker Preset",
    [
        "38x21 mm (5 x 13)",
        "Custom"
    ]
)

if preset == "38x21 mm (5 x 13)":
    default_cols = 5
    default_rows = 13
    default_w = 38.0
    default_h = 21.0
else:
    default_cols = 5
    default_rows = 13
    default_w = 38.0
    default_h = 21.0

st.subheader("Paper & Label Settings")

c1, c2 = st.columns(2)

with c1:
    cols = st.number_input("Columns", 1, 20, default_cols)
    label_width_mm = st.number_input("Label Width (mm)", value=default_w)
    left_margin_mm = st.number_input("Left Margin (mm)", value=6.0)
    col_spacing_mm = st.number_input("Column Gap (mm)", value=2.0)

with c2:
    rows = st.number_input("Rows", 1, 30, default_rows)
    label_height_mm = st.number_input("Label Height (mm)", value=default_h)
    top_margin_mm = st.number_input("Top Margin (mm)", value=13.0)
    row_spacing_mm = st.number_input("Row Gap (mm)", value=0.0)

st.subheader("Start Position")

sc1, sc2 = st.columns(2)

with sc1:
    start_col = st.number_input("Start Column", 0, 20, 0)

with sc2:
    start_row = st.number_input("Start Row", 0, 50, 0)

st.divider()

sticker_type = st.radio(
    "Sticker Type",
    [
        "Barcode Sticker",
        "Side Sticker",
        "Combined (2 Barcode + 1 Side)"
    ]
)

if sticker_type == "Barcode Sticker":

    library_name = st.text_input(
        "Library Name",
        "DARUL IRFAN LIBRARY"
    )

    library_place = st.text_input(
        "Library Place",
        "PANDIKKAD"
    )

    uploaded = st.file_uploader(
        "Upload CSV/XLSX/ODS",
        type=["csv", "xlsx", "ods"]
    )

    if uploaded:
        try:
            df = read_uploaded_table(uploaded)

            if df.shape[1] == 0:
                st.error("The uploaded file has no columns.")
            else:
                column_names = [
                    f"{i + 1}: {str(col)}"
                    for i, col in enumerate(df.columns)
                ]

                default_index = 3 if df.shape[1] >= 4 else 0

                barcode_choice = st.selectbox(
                    "Barcode Number Column",
                    options=range(len(df.columns)),
                    index=default_index,
                    format_func=lambda i: column_names[i]
                )

                st.caption(
                    f"Loaded {len(df)} rows and {df.shape[1]} columns. "
                    "The 4th column is selected automatically when available."
                )

                if df.shape[1] < 4:
                    st.warning(
                        "Your file has fewer than 4 columns. "
                        "Select the column that actually contains the barcode number."
                    )

                if st.button("Generate Barcode PDF"):
                    try:
                        pdf = generate_barcode_pdf(
                            df,
                            library_name,
                            library_place,
                            cols,
                            rows,
                            label_width_mm,
                            label_height_mm,
                            left_margin_mm,
                            top_margin_mm,
                            col_spacing_mm,
                            row_spacing_mm,
                            start_col,
                            start_row,
                            barcode_choice
                        )

                        st.download_button(
                            "Download PDF",
                            pdf,
                            file_name="barcode_stickers.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"Could not generate the barcode PDF: {e}")

        except Exception as e:
            st.error(
                "Could not read the uploaded file. "
                f"Please check the CSV/XLSX/ODS format. Details: {e}"
            )

elif sticker_type == "Side Sticker":

    font_size = st.number_input(
        "Font Size",
        6,
        30,
        12
    )

    uploaded = st.file_uploader(
        "Upload CSV/XLSX/ODS",
        type=["csv", "xlsx", "ods"]
    )

    if uploaded and st.button("Generate Side Sticker PDF"):

        name = uploaded.name.lower()

        if name.endswith(".csv"):
            df = pd.read_csv(uploaded, header=None)
        elif name.endswith(".xlsx"):
            df = pd.read_excel(uploaded, header=None)
        else:
            df = pd.read_excel(
                uploaded,
                engine="odf",
                header=None
            )

        pdf = generate_side_pdf(
            df,
            cols,
            rows,
            label_width_mm,
            label_height_mm,
            left_margin_mm,
            top_margin_mm,
            col_spacing_mm,
            row_spacing_mm,
            start_col,
            start_row,
            font_size
        )

        st.download_button(
            "Download PDF",
            pdf,
            file_name="side_stickers.pdf",
            mime="application/pdf"
        )
elif sticker_type == "Combined (2 Barcode + 1 Side)":

    uploaded = st.file_uploader(
        "Upload Excel File",
        type=["xlsx", "csv","ods"]
    )

    barcode_copies = st.number_input(
        "Barcode Copies",
        min_value=1,
        max_value=10,
        value=2
    )

    side_copies = st.number_input(
        "Side Sticker Copies",
        min_value=1,
        max_value=10,
        value=1
    )

    library_name = st.text_input(
        "Library Name",
        "DARUL IRFAN LIBRARY"
    )

    library_place = st.text_input(
        "Library Place",
        "PANDIKKAD"
    )

    if uploaded:
        try:
            df = read_uploaded_table(uploaded)

            if df.shape[1] == 0:
                st.error("The uploaded file has no columns.")
                st.stop()

            column_names = [
                f"{i + 1}: {str(col)}"
                for i, col in enumerate(df.columns)
            ]

            default_index = 3 if df.shape[1] >= 4 else 0

            barcode_choice = st.selectbox(
                "Barcode Number Column",
                options=range(len(df.columns)),
                index=default_index,
                format_func=lambda i: column_names[i],
                key="combined_barcode_column"
            )

            if df.shape[1] < 4:
                st.warning(
                    "Your file has fewer than 4 columns. "
                    "Select the column containing the barcode number."
                )

        except Exception as e:
            st.error(
                "Could not read the uploaded file. "
                f"Please check the CSV/XLSX/ODS format. Details: {e}"
            )
            st.stop()

    if uploaded and st.button("Generate Combined PDF"):

        combined_rows = []

        for _, row in df.iterrows():

            barcode = str(row.iloc[barcode_choice])

            side_data = [
                str(row.iloc[0]) if pd.notna(row.iloc[0]) else "",
                str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
                str(row.iloc[2]) if pd.notna(row.iloc[2]) else "",
                  str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
            ]

            for _ in range(barcode_copies):
                combined_rows.append({
                    "type": "barcode",
                    "barcode": barcode
                })

            for _ in range(side_copies):
                combined_rows.append({
                    "type": "side",
                    "data": side_data
                })

        buffer = BytesIO()
        pdf = canvas.Canvas(buffer, pagesize=A4)

        page_width, page_height = A4

        label_width = label_width_mm * mm
        label_height = label_height_mm * mm
        left_margin = left_margin_mm * mm
        top_margin = top_margin_mm * mm
        col_spacing = col_spacing_mm * mm
        row_spacing = row_spacing_mm * mm

        for i, item in enumerate(combined_rows):

            position = i + (start_row * cols) + start_col

            col = position % cols
            row = (position // cols) % rows

            x = left_margin + col * (label_width + col_spacing)

            y_top = page_height - top_margin - row * (
                label_height + row_spacing
            )

            y = y_top - label_height

            if item["type"] == "barcode":

                pdf.setFont("Helvetica-Bold", 7)

                pdf.drawCentredString(
                    x + label_width/2,
                    y + 17*mm,
                    library_name
                )

                pdf.drawCentredString(
                    x + label_width/2,
                    y + 14*mm,
                    library_place
                )

                barcode = code128.Code128(
                    item["barcode"],
                    barHeight=8*mm,
                    barWidth=0.53*mm,
                    humanReadable=False
                )

                barcode_x = x + (
                    label_width - barcode.width
                ) / 2

                barcode.drawOn(
                    pdf,
                    barcode_x,
                    y + 4.5*mm
                )

                pdf.drawCentredString(
                    x + label_width/2,
                    y + 2*mm,
                    item["barcode"]
                )

            else:

                pdf.setFont("Times-Bold", 12)

                lines = item["data"]

                positions = [
                    17.5*mm,
                    13*mm,
                    8.5*mm,
                    4*mm
                ]

                for n in range(4):

                    text = ""

                    if n < len(lines):
                        text = str(lines[n])

                    pdf.drawCentredString(
                        x + label_width/2,
                        y + positions[n],
                        text
                    )

            if (i + 1) % (cols * rows) == 0:
                pdf.showPage()

        pdf.save()

        buffer.seek(0)

        st.download_button(
            "Download Combined PDF",
            buffer,
            file_name="combined_stickers.pdf",
            mime="application/pdf"
        )