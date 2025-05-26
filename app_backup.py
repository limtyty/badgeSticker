import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import os

# Config
COLUMNS = 4
ROWS = 9
CELL_SPACING = 1  # mm
LEFT_MARGIN = 10
TOP_MARGIN = 10
LINE_HEIGHT = 5
PADDING = 2
MAX_FONT_SIZE = 10
MIN_FONT_SIZE = 6

# Font setup
FONT_MULI = "Muli"
FONT_DIN = "DIN"
MULI_FONT_PATH = "fonts/Muli-Black.ttf"
DIN_FONT_PATH = "fonts/DINRegular.ttf"

# Helper: Shrink text to fit max width
def shrink_text_to_fit(pdf, text, font, max_width, max_size=MAX_FONT_SIZE, min_size=MIN_FONT_SIZE):
    size = max_size
    pdf.set_font(font, "", size)
    while pdf.get_string_width(text) > max_width and size > min_size:
        size -= 0.5
        pdf.set_font(font, "", size)
    return size

# PDF Class
class GridPDF(FPDF):
    def __init__(self):
        super().__init__(orientation='L', format='A4')
        self.set_auto_page_break(auto=False)
        self.set_margins(LEFT_MARGIN, TOP_MARGIN, LEFT_MARGIN)
        self.add_page()
        self.add_font(FONT_MULI, "", MULI_FONT_PATH, uni=True)
        self.add_font(FONT_DIN, "", DIN_FONT_PATH, uni=True)

    def draw_grid(self, data):
        page_width = self.w - 2 * LEFT_MARGIN
        page_height = self.h - 2 * TOP_MARGIN
        cell_width = (page_width - (COLUMNS - 1) * CELL_SPACING) / COLUMNS
        cell_height = (page_height - (ROWS - 1) * CELL_SPACING) / ROWS
        max_width = cell_width - 2 * PADDING

        for index, row in enumerate(data):
            col = index % COLUMNS
            row_pos = (index // COLUMNS) % ROWS
            if index % (COLUMNS * ROWS) == 0 and index != 0:
                self.add_page()

            x = LEFT_MARGIN + col * (cell_width + CELL_SPACING)
            y = TOP_MARGIN + row_pos * (cell_height + CELL_SPACING)

            # Draw box
            self.rect(x, y, cell_width, cell_height)

            # Full Name (Muli)
            name = str(row['Full Name'])
            name_size = shrink_text_to_fit(self, name, FONT_MULI, max_width)
            self.set_font(FONT_MULI, "", name_size)
            self.set_xy(x + PADDING, y + PADDING)
            self.cell(max_width, LINE_HEIGHT, name, align='C')

            # Position (DIN)
            Position = str(row['Position'])
            Position_size = shrink_text_to_fit(self, Position, FONT_DIN, max_width)
            self.set_font(FONT_DIN, "", Position_size)
            self.set_xy(x + PADDING, y + PADDING + LINE_HEIGHT)
            self.cell(max_width, LINE_HEIGHT, Position, align='C')

            # Company (DIN)
            company = str(row['Company'])
            company_size = shrink_text_to_fit(self, company, FONT_DIN, max_width)
            self.set_font(FONT_DIN, "", company_size)
            self.set_xy(x + PADDING, y + PADDING + 2 * LINE_HEIGHT)
            self.cell(max_width, LINE_HEIGHT, company, align='C')

# Streamlit app
st.set_page_config(page_title="PDF Grid Generator", layout="wide")
st.title("📄 PDF Grid Generator (Shrink-to-Fit Text)")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Font existence check
font_errors = []
if not os.path.exists(MULI_FONT_PATH):
    font_errors.append("❌ Muli-Black.ttf not found in fonts/")
if not os.path.exists(DIN_FONT_PATH):
    font_errors.append("❌ DINRegular.ttf not found in fonts/")

if font_errors:
    for e in font_errors:
        st.error(e)
    st.info("📁 Please place required font files in a `fonts/` folder.")
else:
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)

            required_cols = {"Full Name", "Position", "Company"}
            if not required_cols.issubset(df.columns):
                st.error("❌ Excel must contain: Full Name, Position, Company")
            else:
                st.success("✅ File loaded!")
                st.dataframe(df)

                if st.button("Generate PDF"):
                    pdf = GridPDF()
                    records = df.to_dict(orient="records")
                    pdf.draw_grid(records)

                    buffer = BytesIO(pdf.output(dest="S").encode("latin-1"))

                    st.download_button(
                        label="📥 Download PDF",
                        data=buffer,
                        file_name="Chamber_Profiles_ShrinkWrap.pdf",
                        mime="application/pdf"
                    )

        except Exception as e:
            st.error(f"⚠️ Error reading file: {e}")
