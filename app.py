import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import os

# --- Page and Grid Configuration ---
PAGE_WIDTH = 297  # A4 landscape width in mm
PAGE_HEIGHT = 210  # A4 height in mm

COLUMNS = 4
ROWS = 8
CELL_WIDTH = 70  # mm
CELL_HEIGHT = 23  # mm
CELL_SPACING = 1  # mm between boxes
LINE_HEIGHT = 5.5
PADDING = 2
MAX_FONT_SIZE = 11
MIN_FONT_SIZE = 7

# Calculate total grid size
grid_width = COLUMNS * CELL_WIDTH + (COLUMNS - 1) * CELL_SPACING
grid_height = ROWS * CELL_HEIGHT + (ROWS - 1) * CELL_SPACING

# Center margins
LEFT_MARGIN = RIGHT_MARGIN = (PAGE_WIDTH - grid_width) / 2
TOP_MARGIN = BOTTOM_MARGIN = (PAGE_HEIGHT - grid_height) / 2

# --- Font Paths ---
FONT_MULI = "Muli"
FONT_DIN = "DIN"
MULI_FONT_PATH = "fonts/Muli-Black.ttf"
DIN_FONT_PATH = "fonts/DINRegular.ttf"

# --- Utility: Shrink font to fit ---
def shrink_text_to_fit(pdf, text, font, max_width, max_size=MAX_FONT_SIZE, min_size=MIN_FONT_SIZE):
    size = max_size
    pdf.set_font(font, "", size)
    while pdf.get_string_width(text) > max_width and size > min_size:
        size -= 0.5
        pdf.set_font(font, "", size)
    return size

# --- Utility: Wrap text to fit within max width ---
def wrap_text(pdf, text, font, font_size, max_width):
    pdf.set_font(font, "", font_size)
    words = text.split()
    lines = []
    current_line = ""
    
    for word in words:
        test_line = current_line + word + " "
        if pdf.get_string_width(test_line.strip()) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line.strip())
            current_line = word + " "
    if current_line:
        lines.append(current_line.strip())
    
    return lines

# --- PDF Generator Class ---
class GridPDF(FPDF):
    def __init__(self):
        super().__init__(orientation='L', format='A4')
        self.set_auto_page_break(auto=False)
        self.set_margins(LEFT_MARGIN, TOP_MARGIN, RIGHT_MARGIN)
        self.add_page()
        self.add_font(FONT_MULI, "", MULI_FONT_PATH, uni=True)
        self.add_font(FONT_DIN, "", DIN_FONT_PATH, uni=True)

    def draw_grid(self, data):
        max_width = CELL_WIDTH - 2 * PADDING

        for index, row in enumerate(data):
            cell_index = index % (COLUMNS * ROWS)
            col = cell_index % COLUMNS
            row_pos = cell_index // COLUMNS

            if cell_index == 0 and index != 0:
                self.add_page()

            x = LEFT_MARGIN + col * (CELL_WIDTH + CELL_SPACING)
            y = TOP_MARGIN + row_pos * (CELL_HEIGHT + CELL_SPACING)

            # Draw border
            self.rect(x, y, CELL_WIDTH, CELL_HEIGHT)

            # Prepare text entries with wrapping, only include non-empty and non-NaN fields
            entries = []
            for field, font, key in [
                ('Full Name', FONT_MULI, 'Full Name'),
                ('Position', FONT_DIN, 'Position'),
                ('Company', FONT_DIN, 'Company')
            ]:
                value = str(row.get(key, ''))
                # Skip if value is NaN, None, empty, or 'nan' (case-insensitive)
                if pd.notna(value) and value.strip() and value.lower() != 'nan':
                    entries.append((field, font, value))

            lines = []
            for field, font, text in entries:
                font_size = shrink_text_to_fit(self, text, font, max_width)
                wrapped_lines = wrap_text(self, text, font, font_size, max_width)
                lines.extend([(line, font, font_size) for line in wrapped_lines])

            # Center vertically
            total_height = len(lines) * LINE_HEIGHT
            start_y = y + (CELL_HEIGHT - total_height) / 2 if total_height < CELL_HEIGHT else y

            # Draw wrapped text
            for i, (text, font_name, font_size) in enumerate(lines):
                if start_y + i * LINE_HEIGHT + LINE_HEIGHT <= y + CELL_HEIGHT:  # Ensure text fits in cell
                    self.set_font(font_name, "", font_size)
                    self.set_xy(x + PADDING, start_y + i * LINE_HEIGHT)
                    self.cell(max_width, LINE_HEIGHT, text, align='C')

# --- Streamlit App ---
st.set_page_config(page_title="PDF Badge Generator", layout="wide")
st.title("PDF Generator — 4x8 Badge Canada Cambodia Mission")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Check font files
font_errors = []
if not os.path.exists(MULI_FONT_PATH):
    font_errors.append("❌ Muli-Black.ttf not found in fonts/")
if not os.path.exists(DIN_FONT_PATH):
    font_errors.append("❌ DINRegular.ttf not found in fonts/")

if font_errors:
    for e in font_errors:
        st.error(e)
    st.info("📁 Please make sure both font files are in a `fonts/` folder.")
else:
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)

            required = {"Full Name", "Position", "Company"}
            if not required.issubset(df.columns):
                st.error("❌ Excel must contain: Full Name, Position, Company")
            else:
                st.success("✅ Excel loaded successfully!")
                st.dataframe(df)

                if st.button("Generate PDF"):
                    pdf = GridPDF()
                    records = df.to_dict(orient="records")
                    pdf.draw_grid(records)

                    buffer = BytesIO(pdf.output(dest="S").encode("latin-1"))
                    st.download_button(
                        label="📥 Download PDF",
                        data=buffer,
                        file_name="Chamber_Profiles_4x8_Centered.pdf",
                        mime="application/pdf"
                    )

        except Exception as e:
            st.error(f"⚠️ Error reading file: {e}")