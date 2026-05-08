from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.styles.colors import Color
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
import re

# --- CONSTANTS ---
TEMPLATE_PATH = "POTemplate.xlsx"

# --- SMART MARKET EXPANSION DICTIONARY ---
MARKET_MAPPING = {
    "INDO": "INDONESIA",
    "KOR": "KOREA",
    "JAP": "JAPAN",
    "JPN": "JAPAN",
    "CHI": "CHINA",
    "CHN": "CHINA",
    "TAI": "TAIWAN",
    "TWN": "TAIWAN",
    "OTHER": "Other",
    "OTHERS": "Other"
}

# --- BORDERS ---
THIN_BORDER = Border(
    left=Side(style='thin', color='A5A5A5'), 
    right=Side(style='thin', color='A5A5A5'), 
    top=Side(style='thin', color='A5A5A5'), 
    bottom=Side(style='thin', color='A5A5A5')
)

THICK_BORDER = Border(
    left=Side(style='thin', color='000000'), 
    right=Side(style='thin', color='000000'), 
    top=Side(style='thin', color='000000'), 
    bottom=Side(style='thin', color='000000')
)

# --- FONTS ---
GREEN_DATA_FONT = Font(color="008608")
FONT_HEADER_DEFAULT = Font(bold=True, color="000000")
FONT_HEADER_RED = Font(bold=True, color="FF0000")

# --- ALIGNMENTS ---
HEADER_ALIGNMENT = Alignment(horizontal='center', vertical='center', wrap_text=True)
DATA_ALIGNMENT = Alignment(horizontal='center', vertical='center')

# --- FILLS ---
HEADER_FILL_NEUTRAL = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")

# --- NONE TEMPLATE HEADER CONFIGURATION ---
# (Start Col, End Col, Text)
NONE_TEMPLATE_HEADER_CONFIG = [
    (1, 1, "PO"),
    (2, 6, "customer ordered size"),
    (7, 11, "Sourcing Size"),
    (12, 16, "Tecn. Notation"),
    (17, 20, "size run ID"),
    (21, 24, "Chinese Customize Size - 2nd line size"),
    (25, 28, "Quantity"),
    (29, 33, "Age group"),
    (34, 38, "Gender"),
    (39, 43, "Size Spec / Pattern"),
    (44, 47, "Size page"),
    (48, 51, "Top/Bottom"),
    (52, 54, "Product division"),
    (55, 57, "Season"),
    (58, 60, "Style"),
    (61, 63, "Code item")
]

# --- UTILITY FUNCTIONS ---

def format_factory_name(name):
    if not name: return name
    name = str(name)
    name = re.sub(r'([a-zA-Z])\(', r'\1 (', name)
    name = re.sub(r'\)([a-zA-Z])', r') \1', name)
    name = re.sub(r'(Trax)(Apparel)', r'\1 \2', name, flags=re.IGNORECASE)
    return name.strip()

def format_ak_column(text_val):
    if not text_val: return text_val
    text_str = str(text_val)
    blue_color = Color(rgb="FF0070C0")
    red_color = Color(rgb="FFFF0000")
    blue_font = InlineFont(rFont="Calibri", sz=11, color=blue_color)
    red_font = InlineFont(rFont="Calibri", sz=11, color=red_color)

    if "between 50% to 74.9%" in text_str or "between 0% to 49.9%" in text_str:
        parts = re.split(r'(Goose down|Duck down)', text_str)
        rich_text_elements = []
        for part in parts:
            if not part: continue
            if part in ["Goose down", "Duck down"]:
                rich_text_elements.append(TextBlock(red_font, part))
            else:
                rich_text_elements.append(TextBlock(blue_font, part))
        return CellRichText(*rich_text_elements)

    if "Filling:" in text_str:
        return CellRichText(TextBlock(blue_font, text_str))

    return text_str

def clean_none_template_val(val):
    if not val: return ""
    v_str = str(val).strip()
    if v_str.upper() in ["NOT FOUND", "NAN", "NONE"]:
        return ""
    return v_str