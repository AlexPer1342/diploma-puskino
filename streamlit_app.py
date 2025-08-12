import io
import zipfile
import os
from datetime import datetime

import streamlit as st
import pandas as pd

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from pypdf import PdfReader, PdfWriter

# -------------------- BENDRA KONFIGŪRA --------------------
st.set_page_config(page_title="Padėkos raštų generatorius", layout="wide")

st.title("🎓 Padėkos raštų generatorius (Excel ➜ PDF)")
st.markdown(
    "Įkelkite **Excel** failą su stulpeliais "
    "`Vardas`, `Klasė`, `TIPAS`, `Komentaras`, `Metai` (jei tuščia, užpildysime automatiškai). "
    "PDF šabloną pasirinksite iš projekto aplanke esančių variantų."
)

# -------------------- LOKALŪS FAILAI / NUSTATYMAI --------------------
TEMPLATE_CANDIDATES = ["sablon2025.pdf", "sablon2025 s logo.pdf"]
CITY_PREFIX = "Vilnius"

FONT_REGULAR_FILE = "JosefinSans-Regular.ttf"
FONT_BOLD_FILE    = "JosefinSans-Bold.ttf"
FONT_LIGHT_FILE   = "JosefinSans-ExtraLight.ttf"

FONT_REGULAR_NAME = "JosefinSans"
FONT_BOLD_NAME    = "JosefinSansBold"
FONT_LIGHT_NAME   = "JosefinSansLight"

def register_font_safe(font_file: str, font_name: str) -> bool:
    if os.path.exists(font_file):
        try:
            pdfmetrics.registerFont(TTFont(font_name, font_file))
            st.success(f"Naudojamas šriftas: {font_name}")
            return True
        except Exception as e:
            st.warning(f"Nepavyko užregistruoti šrifto {font_file}: {e}")
    else:
        st.warning(f"Šriftas '{font_file}' nerastas.")
    return False

has_regular = register_font_safe(FONT_REGULAR_FILE, FONT_REGULAR_NAME)
has_bold    = register_font_safe(FONT_BOLD_FILE,    FONT_BOLD_NAME)
has_light   = register_font_safe(FONT_LIGHT_FILE,   FONT_LIGHT_NAME)

if not has_regular:
    FONT_REGULAR_NAME = "Helvetica"
if not has_bold:
    FONT_BOLD_NAME = FONT_REGULAR_NAME
if not has_light:
    FONT_LIGHT_NAME = FONT_REGULAR_NAME

# -------------------- ŠABLONO PASIRINKIMAS --------------------
available_templates = [p for p in TEMPLATE_CANDIDATES if os.path.exists(p)]
if not available_templates:
    st.error(
        "Nerasta nė vieno šablono faile: "
        + ", ".join(TEMPLATE_CANDIDATES)
        + ". Įkelk bent vieną iš jų į projekto aplanką."
    )
    st.stop()

selected_template = st.selectbox("Pasirinkite PDF šabloną", options=available_templates, index=0)

def probe_template_size(path: str):
    with open(path, "rb") as f:
        reader = PdfReader(f)
        page0 = reader.pages[0]
        mb = page0.mediabox
        return float(mb.width), float(mb.height)

try:
    TEMPLATE_PAGE_WIDTH, TEMPLATE_PAGE_HEIGHT = probe_template_size(selected_template)
    st.success(f"Šablonas: '{selected_template}' ({int(TEMPLATE_PAGE_WIDTH)} × {int(TEMPLATE_PAGE_HEIGHT)} pt)")
except Exception as e:
    st.error(f"Nepavyko perskaityti PDF šablono '{selected_template}': {e}")
    st.stop()

# -------------------- ĮKELIMAS: tik Excel + ŠABLONO ATSISIUNTIMAS --------------------
EXCEL_TEMPLATE_FILE = "padekos_testas.xlsx"  # <- jei tavo failas vadinasi kitaip, pakeisk čia

with st.expander("📄 Neturi Excel? Atsisiųsk paruoštą šabloną"):
    if os.path.exists(EXCEL_TEMPLATE_FILE):
        with open(EXCEL_TEMPLATE_FILE, "rb") as f:
            st.download_button(
                "📥 Atsisiųsti Excel šabloną (padekos_testas.xlsx)",
                data=f.read(),
                file_name=os.path.basename(EXCEL_TEMPLATE_FILE),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.warning(f"Šablono failas „{EXCEL_TEMPLATE_FILE}“ nerastas projekto aplanke.")

xls_file = st.file_uploader("Excel sąrašas", type=["xls", "xlsx"], key="xls")

xls_file = st.file_uploader("Excel sąrašas", type=["xls", "xlsx"], key="xls")

# -------------------- IŠDĖSTYMAS / NUSTATYMAI --------------------
st.sidebar.header("🛠️ Išdėstymo nustatymai (taškai)")
st.sidebar.caption("Koordinatės nuo kairio-apatinio kampo. Tekstas bus dedamas ant šablono.")

def xy_slider(label_x, label_y, default_x, default_y):
    x = st.sidebar.number_input(f"{label_x} (X)", value=float(default_x), step=1.0)
    y = st.sidebar.number_input(f"{label_y} (Y)", value=float(default_y), step=1.0)
    return x, y

tipas_x, tipas_y           = xy_slider("TIPAS X",         "TIPAS Y",         TEMPLATE_PAGE_WIDTH/2, 540)
vardas_x, vardas_y         = xy_slider("VARDAS X",        "VARDAS Y",        TEMPLATE_PAGE_WIDTH/2, 400)
klase_x, klase_y           = xy_slider("KLASĖ X",         "KLASĖ Y",         TEMPLATE_PAGE_WIDTH/2, 460)
komentaras_x, komentaras_y = xy_slider("KOMENTARAS X",    "KOMENTARAS Y",    TEMPLATE_PAGE_WIDTH/2, 360)
metai_x, metai_y           = xy_slider("MIESTAS/METAI X", "MIESTAS/METAI Y", TEMPLATE_PAGE_WIDTH/2, 55)

st.sidebar.subheader("Šriftų dydžiai")
fs_tipas      = st.sidebar.number_input("TIPAS (pt)",         value=46, min_value=8, max_value=96)
fs_vardas     = st.sidebar.number_input("VARDAS (pt)",        value=46, min_value=8, max_value=96)
fs_klase      = st.sidebar.number_input("KLASĖ (pt)",         value=20, min_value=8, max_value=96)
fs_komentaras = st.sidebar.number_input("KOMENTARAS (pt)",    value=20, min_value=8, max_value=96)
fs_metai      = st.sidebar.number_input("MIESTAS/METAI (pt)", value=14, min_value=8, max_value=96)

st.sidebar.subheader("Tekstų derinimas")
center_text = st.sidebar.checkbox("Centruoti tekstus pagal X", value=True)
wrap_comment = st.sidebar.checkbox("Laužyti komentarą iki pločio", value=True)
comment_width = st.sidebar.number_input("Komentaro maksimalus plotis (pt)",
                                        value=420, min_value=100, max_value=int(TEMPLATE_PAGE_WIDTH))

# vardo laužymas iki 2 eilučių
vardas_width = st.sidebar.number_input("Vardo maksimalus plotis (pt)",
                                       value=int(TEMPLATE_PAGE_WIDTH * 0.75),
                                       min_value=100, max_value=int(TEMPLATE_PAGE_WIDTH))

st.sidebar.subheader("Išvestis")
make_single_pdf = st.sidebar.checkbox("Sujungti visus į vieną PDF", value=False)
out_prefix = st.sidebar.text_input("Failų vardų priešdėlis", value="Padekos_rastas")

# -------------------- TEKSTO LAUŽYMAS --------------------
def _wrap_text_to_lines(c, text, font_used, size, max_width, max_lines=None):
    """
    Laužo tekstą į eilutes pagal max_width. Jei max_lines nurodytas (pvz., 2),
    sujungia likutį į paskutinę eilutę, kad bendras eilučių skaičius neviršytų max_lines.
    """
    if text is None:
        return [""]
    text = str(text).strip()
    if not max_width:
        return [text]

    words = text.split()
    if not words:
        return [""]

    lines = []
    cur = ""
    for w in words:
        test = (cur + " " + w).strip()
        if pdfmetrics.stringWidth(test, font_used, size) <= max_width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)

    if max_lines and len(lines) > max_lines:
        head = lines[:max_lines-1]
        tail = " ".join(lines[max_lines-1:])
        lines = head + [tail]

    return lines

# -------------------- PIEŠIMAS / MERGE --------------------
def make_overlay_pdf(row, page_width, page_height):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_width, page_height))

    def draw_text(x, y, text, size, font_used, align_center=True, max_width=None, max_lines=None):
        """
        Piešia tekstą. Jei nurodytas max_width, laužo tekstą į eilutes.
        Grąžina panaudotų eilučių skaičių (naudinga vardo atvejui).
        """
        c.setFont(font_used, size)
        if max_width:
            lines = _wrap_text_to_lines(c, text, font_used, size, max_width, max_lines=max_lines)
            line_height = size * 1.2
            start_y = y
            for i, line in enumerate(lines):
                if align_center:
                    c.drawCentredString(x, start_y - i * line_height, line)
                else:
                    c.drawString(x, start_y - i * line_height, line)
            return len(lines)
        else:
            s = "" if text is None else str(text)
            if align_center:
                c.drawCentredString(x, y, s)
            else:
                c.drawString(x, y, s)
            return 1

    c.setTitle("Padėkos raštas")

    # TIPAS – Bold
    draw_text(tipas_x, tipas_y, row.get("TIPAS", ""), fs_tipas, FONT_BOLD_NAME, align_center=center_text)

    # VARDAS – Regular, laužyti iki 2 eilučių pagal 'vardas_width'
    name_lines_used = draw_text(
        vardas_x, vardas_y,
        row.get("Vardas", ""),
        fs_vardas, FONT_REGULAR_NAME,
        align_center=center_text,
        max_width=vardas_width,
        max_lines=2
    )

    # KLASĖ – Regular
    draw_text(klase_x, klase_y, row.get("Klasė", ""), fs_klase, FONT_REGULAR_NAME, align_center=center_text)

    # Jei vardas užėmė dvi eilutes – nuleidžiame komentarą 40 pt žemiau
    komentaras_y_adj = komentaras_y - 40 if name_lines_used > 1 else komentaras_y

    # KOMENTARAS – ExtraLight (laužomas pagal comment_width)
    draw_text(
        komentaras_x, komentaras_y_adj,
        row.get("Komentaras", ""),
        fs_komentaras, FONT_LIGHT_NAME,
        align_center=center_text,
        max_width=comment_width,
        max_lines=None
    )

    # MIESTAS/METAI – ExtraLight
    draw_text(metai_x, metai_y, row.get("Metai", ""), fs_metai, FONT_LIGHT_NAME, align_center=center_text)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf

def merge_overlay_with_template(template_bytes, overlay_bytes):
    tpl_reader = PdfReader(template_bytes)
    ovl_reader = PdfReader(overlay_bytes)

    tpl_page = tpl_reader.pages[0]
    ovl_page = ovl_reader.pages[0]

    tpl_page.merge_page(ovl_page)

    writer = PdfWriter()
    writer.add_page(tpl_page)

    out_buf = io.BytesIO()
    writer.write(out_buf)
    out_buf.seek(0)
    return out_buf

# -------------------- EXCEL NUSKAITYMAS + „Metai“ AUTO --------------------
df = None
if xls_file is not None:
    try:
        df = pd.read_excel(xls_file)
        required = ["Vardas", "Klasė", "TIPAS", "Komentaras"]
        missing_req = [c for c in required if c not in df.columns]
        if missing_req:
            st.error(f"Trūksta privalomų stulpelių: {', '.join(missing_req)}")
            df = None
        else:
            if "Metai" not in df.columns:
                df["Metai"] = ""
            current_year = datetime.now().year
            df["Metai"] = df["Metai"].apply(lambda v: f"{CITY_PREFIX}, {current_year}" if pd.isna(v) or str(v).strip() == "" else v)
            st.success("Excel nuskaitytas sėkmingai. Tuščios „Metai“ reikšmės užpildytos automatiškai.")
            st.dataframe(df.head(20))
    except Exception as e:
        st.error(f"Nepavyko nuskaityti Excel failo: {e}")

# -------------------- PREVIEW FUNKCIJA --------------------
if df is not None and len(df) > 0:
    st.subheader("👁️ Peržiūros režimas")
    row_index = st.number_input("Pasirinkite eilutės indeksą peržiūrai", min_value=0, max_value=len(df)-1, value=0, step=1)
    if st.button("🔍 Generuoti peržiūrą pasirinktam įrašui"):
        with open(selected_template, "rb") as base_tpl:
            template_bytes_data = base_tpl.read()
        preview_buf = merge_overlay_with_template(
            io.BytesIO(template_bytes_data),
            make_overlay_pdf(df.iloc[row_index], TEMPLATE_PAGE_WIDTH, TEMPLATE_PAGE_HEIGHT)
        )
        st.download_button(
            "⬇️ Atsisiųsti peržiūros PDF",
            data=preview_buf,
            file_name=f"preview_{df.iloc[row_index]['Vardas']}.pdf",
            mime="application/pdf",
        )

# -------------------- GENERAVIMAS --------------------
st.divider()
generate = st.button("🚀 Generuoti PDF(-us)", type="primary", disabled=df is None)

if generate:
    try:
        with open(selected_template, "rb") as base_tpl:
            template_bytes_data = base_tpl.read()
        pdf_buffers = []
        for idx, row in df.iterrows():
            overlay_buf = make_overlay_pdf(row, TEMPLATE_PAGE_WIDTH, TEMPLATE_PAGE_HEIGHT)
            merged_buf = merge_overlay_with_template(io.BytesIO(template_bytes_data), overlay_buf)
            safe_name = str(row.get("Vardas", f"asmuo_{idx}")).strip().replace("/", "_").replace("\\", "_")
            pdf_buffers.append((f"{out_prefix}_{safe_name}.pdf", merged_buf))
        if make_single_pdf:
            writer = PdfWriter()
            for _, buf in pdf_buffers:
                reader = PdfReader(buf)
                for p in reader.pages:
                    writer.add_page(p)
            single_buf = io.BytesIO()
            writer.write(single_buf)
            single_buf.seek(0)
            st.success(f"Sukurta {len(pdf_buffers)} padėkos raštų į vieną PDF.")
            st.download_button(
                "⬇️ Atsisiųsti vieną PDF",
                data=single_buf,
                file_name=f"{out_prefix}_visi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
            )
        else:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, buf in pdf_buffers:
                    zf.writestr(fname, buf.getvalue())
            zip_buf.seek(0)
            st.success(f"Sukurta {len(pdf_buffers)} atskirų PDF.")
            st.download_button(
                "⬇️ Atsisiųsti ZIP archyvą",
                data=zip_buf,
                file_name=f"{out_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
            )
    except Exception as e:
        st.error(f"Generuojant įvyko klaida: {e}")
