from shiny import App, ui, reactive, render
import io
import os
import uuid
import base64
import tempfile
import pandas as pd
import datetime
from datetime import datetime as dd
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Image as RLImage
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.pdfbase.ttfonts import TTFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
font_path = os.path.join(BASE_DIR, "fonts", "NotoSansTC-Regular.ttf")

# image
HEADER_BASE64 = """
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
"""


def format_date(raw_date):
    """
    Convert any date string like '2024/5/20' to '2024/05/20'.
    If the input is empty or invalid, return an empty string.
    """
    if not raw_date:
        return ""
    
    # Replace different delimiters with '/'
    raw_date = raw_date.replace("-", "/").replace(".", "/")
    
    # Try parsing the date
    for fmt in ("%Y/%m/%d", "%Y/%d/%m", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            parsed = dd.strptime(raw_date, fmt)
            return parsed.strftime("%Y/%m/%d")  # zero-padded
        except ValueError:
            continue
    
    # fallback: return raw if parsing fails
    return raw_date


# ---------- Reactive Store ----------
cards = reactive.Value({})


# ---------- UI ----------
app_ui = ui.page_fluid(
    ui.tags.style(
        """
        body {
            margin: 0;
            padding: 0;
        }

        .card-container {
            display: grid;
            grid-template-columns: 105mm 105mm;
            grid-template-rows: 74.25mm 74.25mm 74.25mm 74.25mm;
            width: 210mm;
            height: 297mm;
        }

        .art-card {
            width: 105mm;
            height: 74.25mm;
            background: white;
            border: 0.01mm solid grey;
            box-sizing: border-box;
            padding: 6mm;
            display: flex;
            flex-direction: column;
            font-family: sans-serif;
        }

        .art-card img {
            width: 50mm;        /* fixed physical width */
            height: auto;
            object-fit: contain;
            display: block;
            margin: 0 auto 4mm auto;
        }

        .card-line {
            display: grid;
            grid-template-columns: 100mm auto;
            font-size: 14pt;
            margin-bottom: 3mm;
        }

        .card-label {
            text-align: right;
            padding-right: 3mm;
        }

        .card-value {
            text-align: left;
        }
        """
    ),
    ui.layout_sidebar(
        ui.sidebar(
            ui.input_file("csv_upload", "上傳 CSV", accept=[".csv"]),
            ui.hr(),
            ui.input_text("title", "作品名稱"),
            ui.input_text("author", "作者"),
            ui.input_text("height", "高度 (cm)"),
            ui.input_text("width", "寬度 (cm)"),
            ui.input_text("medium", "創作媒材"),
            ui.input_date("date", "創作日期", value=str(datetime.date.today())),
            ui.input_numeric("quantity", "數量", value=1, min=1, step=1),
            ui.input_action_button("add", "Add Card"),
            ui.hr(),
            ui.input_select("delete_select", "Select Card", choices=[]),
            ui.input_action_button("delete", "Delete Card"),
            ui.hr(),
            ui.download_button("download_pdf", "Download PDF"),
        ),
        ui.div({"class": "card-container"}, ui.output_ui("card_display")),
    ),
)


# ---------- Server ----------
def server(input, output, session):

    @reactive.effect
    @reactive.event(input.csv_upload)
    def load_csv():
        file = input.csv_upload()
        if not file:
            return

        # Read file, replace delimiters with comma
        with open(file[0]["datapath"], "r", encoding="utf-8") as f:
            content = f.read()
            content = content.replace("，", ",").replace(";", ",").replace("\t", ",").replace(", ", ",")  # replace ; or tab
            content = content.replace("（", " (").replace("）", ") ") # replace full sized paranthesis
            df = pd.read_csv(io.StringIO(content), delimiter=",", comment="#")

        data = cards.get().copy()

        for _, row in df.iterrows():
            cid = str(uuid.uuid4())[:8]

            data[cid] = {
                "title": str(row.get("title", "")),
                "author": str(row.get("author", "")),
                "height": str(row.get("height", "")),
                "width": str(row.get("width", "")),
                "medium": str(row.get("medium", "")),
                "date": format_date(str(row.get("date", ""))),
            }

        cards.set(data)

        ui.update_select(
            "delete_select", choices={cid: data[cid]["title"] for cid in data}
        )

    @reactive.effect
    @reactive.event(input.add)
    def add_card():
        data = cards.get().copy()
        qty = int(input.quantity() or 1)
        date_value = format_date(str(input.date()))

        for _ in range(qty):
            cid = str(uuid.uuid4())[:8]
            data[cid] = {
                "title": input.title(),
                "author": input.author(),
                "height": input.height(),
                "width": input.width(),
                "medium": input.medium(),
                "date": date_value,
            }

        cards.set(data)

        # Update dropdown with indexed labels
        choices = {}
        counter = {}

        for cid, content in data.items():
            title = content["title"]
            counter[title] = counter.get(title, 0) + 1
            label = f"{title} ({counter[title]})"
            choices[cid] = label

        ui.update_select("delete_select", choices=choices)

    @reactive.effect
    @reactive.event(input.delete)
    def delete_card():
        selected = input.delete_select()
        if selected:
            data = cards.get().copy()
            data.pop(selected, None)
            cards.set(data)

            # rebuild dropdown
            choices = {}
            counter = {}

            for cid, content in data.items():
                title = content["title"]
                counter[title] = counter.get(title, 0) + 1
                label = f"{title} ({counter[title]})"
                choices[cid] = label

            ui.update_select("delete_select", choices=choices)

    # ---------- Render Cards ----------
    @output
    @render.ui
    def card_display():

        elements = []

        for cid, content in cards.get().items():

            size_display = f"{content['height']} cm x {content['width']} cm"

            card = ui.div(
                {"class": "art-card"},
                ui.tags.img(src=f"data:image/png;base64,{HEADER_BASE64}"),
                ui.div({"class": "card-line"}, f"作品名稱: {content['title']}"),
                ui.div({"class": "card-line"}, f"作   者: {content['author']}"),
                ui.div({"class": "card-line"}, f"尺   寸: {size_display}"),
                ui.div({"class": "card-line"}, f"創作媒材: {content['medium']}"),
                ui.div(
                    {"class": "card-line"}, f"創作日期: {content.get('date', '')}"
                ),
            )

            elements.append(card)

        return elements

    # ---------- PDF Download ----------

    @output
    @render.download(filename="gallery-labels-print-ready.pdf")
    def download_pdf():
        # pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        pdfmetrics.registerFont(TTFont("NotoSansTC", font_path))

        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)

        page_width, page_height = A4

        card_width = 105 * mm
        card_height = 74.25 * mm

        cols, rows = 2, 4

        style = ParagraphStyle(
            name="Label", fontName="NotoSansTC", fontSize=14, leading=28
        )

        cards_data = list(cards.get().values())

        for index, content in enumerate(cards_data):

            position = index % 8
            col = position % cols
            row = position // cols

            x = col * card_width
            y = page_height - ((row + 1) * card_height)

            # Draw card border
            c.setLineWidth(0.2)
            c.rect(x, y, card_width, card_height)

            # Internal padding
            padding = 4 * mm
            cursor_y = y + card_height - padding

            # ---- Header Image ----
            if HEADER_BASE64:
                img_data = base64.b64decode(HEADER_BASE64)
                img_buffer = io.BytesIO(img_data)

                img_width = 50 * mm
                img = RLImage(img_buffer)
                scale = img_width / img.imageWidth
                img.drawWidth = img_width
                img.drawHeight = img.imageHeight * scale

                img_x = x + (card_width - img_width) / 2
                img_y = cursor_y - img.drawHeight

                img.drawOn(c, img_x, img_y)

                # Move cursor below image
                cursor_y = img_y - 4 * mm

            # ---- Text Content ----
            size_display = f"{content['height']} cm x {content['width']} cm"

            text_content = f"""
            作品名稱: {content['title']}<br/>
            作     者: {content['author']}<br/>
            尺     寸: {size_display}<br/>
            創作媒材: {content['medium']}<br/>
            創作日期: {content.get('date', '')}
            """

            paragraph = Paragraph(text_content, style)

            text_width = card_width - 2 * padding
            text_height = card_height

            paragraph.wrapOn(c, text_width, text_height)
            paragraph.drawOn(c, x + padding, cursor_y - paragraph.height)

            # New page every 8 cards
            if position == 7:
                c.showPage()

        c.save()
        buffer.seek(0)
        yield buffer.read()


# ---------- App ----------
app = App(app_ui, server)
