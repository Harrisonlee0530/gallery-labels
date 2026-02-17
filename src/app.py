from shiny import App, ui, reactive, render
from shiny.types import FileInfo
import uuid
import os
from pathlib import Path
from reportlab.platypus import SimpleDocTemplate, Image, Spacer
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Frame, PageTemplate
from reportlab.platypus import BaseDocTemplate
from reportlab.platypus import Flowable
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import PageBreak
from reportlab.platypus import KeepTogether
from reportlab.platypus import Image as RLImage
import tempfile

# ---------- Reactive Card Store ----------
cards = reactive.Value({})

# ---------- UI ----------
app_ui = ui.page_fluid(

    ui.tags.style("""
    .card-container {
        display: flex;
        flex-wrap: wrap;
        gap: 10mm;
    }

    .art-card {
        width: 105mm;
        height: 74.25mm;
        background-image: url('/template.png');
        background-size: cover;
        background-repeat: no-repeat;
        position: relative;
        padding: 12mm 8mm;
        box-sizing: border-box;
        font-family: sans-serif;
    }

    .card-text {
        font-size: 10pt;
        line-height: 1.4;
        margin-top: 6mm;
    }
    """),

    ui.layout_sidebar(
        ui.sidebar(
            ui.input_text("title", "作品名稱"),
            ui.input_text("author", "作者"),
            ui.input_text("size", "尺寸"),
            ui.input_text("medium", "創作媒材"),
            ui.input_action_button("add", "Add Card"),
            ui.hr(),
            ui.input_select("delete_select", "Select Card", choices=[]),
            ui.input_action_button("delete", "Delete Card"),
            ui.hr(),
            ui.download_button("download_png", "Download PNG"),
            ui.download_button("download_pdf", "Download PDF"),
        ),
        ui.div(
            {"class": "card-container"},
            ui.output_ui("card_display")
        )
    )
)

# ---------- Server ----------
def server(input, output, session):

    @reactive.effect
    @reactive.event(input.add)
    def add_card():
        new_id = str(uuid.uuid4())[:8]
        cards_dict = cards.get().copy()
        cards_dict[new_id] = {
            "title": input.title(),
            "author": input.author(),
            "size": input.size(),
            "medium": input.medium(),
        }
        cards.set(cards_dict)

        ui.update_select(
            "delete_select",
            choices=list(cards_dict.keys())
        )

    @reactive.effect
    @reactive.event(input.delete)
    def delete_card():
        selected = input.delete_select()
        if selected:
            cards_dict = cards.get().copy()
            cards_dict.pop(selected, None)
            cards.set(cards_dict)

            ui.update_select(
                "delete_select",
                choices=list(cards_dict.keys())
            )

    @output
    @render.ui
    def card_display():
        card_list = []
        for cid, content in cards.get().items():
            card = ui.div(
                {"class": "art-card"},
                ui.div(
                    {"class": "card-text"},
                    f"作品名稱：{content['title']}<br>"
                    f"作者：{content['author']}<br>"
                    f"尺寸：{content['size']}<br>"
                    f"創作媒材：{content['medium']}"
                )
            )
            card_list.append(card)
        return card_list

    # -------- PNG Download --------
    @output
    @render.download(filename="cards.html")
    def download_png():
        # For true PNG export you'd normally use:
        # - playwright
        # - selenium
        # - html2image
        #
        # For now export static HTML snapshot
        html = "<html><body>"
        for cid, content in cards.get().items():
            html += f"""
            <div style="width:105mm;height:74.25mm;
                        background-image:url('template.png');
                        background-size:cover;
                        padding:12mm 8mm;
                        box-sizing:border-box;">
                <div style="margin-top:6mm;font-size:10pt;">
                作品名稱：{content['title']}<br>
                作者：{content['author']}<br>
                尺寸：{content['size']}<br>
                創作媒材：{content['medium']}
                </div>
            </div>
            <br>
            """
        html += "</body></html>"
        yield html

    # -------- PDF Download --------
    @output
    @render.download(filename="cards.pdf")
    def download_pdf():

        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))

        temp = tempfile.NamedTemporaryFile(delete=False)
        doc = SimpleDocTemplate(
            temp.name,
            pagesize=A4
        )

        elements = []

        style = ParagraphStyle(
            name="Chinese",
            fontName="STSong-Light",
            fontSize=10
        )

        for cid, content in cards.get().items():
            text = f"""
            作品名稱：{content['title']}<br/>
            作者：{content['author']}<br/>
            尺寸：{content['size']}<br/>
            創作媒材：{content['medium']}
            """

            elements.append(Paragraph(text, style))
            elements.append(Spacer(1, 20))

        doc.build(elements)

        with open(temp.name, "rb") as f:
            yield f.read()

# ---------- App ----------
app = App(app_ui, server)