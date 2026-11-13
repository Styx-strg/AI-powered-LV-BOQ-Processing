from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    HRFlowable
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from xml.sax.saxutils import escape
from decimal import Decimal
from pathlib import Path

def create_validation_pdf(
    validation_json,
    ai_report=None,
    Extraktionskommentare= None,
    output_path="Validation.pdf",
    position_count=None,
    Bieterkommentare=None,
    history=None
):

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=23 * mm,
        bottomMargin=20 * mm
    )

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="ReportTitle",
        parent=styles["Title"],
        fontSize=22,
        leading=27,
        alignment=0,
        textColor=colors.HexColor("#243447"),
        spaceAfter=6
    ))

    styles.add(ParagraphStyle(
        name="Section",
        parent=styles["Heading2"],
        fontSize=13,
        leading=17,
        textColor=colors.HexColor("#243447"),
        spaceBefore=15,
        spaceAfter=8
    ))

    styles.add(ParagraphStyle(
        name="Cell",
        parent=styles["Normal"],
        fontSize=9,
        leading=12
    ))

    styles.add(ParagraphStyle(
        name="SmallReport",
        parent=styles["Normal"],
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#555555")
    ))

    styles.add(ParagraphStyle(
        name="TableHeader",
        parent=styles["Normal"],
        fontSize=9,
        leading=12,
        textColor=colors.white,
        fontName="Helvetica-Bold"
    ))

    elements = []

    def paragraph(value, style="Cell"):
        return Paragraph(escape(str(value)), styles[style])

    def format_eur(value):
        if value is None:
            return "Nicht verfügbar"

        try:
            value = Decimal(str(value))
            formatted = f"{value:,.2f}"
            formatted = formatted.replace(",", "X")
            formatted = formatted.replace(".", ",")
            formatted = formatted.replace("X", ".")
            return formatted + " €"
        except Exception:
            return str(value)

    def create_table(headers, rows, widths):

        data = [
            [paragraph(header, "TableHeader") for header in headers]
        ]

        for row in rows:
            data.append([
                paragraph(value) for value in row
            ])

        table = Table(
            data,
            colWidths=widths,
            repeatRows=1,
            splitByRow=1
        )

        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0),
             colors.HexColor("#243447")),

            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),

            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#F2F4F6")]),

            ("GRID", (0, 0), (-1, -1), 0.4,
             colors.HexColor("#D5D9DE")),

            ("VALIGN", (0, 0), (-1, -1), "TOP"),

            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))

        return table

    positions = validation_json.get("positions", {})
    prices = validation_json.get("prices", {})

    existing = positions.get("existing", [])
    missing = positions.get("missing", [])
    unknown = positions.get("unknown", [])
    duplicates = positions.get("duplicates", {})
    orphans = positions.get("orphans", [])
    offered_total = prices.get("AI_scanned_total")
    gp_errors = prices.get("gp_errors", [])
    incomplete = prices.get("incomplete", [])
    price_valid = prices.get("prices_valid")
    calculated_total = prices.get("calculated_total")

    elements.append(paragraph("Angebotsvalidierung", "ReportTitle"))
    elements.append(paragraph("Automatisch erstellter Prüfbericht zur " "Auswertung des Anbieterangebots.", "SmallReport"))
    elements.append(Spacer(1, 8))

    elements.append(HRFlowable(
        width="100%",
        thickness=1,
        color=colors.HexColor("#D5D9DE")
    ))

    elements.append(
        paragraph("1. Zusammenfassung", "Section")
    )

    overall_valid = (
        price_valid is True
        and len(missing) == 0
        and len(duplicates) == 0
        and len(orphans) == 0
        and len(incomplete) == 0
    )

    status = (
        "Keine Auffälligkeiten"
        if overall_valid
        else "Manuelle Prüfung erforderlich"
    )

    summary_rows = [
        ["Positionen im LV", position_count],
        ["Erkannte Positionen", existing],
        ["Fehlende Positionen", len(missing)],
        ["Unbekannte Positionen", len(unknown)],
        ["Doppelte Positionen", len(duplicates)],
        ["Positionen ohne Positionsnummer", len(orphans)],
        ["GP-Abweichungen", len(gp_errors)],
        ["Unvollständige Angaben", len(incomplete)],
        ["Gesamtsummenprüfung", status],
        ["Selbst berechneter Angebotspreis",
         format_eur(calculated_total)]
    ]

    if offered_total is not None:
        summary_rows.insert(
            -1,
            ["Angebotspreis laut KI", format_eur(offered_total)]
        )

    elements.append(create_table(
        ["Prüfung", "Ergebnis"],
        summary_rows,
        [105 * mm, 69 * mm]
    ))
    elements.append(paragraph("2. Positionsvalidierung", "Section"))

    def add_position_section(title, positions_list):
        if positions_list:
            rows = [[pos] for pos in positions_list]
            elements.append(create_table(
                [title],
                rows,
                [174 * mm]
            ))
        else:
            elements.append(
                paragraph("Keine Auffälligkeiten gefunden.")
            )

        elements.append(Spacer(1, 10))

    add_position_section("Fehlende Positionen", missing)

    # add_position_section("Unbekannte Positionen", unknown)

    if duplicates:
        duplicate_rows = [
            [pos, count]
            for pos, count in duplicates.items()
        ]
        elements.append(create_table(
            ["Doppelte Positionen", "Anzahl der Erkennungen"],
            duplicate_rows,
            [100 * mm, 74 * mm]
        ))
    else:
        elements.append(paragraph("Keine Duplikate gefunden."))

    elements.append(Spacer(1, 10))

    elements.append(
        paragraph(
            f"Positionen ohne erkannte Positionsnummer: "
            f"{len(orphans)}",
            "SmallReport"
        )
    )

    elements.append(paragraph("3. Preisvalidierung", "Section"))

    add_position_section(
        "Positionen mit GP-Abweichungen",
        gp_errors
    )
    add_position_section(
        "Unvollständige Positionen / Angaben",
        incomplete
    )

    elements.append(
        paragraph(
            "Der Gesamtpreis wurde anhand der erkannten "
            "Einheitspreise und der Mengen des Soll-LV "
            "selbstständig berechnet.",
            "SmallReport"
        )
    )

    if Bieterkommentare is not None and len(Bieterkommentare) != 0:

        elements.append(Spacer(1, 5))

        elements.append(
            paragraph("4. Bieterkommentare", "Section")
        )
        BK_rows = []
        # dictionarys und listen absichern
        for kommentar in Bieterkommentare:
            BK_rows.append([kommentar["positionsnummer"], kommentar["kommentar"]])
        elements.append(create_table(
            ["Position", "Kommentar"],
            BK_rows,
            [69 * mm, 105 * mm]
        ))
        elements.append(Spacer(1, 5))

    if ai_report is not None:

        elements.append(PageBreak())

        elements.append(
            paragraph("5. KI-Prüfbericht", "Section")
        )

        if hasattr(ai_report, "model_dump"):
            ai_report = ai_report.model_dump()

        if isinstance(ai_report, dict):
            ai_lines = [
                f"{key}: {value}"
                for key, value in ai_report.items()
            ]

        elif isinstance(ai_report, list):
            ai_lines = [str(item) for item in ai_report]

        else:
            ai_lines = str(ai_report).splitlines()

        for line in ai_lines:

            if line.strip():
                elements.append(
                    paragraph(line)
                )
                elements.append(Spacer(1, 5))

    if Extraktionskommentare is not None:
        elements.append(
        paragraph("Extraktionskommentare:", "Section")
        )
        elements.append(
                paragraph(str(Extraktionskommentare))
        )
        elements.append(Spacer(1, 5))

    if history:
        elements.append(PageBreak())

        elements.append(
            paragraph("Änderungshistorie des Agenten:", "Section")
        )

        if isinstance(history, list):
            history_lines = [str(item) for item in history]

        elif isinstance(history, dict):
            history_lines = [
                f"{key}: {value}"
                for key, value in history.items()
            ]

        else:
            history_lines = str(history).splitlines()

        for line in history_lines:
            if line.strip():
                elements.append(paragraph(line))
                elements.append(Spacer(1, 5))

    def page_decoration(canvas, doc):
        canvas.saveState()
        width, height = A4
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#777777"))

        canvas.drawString(
            18 * mm,
            height - 12 * mm,
            "ANGEBOTSPRÜFUNG"
        )

        canvas.drawRightString(
            width - 18 * mm,
            10 * mm,
            f"Seite {doc.page}"
        )

        canvas.setStrokeColor(colors.HexColor("#D5D9DE"))

        canvas.line(
            18 * mm,
            height - 15 * mm,
            width - 18 * mm,
            height - 15 * mm
        )

        canvas.restoreState()

    Path(output_path).parent.mkdir(
        parents=True,
        exist_ok=True
    )

    doc.build(
        elements,
        onFirstPage=page_decoration,
        onLaterPages=page_decoration
    )

    return output_path
