from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageTemplate, Frame
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from datetime import datetime

# Define add_footer FIRST
def add_footer(canvas, doc):
    canvas.saveState()
    
    # Set footer styles
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.black)
    
    # Footer content with absolute positioning
    page_width, page_height = letter
    margin = 0.5 * inch
    
    # Company Name
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawRightString(
        page_width - margin,
        0.7 * inch,
        "AGROPARTNERS S.R.L."
    )
    
    # Horizontal line
    canvas.line(
        margin,
        0.6 * inch,
        page_width - margin,
        0.6 * inch
    )
    
    # Contact info
    canvas.setFont("Helvetica", 8)
    info = [
        "4to anillo entre mutualista y paragua",
        "Telf: 591-3-3369000 FAX: 3620100",
        "Santa Cruz - Bolivia"
    ]
    
    y_position = 0.45 * inch
    for line in reversed(info):
        canvas.drawRightString(page_width - margin, y_position, line)
        y_position -= 0.15 * inch

    canvas.restoreState()

def generate_printer_report(filename, filtered_data, all_data, image_path):
    PAGE_WIDTH, PAGE_HEIGHT = letter
    MARGIN = 0.5 * inch
    
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=0.25 * inch,
        bottomMargin=1 * inch  # Space reserved for footer
    )
    
    # Create page template with footer
    frame = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        PAGE_HEIGHT - doc.topMargin - doc.bottomMargin,
        id='main_frame'
    )
    
    doc.addPageTemplates([
        PageTemplate(
            id='AllPages',
            frames=frame,
            onPage=add_footer,
            pagesize=letter
        )
    ])
    
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='TableHeader',
        alignment=TA_CENTER,
        fontSize=7,
        fontName='Helvetica-Bold',
        leading=9
    ))
    
    story = []
    
    # Header Section
    img_width = PAGE_WIDTH - 2 * MARGIN
    header_img = Image(image_path, width=img_width, height=1 * inch)
    story.append(header_img)
    story.append(Spacer(1, 0.2 * inch))
    
    # Title Section
    title = Paragraph("<b>REPORTE DE IMPRESORAS</b>", styles['Title'])
    story.append(title)
    date_str = datetime.now().strftime("%d/%m/%Y")
    date_para = Paragraph(f"<font size=10>Fecha: {date_str}</font>", styles['Normal'])
    story.append(date_para)
    story.append(Spacer(1, 0.1 * inch))
    story.append(Table([[""]], colWidths=[img_width], 
                     style=[('LINEABOVE', (0,0), (-1,-1), 1, colors.black)]))
    
    # Cambios a Realizar Section
    cambios_title = Paragraph("<b>Cambios a Realizar</b>", styles['Heading3'])
    story.append(cambios_title)
    story.append(Spacer(1, 0.2 * inch))
    
    # Table Configuration
    headers = [
        "IP", "Tinta\nNegra", "Cian", "Amarillo", "Magenta",
        "Kit\nde\nMantenimiento", "Unidad\nde Imagen", "Estado", "Area"
    ]
    
    col_widths = [
        0.97*inch, 0.6*inch, 0.5*inch, 0.7*inch,
        0.7*inch, 0.9*inch, 0.9*inch, 0.6*inch, 0.97*inch
    ]
    # Adds percent sign and None items adds N/A value
    def process_data(data):
        table_data = [headers]
        for item in data:
            row = [
                item['ip'],
                f"{item['black_ink']}%" if item['black_ink'] else 'N/A',
                f"{item['cyan_ink']}%" if item['cyan_ink'] else 'N/A',
                f"{item['yellow_ink']}%" if item['yellow_ink'] else 'N/A',
                f"{item['magenta_ink']}%" if item['magenta_ink'] else 'N/A',
                f"{item['maintenance_kit']}%" if item['maintenance_kit'] else 'N/A',
                f"{item['imaging_unit']}%" if item['imaging_unit'] else 'N/A',
                'N/A' if item['status'] == None else item['status'],
                item['area']
            ]
            table_data.append(row)
        return table_data
    
    def add_conditional_styles(data):
        styles = []
        for row_idx in range(1, len(data)):
            for col_idx in [1, 2, 3, 4, 5, 6]:
                cell_value = data[row_idx][col_idx]
                if cell_value != 'N/A':
                    try:
                        value = int(cell_value.strip('%'))
                        if value < 30:
                            styles.extend([
                                ('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), colors.red),
                                ('TEXTCOLOR', (col_idx, row_idx), (col_idx, row_idx), colors.white)
                            ])
                    except (ValueError, AttributeError):
                        pass
            
            if data[row_idx][7] != "OK" and data[row_idx][7] != "N/A":
                styles.extend([
                    ('BACKGROUND', (7, row_idx), (7, row_idx), colors.red),
                    ('TEXTCOLOR', (7, row_idx), (7, row_idx), colors.white)
                ])
        return styles
    
    # First Table
    cambios_table_data = process_data(filtered_data)
    table_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 7),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('WORDWRAP', (0,0), (-1,-1)),
    ]
    table_style += add_conditional_styles(cambios_table_data)
    
    cambios_table = Table(cambios_table_data, colWidths=col_widths, repeatRows=1)
    cambios_table.setStyle(TableStyle(table_style))
    story.append(cambios_table)
    
    # Separator
    story.append(Spacer(1, 0.2 * inch))
    story.append(Table([[""]], colWidths=[img_width], 
                     style=[('LINEABOVE', (0,0), (-1,-1), 1, colors.black)]))
    
    # Full Report Section
    reporte_title = Paragraph("<b>Reporte Completo</b>", styles['Heading3'])
    story.append(reporte_title)
    story.append(Spacer(1, 0.2 * inch))
    
    # Correct the second table data
    # Second Table
    full_table_data = process_data(all_data)
    table_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 7),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('WORDWRAP', (0,0), (-1,-1)),
    ]
    table_style += add_conditional_styles(full_table_data)
    full_table = Table(full_table_data, colWidths=col_widths, repeatRows=1)
    full_table.setStyle(TableStyle(table_style))
    story.append(full_table)
    
    doc.build(story)

