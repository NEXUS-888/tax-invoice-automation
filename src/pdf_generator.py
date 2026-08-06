import os
import re
from xml.sax.saxutils import escape
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors


class PDFGenerator:
    def __init__(self, output_dir, signature_path=None):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)

        # 600 DPI signature block image (white bg, includes For ANANYA + ink + Proprietor)
        base_dir = os.path.dirname(os.path.dirname(__file__))
        self.signature_path = signature_path or os.path.join(base_dir, "data", "signature_final.png")

    def generate_agency_pdf(self, agency_data):
        """
        Generates a professional PDF invoice matching MAHALASA.pdf format.
        Uses a single 600-DPI signature block image — no duplicate vector text.
        """
        safe_sheet_name = re.sub(r'[^A-Za-z0-9_.-]+', '_', str(agency_data['sheet_name'])).strip('_')
        base_file_name = f"Invoice_{agency_data['invoice_no']}_{safe_sheet_name}"
        file_name = f"{base_file_name}.pdf"
        pdf_path = os.path.join(self.output_dir, file_name)

        styles = getSampleStyleSheet()

        # --- Styles matching MAHALASA.pdf ---
        title_style = ParagraphStyle(
            'CompanyTitle', parent=styles['Normal'],
            fontName='Times-Bold', fontSize=24, leading=26,
            alignment=1, textColor=colors.HexColor('#C00000'),
        )
        addr_style = ParagraphStyle(
            'CompanyAddr', parent=styles['Normal'],
            fontName='Times-Bold', fontSize=8.5, leading=10.5,
            alignment=1, textColor=colors.HexColor('#002060'),
        )
        contact_style = ParagraphStyle(
            'CompanyContact', parent=styles['Normal'],
            fontName='Times-Bold', fontSize=9, leading=11,
            alignment=1, textColor=colors.HexColor('#C00000'),
        )
        h2_style = ParagraphStyle(
            'HeaderDesc', parent=styles['Normal'],
            fontName='Helvetica-Bold', fontSize=10, leading=13,
            alignment=1, textColor=colors.black,
        )
        bold_style = ParagraphStyle(
            'BoldTxt', parent=styles['Normal'],
            fontName='Helvetica-Bold', fontSize=8.5, leading=10.5,
            textColor=colors.black,
        )
        norm_style = ParagraphStyle(
            'NormTxt', parent=styles['Normal'],
            fontName='Helvetica', fontSize=8.5, leading=10.5,
            textColor=colors.black,
        )
        center_bold = ParagraphStyle('CenterBold', parent=bold_style, alignment=1)
        center_norm = ParagraphStyle('CenterNorm', parent=norm_style, alignment=1)

        elements = []

        # ── 1. Company header ──────────────────────────────────────────────
        elements.append(Paragraph('ANANYA ENTERPRISES', title_style))
        elements.append(Spacer(1, 2))
        elements.append(Paragraph(
            'No. 10,1st Floor,4th \u2018E\u2019 1st Cross,Magadi Main Road,'
            'Vrushabavathi nagara,Kamakshipalya,Banglore \u2013 560079.',
            addr_style,
        ))
        elements.append(Spacer(1, 2))
        elements.append(Paragraph(
            'Email:ananyaenterprise347@gmail.com'
            ' &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; '
            'Mob: 9141844840,9880183840,7353306462',
            contact_style,
        ))
        elements.append(Spacer(1, 4))
        elements.append(HRFlowable(width='100%', thickness=3, color=colors.black, spaceAfter=6))

        # ── 2. Recipient TO box + metadata grid ───────────────────────────
        addr_lines = [escape(str(a)) for a in agency_data.get('address', []) if a]
        addr_str = "<br/>".join(addr_lines)
        date_str = str(agency_data.get('date', '')).split(' ')[0]

        left_to_content = [
            Paragraph('<b>TO</b>', bold_style),
            Paragraph(f"<b>{escape(str(agency_data['agency_name']))}</b>", bold_style),
            Paragraph(addr_str, norm_style) if addr_str else Paragraph('', norm_style),
        ]

        gstin_str = escape(str(agency_data.get('gstin', '')))
        if gstin_str:
            left_to_content.append(Spacer(1, 2))
            left_to_content.append(Paragraph(
                f"<b>{gstin_str}</b>",
                ParagraphStyle('GSTINStyle', parent=bold_style, fontSize=10),
            ))

        meta_table_data = [
            [Paragraph('<b>DATE:</b>', bold_style), Paragraph(escape(date_str), norm_style)],
            [Paragraph('<b>PAN NO</b>', bold_style), Paragraph(escape(str(agency_data.get('pan_no', ''))), bold_style)],
            [Paragraph('<b>VENDOR CODE</b>', bold_style), Paragraph(escape(str(agency_data.get('vendor_code', ''))), bold_style)],
            [Paragraph('<b>GST NO</b>', bold_style), Paragraph(escape(str(agency_data.get('gst_no', ''))), bold_style)],
            [Paragraph('<b>INVOICE NO</b>', bold_style), Paragraph(f"<b>{agency_data['invoice_no']}</b>", bold_style)],
        ]
        t_meta = Table(meta_table_data, colWidths=[105, 120])
        t_meta.setStyle(TableStyle([
            ('GRID',          (0, 0), (-1, -1), 0.75, colors.black),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
        ]))

        to_box = Table([[left_to_content, t_meta]], colWidths=[305, 230])
        to_box.setStyle(TableStyle([
            ('GRID',          (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING',   (0, 0), (-1, -1), 5),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 5),
        ]))
        elements.append(to_box)
        elements.append(Spacer(1, 6))

        # ── 3. Month description banner ───────────────────────────────────
        elements.append(Paragraph(f"<b>{escape(str(agency_data.get('month_desc', '')))}</b>", h2_style))
        elements.append(Spacer(1, 6))

        # ── 4. Items table ────────────────────────────────────────────────
        headers = ['SL NO', 'HSN/SAC', 'ITEM DESCRIPTION VEHICLE NO',
                    'TOTAL LOADS', 'PER/LOAD AMOUNT', 'TOTAL AMOUNT']
        items_table_data = [[Paragraph(f'<b>{h}</b>', center_bold) for h in headers]]

        for v in agency_data.get('vehicles', []):
            items_table_data.append([
                Paragraph(escape(str(v.get('sl', ''))), center_norm),
                Paragraph(escape(str(v.get('hsn', ''))), center_norm),
                Paragraph(f"<b>{escape(str(v.get('vehicle_no', '')))}</b>", bold_style),
                Paragraph(f"{float(v.get('loads', 0)):.2f}", center_norm),
                Paragraph(f"{float(v.get('rate', 0)):.2f}", center_norm),
                Paragraph(f"<b>{float(v.get('total_amount', 0)):.2f}</b>", center_bold),
            ])

        min_rows = max(5, len(agency_data.get('vehicles', [])) + 1)
        for _ in range(len(items_table_data), min_rows):
            items_table_data.append([''] * 6)

        # Summary rows
        items_table_data.append([
            Paragraph('<b>BANK ACCOUNT DETAILS</b>', bold_style), '', '',
            Paragraph('<b>TOTAL</b>', bold_style), '',
            Paragraph(f"<b>{float(agency_data.get('total', 0)):.2f}</b>", center_bold),
        ])
        items_table_data.append([
            Paragraph('ANANYA ENTERPRISES, HDFC BANK', norm_style), '', '',
            Paragraph('<b>SGST -9%</b>', bold_style), '',
            Paragraph(f"{float(agency_data.get('sgst', 0)):.2f}", center_norm),
        ])
        items_table_data.append([
            Paragraph('ACCOUNT NO : 50200076002672', norm_style), '', '',
            Paragraph('<b>CGST - 9%</b>', bold_style), '',
            Paragraph(f"{float(agency_data.get('cgst', 0)):.2f}", center_norm),
        ])
        items_table_data.append([
            Paragraph('IFSC CODE : HDFC0004261', norm_style), '', '',
            Paragraph('<b>GRAND TOTAL</b>', bold_style), '',
            Paragraph(f"<b>{float(agency_data.get('grand_total', 0)):.2f}</b>", center_bold),
        ])

        words_txt = agency_data.get('in_words', '')
        items_table_data.append([
            Paragraph(f"<b>{escape(str(words_txt))}</b>", center_bold), '', '', '', '', '',
        ])

        t_main = Table(items_table_data, colWidths=[40, 55, 205, 65, 80, 90])
        t_main.setStyle(TableStyle([
            ('GRID',          (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('SPAN', (0, min_rows),   (2, min_rows)),
            ('SPAN', (3, min_rows),   (4, min_rows)),
            ('SPAN', (0, min_rows+1), (2, min_rows+1)),
            ('SPAN', (3, min_rows+1), (4, min_rows+1)),
            ('SPAN', (0, min_rows+2), (2, min_rows+2)),
            ('SPAN', (3, min_rows+2), (4, min_rows+2)),
            ('SPAN', (0, min_rows+3), (2, min_rows+3)),
            ('SPAN', (3, min_rows+3), (4, min_rows+3)),
            ('SPAN', (0, min_rows+4), (-1, min_rows+4)),
        ]))
        elements.append(t_main)
        elements.append(Spacer(1, 6))

        # ── 5. Signature block — single 600-DPI image, right-aligned ─────
        # The image already contains "For ANANYA ENTERPRISES", ink sig, and
        # "Proprietor" — NO extra text is rendered by code.
        sig_path = self.signature_path
        if os.path.exists(sig_path):
            # Original placement in MAHALASA.pdf: 142.1 × 51.2 pt
            sig_img = Image(sig_path, width=142, height=53)
            sig_img.hAlign = 'RIGHT'
            elements.append(sig_img)
        else:
            elements.append(Spacer(1, 55))

        # ── Build PDF ─────────────────────────────────────────────────────
        target_path = pdf_path
        try:
            doc = SimpleDocTemplate(
                target_path, pagesize=A4,
                leftMargin=30, rightMargin=30, topMargin=20, bottomMargin=20,
            )
            doc.build(elements)
        except PermissionError:
            target_path = os.path.join(self.output_dir, f"{base_file_name}_updated.pdf")
            doc = SimpleDocTemplate(
                target_path, pagesize=A4,
                leftMargin=30, rightMargin=30, topMargin=20, bottomMargin=20,
            )
            doc.build(elements)

        return target_path

    def batch_generate(self, agency_data_list):
        """Generates PDF invoices for all agencies in batch."""
        generated_paths = []
        for agency in agency_data_list:
            pdf_path = self.generate_agency_pdf(agency)
            generated_paths.append({
                'sheet_name': agency['sheet_name'],
                'agency_name': agency['agency_name'],
                'invoice_no': agency['invoice_no'],
                'pdf_path': pdf_path,
            })
        return generated_paths
