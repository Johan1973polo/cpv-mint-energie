"""
Générateur de CPV au format PDF
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime


class PDFCPVGenerator:
    """Génère un CPV au format PDF"""

    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._setup_styles()

    def _setup_styles(self):
        """Configure les styles personnalisés"""
        # Titre principal
        self.styles.add(ParagraphStyle(
            name='CPVTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#667eea'),
            spaceAfter=30,
            alignment=1  # Centré
        ))

        # Sous-titre
        self.styles.add(ParagraphStyle(
            name='CPVSubtitle',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=colors.HexColor('#764ba2'),
            spaceAfter=20
        ))

        # Section
        self.styles.add(ParagraphStyle(
            name='CPVSection',
            parent=self.styles['Heading3'],
            fontSize=12,
            textColor=colors.HexColor('#667eea'),
            spaceBefore=15,
            spaceAfter=10
        ))

    def generate(self, output_path, extracted_data, form_data):
        """
        Génère le CPV PDF

        Args:
            output_path: Chemin du fichier PDF à créer
            extracted_data: Données extraites des PDFs
            form_data: Données du formulaire
        """
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            rightMargin=2*cm,
            leftMargin=2*cm,
            topMargin=2*cm,
            bottomMargin=2*cm
        )

        # Construction du contenu
        story = []

        # En-tête
        story.append(Paragraph("CONTRAT DE FOURNITURE D'ÉLECTRICITÉ", self.styles['CPVTitle']))
        story.append(Paragraph("Conditions Particulières de Vente - Offre Maîtrise", self.styles['CPVSubtitle']))
        story.append(Spacer(1, 0.5*cm))

        # Informations client
        story.extend(self._build_client_info(extracted_data))

        # Informations contractuelles
        story.extend(self._build_contract_info(extracted_data, form_data))

        # Données énergétiques
        story.extend(self._build_energy_data(extracted_data, form_data))

        # Données tarifaires
        story.extend(self._build_pricing_data(form_data))

        # Coordonnées bancaires
        story.extend(self._build_banking_info(form_data))

        # Générer le PDF
        doc.build(story)
        print(f"✅ PDF CPV généré: {output_path}")

    def _build_client_info(self, data):
        """Construction de la section informations client"""
        elements = []

        elements.append(Paragraph("INFORMATIONS CLIENT", self.styles['CPVSection']))

        client_data = [
            ["Raison Sociale:", data.get('raison_sociale', 'N/A')],
            ["Forme Juridique:", data.get('forme_juridique', 'N/A')],
            ["SIREN:", data.get('siren', 'N/A')],
            ["SIRET:", data.get('siret_complet', 'N/A')],
            ["RCS:", f"{data.get('ville_rcs', 'N/A')} {data.get('siren', '')}"],
            ["Adresse Siège:", data.get('adresse_siege', 'N/A')],
            ["Signataire:", data.get('signataire_nom', 'N/A')],
            ["Fonction:", data.get('fonction_signataire', 'N/A')],
            ["Email:", data.get('email', 'N/A')],
            ["Téléphone:", data.get('telephone', 'N/A')],
        ]

        table = Table(client_data, colWidths=[5*cm, 12*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f8f9ff')),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#667eea')),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))

        elements.append(table)
        elements.append(Spacer(1, 0.5*cm))

        return elements

    def _build_contract_info(self, data, form_data):
        """Construction de la section informations contractuelles"""
        elements = []

        elements.append(Paragraph("INFORMATIONS CONTRACTUELLES", self.styles['CPVSection']))

        contract_data = [
            ["Segment:", data.get('segment', 'N/A')],
            ["Date Début Livraison:", data.get('date_debut_livraison', 'N/A')],
            ["Date Fin Livraison:", data.get('date_fin_livraison', 'N/A')],
            ["PDL Principal:", data.get('pdl_principal', 'N/A')],
            ["Puissance Souscrite:", data.get('puissance_principale', 'N/A')],
            ["Nombre de Sites:", data.get('nombre_pdl', 'N/A')],
        ]

        table = Table(contract_data, colWidths=[5*cm, 12*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f8f9ff')),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#667eea')),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))

        elements.append(table)
        elements.append(Spacer(1, 0.5*cm))

        return elements

    def _build_energy_data(self, data, form_data):
        """Construction de la section données énergétiques"""
        elements = []

        elements.append(Paragraph("DONNÉES ÉNERGÉTIQUES", self.styles['CPVSection']))

        segment = data.get('segment', '').upper()

        energy_data = [["Poste", "Consommation (MWh)"]]

        if segment == 'C2':
            energy_data.extend([
                ["PTE - Pointe", form_data.get('conso_pte', '0')],
                ["HPH - Heures Pleines Hiver", form_data.get('conso_hph', '0')],
                ["HCH - Heures Creuses Hiver", form_data.get('conso_hch', '0')],
                ["HPE - Heures Pleines Été", form_data.get('conso_hpe', '0')],
                ["HCE - Heures Creuses Été", form_data.get('conso_hce', '0')],
            ])
        elif segment == 'C4':
            energy_data.extend([
                ["HPH - Heures Pleines Hiver", form_data.get('conso_hph', '0')],
                ["HCH - Heures Creuses Hiver", form_data.get('conso_hch', '0')],
                ["HPE - Heures Pleines Été", form_data.get('conso_hpe', '0')],
                ["HCE - Heures Creuses Été", form_data.get('conso_hce', '0')],
            ])
        elif segment == 'C5':
            c5_option = form_data.get('c5_option', 'base')
            if c5_option == 'base':
                energy_data.append(["BASE", form_data.get('conso_base', '0')])
            else:
                energy_data.extend([
                    ["HP - Heures Pleines", form_data.get('conso_hp', '0')],
                    ["HC - Heures Creuses", form_data.get('conso_hc', '0')],
                ])

        energy_data.append(["TOTAL", data.get('volume_total', '0')])

        table = Table(energy_data, colWidths=[10*cm, 7*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#667eea')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#fff3cd')),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))

        elements.append(table)
        elements.append(Spacer(1, 0.5*cm))

        return elements

    def _build_pricing_data(self, form_data):
        """Construction de la section données tarifaires"""
        elements = []

        elements.append(Paragraph("DONNÉES TARIFAIRES", self.styles['CPVSection']))

        pricing_data = [
            ["Prix Abonnement Annuel:", f"{form_data.get('prix_abonnement', 'N/A')} €/an"],
            ["Prix Capacité:", f"{form_data.get('prix_capacite', 'N/A')} €/MWh"],
            ["Garanties d'Origine:", "Oui" if form_data.get('go_souhaite') == 'oui' else "Non"],
            ["CEE:", "Soumis" if form_data.get('cee') == 'soumis' else "Non Soumis"],
            ["Garantie de Paiement:", f"{form_data.get('garantie_montant', '0')} €"],
        ]

        table = Table(pricing_data, colWidths=[10*cm, 7*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f8f9ff')),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#667eea')),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))

        elements.append(table)
        elements.append(Spacer(1, 0.5*cm))

        return elements

    def _build_banking_info(self, form_data):
        """Construction de la section coordonnées bancaires"""
        elements = []

        elements.append(Paragraph("COORDONNÉES BANCAIRES", self.styles['CPVSection']))

        banking_data = [
            ["IBAN:", form_data.get('iban', 'N/A')],
            ["BIC:", form_data.get('bic', 'N/A')],
        ]

        table = Table(banking_data, colWidths=[5*cm, 12*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f8f9ff')),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#667eea')),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))

        elements.append(table)
        elements.append(Spacer(1, 0.5*cm))

        # Footer
        elements.append(Spacer(1, 1*cm))
        footer_text = f"Document généré automatiquement le {datetime.now().strftime('%d/%m/%Y à %H:%M')}"
        elements.append(Paragraph(footer_text, self.styles['Normal']))
        elements.append(Paragraph("🤖 Généré avec Claude Code", self.styles['Normal']))

        return elements
