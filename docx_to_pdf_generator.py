"""
Générateur de CPV utilisant un template DOCX et convertissant en PDF
"""
from docxtpl import DocxTemplate
from datetime import datetime
import subprocess
import platform


class DOCXtoPDFGenerator:
    """Génère un CPV à partir d'un template DOCX et le convertit en PDF"""

    def __init__(self, template_path):
        """
        Args:
            template_path: Chemin vers le template DOCX
        """
        self.template_path = template_path
        self.template = DocxTemplate(template_path)

    def generate_docx(self, output_docx_path, extracted_data, form_data):
        """
        Génère un DOCX pré-rempli

        Args:
            output_docx_path: Chemin du fichier DOCX à créer
            extracted_data: Données extraites des PDFs
            form_data: Données du formulaire
        """
        # Préparer le contexte de remplacement
        context = self._build_context(extracted_data, form_data)

        # Remplir le template
        self.template.render(context)

        # Sauvegarder le DOCX
        self.template.save(output_docx_path)
        print(f"✅ DOCX généré: {output_docx_path}")

        return output_docx_path

    def convert_to_pdf(self, docx_path, pdf_path):
        """
        Convertit un DOCX en PDF

        Args:
            docx_path: Chemin du fichier DOCX
            pdf_path: Chemin du fichier PDF à créer
        """
        try:
            # Sur macOS, utiliser LibreOffice si disponible
            if platform.system() == 'Darwin':
                # Essayer avec LibreOffice
                try:
                    subprocess.run([
                        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                        '--headless',
                        '--convert-to', 'pdf',
                        '--outdir', str(pdf_path.parent) if hasattr(pdf_path, 'parent') else '.',
                        docx_path
                    ], check=True, capture_output=True)
                    print(f"✅ PDF généré avec LibreOffice: {pdf_path}")
                    return pdf_path
                except (FileNotFoundError, subprocess.CalledProcessError):
                    print("⚠️ LibreOffice non trouvé, génération du DOCX uniquement")
                    return docx_path

            # Sur Windows, utiliser Microsoft Word via COM si disponible
            elif platform.system() == 'Windows':
                try:
                    import win32com.client
                    word = win32com.client.Dispatch('Word.Application')
                    doc = word.Documents.Open(docx_path)
                    doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF
                    doc.Close()
                    word.Quit()
                    print(f"✅ PDF généré avec Word: {pdf_path}")
                    return pdf_path
                except:
                    print("⚠️ Microsoft Word non disponible, génération du DOCX uniquement")
                    return docx_path

            # Fallback: retourner le DOCX
            return docx_path

        except Exception as e:
            print(f"⚠️ Erreur lors de la conversion PDF: {e}")
            print("📄 Le fichier DOCX sera fourni à la place")
            return docx_path

    def generate(self, output_path, extracted_data, form_data):
        """
        Génère le CPV final (DOCX ou PDF selon disponibilité)

        Args:
            output_path: Chemin du fichier final (avec extension .pdf ou .docx)
            extracted_data: Données extraites
            form_data: Données du formulaire

        Returns:
            str: Chemin du fichier généré
        """
        # Générer le DOCX d'abord
        docx_path = str(output_path).replace('.pdf', '.docx')
        self.generate_docx(docx_path, extracted_data, form_data)

        # Si l'utilisateur veut un PDF, essayer de convertir
        if output_path.endswith('.pdf'):
            return self.convert_to_pdf(docx_path, output_path)
        else:
            return docx_path

    def _build_context(self, extracted_data, form_data):
        """Construit le contexte pour le remplacement dans le template"""

        # Récupération des données
        segment = extracted_data.get('segment', '').upper()

        context = {
            # Informations client
            'raison_sociale': extracted_data.get('raison_sociale', 'A COMPLETER'),
            'forme_juridique': extracted_data.get('forme_juridique', 'A COMPLETER'),
            'capital_social': form_data.get('capital_social', extracted_data.get('capital_social', 'A COMPLETER')),
            'adresse_siege': extracted_data.get('adresse_siege', 'A COMPLETER'),
            'ville_rcs': extracted_data.get('ville_rcs', 'A COMPLETER'),
            'numero_rcs': extracted_data.get('siren', 'A COMPLETER'),
            'siren': extracted_data.get('siren', 'A COMPLETER'),
            'siret': extracted_data.get('siret_complet', 'A COMPLETER'),

            # Signataire
            'nom_signataire': extracted_data.get('signataire_nom', 'A COMPLETER'),
            'fonction_signataire': form_data.get('fonction_signataire', extracted_data.get('fonction_signataire', 'A COMPLETER')),

            # Contact
            'email': extracted_data.get('email', 'A COMPLETER'),
            'telephone': extracted_data.get('telephone', 'A COMPLETER'),

            # Dates
            'date_signature': datetime.now().strftime('%d/%m/%Y'),
            'date_debut_livraison': extracted_data.get('date_debut_livraison', '01/01/2026'),
            'date_fin_livraison': extracted_data.get('date_fin_livraison', '31/12/2026'),

            # PDL et puissance
            'pdl_principal': extracted_data.get('pdl_principal', 'A COMPLETER'),
            'puissance_principale': extracted_data.get('puissance_principale', 'A COMPLETER'),
            'segment': segment,
            'nombre_pdl': extracted_data.get('nombre_pdl', '1'),

            # Volume
            'volume_total': extracted_data.get('volume_total', 'A COMPLETER'),

            # Adresse de facturation
            'adresse_facturation': extracted_data.get('adresse_siege', 'A COMPLETER'),
            'adresse_consommation': extracted_data.get('adresse_consommation', 'A COMPLETER'),

            # Interlocuteur facturation
            'interlocuteur_facturation': extracted_data.get('signataire_nom', 'A COMPLETER'),

            # Garantie
            'montant_garantie': form_data.get('garantie_montant', '0'),

            # Prix
            'prix_abonnement': form_data.get('prix_abonnement', 'A COMPLETER'),
            'prix_capacite': form_data.get('prix_capacite', 'A COMPLETER'),

            # GO
            'go_souhaite': 'Oui' if form_data.get('go_souhaite') == 'oui' else 'Non',
            'go_percentage': form_data.get('go_percentage', ''),

            # CEE
            'cee_soumis': 'Oui' if form_data.get('cee') == 'soumis' else 'Non',

            # IBAN/BIC
            'iban': form_data.get('iban', 'A COMPLETER'),
            'bic': form_data.get('bic', 'A COMPLETER'),
        }

        # Ajouter les consommations selon le segment
        if segment == 'C2':
            context.update({
                'conso_pte': form_data.get('conso_pte', '0'),
                'conso_hph': form_data.get('conso_hph', '0'),
                'conso_hch': form_data.get('conso_hch', '0'),
                'conso_hpe': form_data.get('conso_hpe', '0'),
                'conso_hce': form_data.get('conso_hce', '0'),
                'prix_pte': form_data.get('prix_pte', 'A COMPLETER'),
                'prix_hph': form_data.get('prix_hph', 'A COMPLETER'),
                'prix_hch': form_data.get('prix_hch', 'A COMPLETER'),
                'prix_hpe': form_data.get('prix_hpe', 'A COMPLETER'),
                'prix_hce': form_data.get('prix_hce', 'A COMPLETER'),
            })
        elif segment == 'C4':
            context.update({
                'conso_hph': form_data.get('conso_hph', '0'),
                'conso_hch': form_data.get('conso_hch', '0'),
                'conso_hpe': form_data.get('conso_hpe', '0'),
                'conso_hce': form_data.get('conso_hce', '0'),
                'prix_hph': form_data.get('prix_hph', 'A COMPLETER'),
                'prix_hch': form_data.get('prix_hch', 'A COMPLETER'),
                'prix_hpe': form_data.get('prix_hpe', 'A COMPLETER'),
                'prix_hce': form_data.get('prix_hce', 'A COMPLETER'),
            })
        elif segment == 'C5':
            c5_option = form_data.get('c5_option', 'base')
            if c5_option == 'base':
                context.update({
                    'conso_base': form_data.get('conso_base', '0'),
                    'prix_base': form_data.get('prix_base', 'A COMPLETER'),
                })
            else:
                context.update({
                    'conso_hp': form_data.get('conso_hp', '0'),
                    'conso_hc': form_data.get('conso_hc', '0'),
                    'prix_hp': form_data.get('prix_hp', 'A COMPLETER'),
                    'prix_hc': form_data.get('prix_hc', 'A COMPLETER'),
                })

        return context
