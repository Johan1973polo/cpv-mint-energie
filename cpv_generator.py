"""
Générateur de CPV pré-rempli
"""
from datetime import datetime


class CPVGenerator:
    """Génère un document CPV pré-rempli à partir des données extraites"""

    def __init__(self, template_path):
        """
        Args:
            template_path: Chemin vers le template CPV
        """
        self.template_path = template_path
        with open(template_path, 'r', encoding='utf-8') as f:
            self.template = f.read()

    def generate(self, extracted_data, form_data):
        """
        Génère le CPV pré-rempli

        Args:
            extracted_data: Données extraites des PDFs
            form_data: Données saisies manuellement dans le formulaire

        Returns:
            str: Contenu du CPV pré-rempli
        """
        cpv_content = self.template

        # Mapping des champs à remplir
        replacements = self._build_replacements(extracted_data, form_data)

        # Remplacement dans le template
        for placeholder, value in replacements.items():
            cpv_content = cpv_content.replace(placeholder, str(value))

        return cpv_content

    def _build_replacements(self, extracted_data, form_data):
        """Construit le dictionnaire de remplacement"""

        # Récupération des données
        raison_sociale = extracted_data.get('raison_sociale', 'A COMPLETER')
        signataire_nom = extracted_data.get('signataire_nom', 'A COMPLETER')
        email = extracted_data.get('email', 'A COMPLETER')
        telephone = extracted_data.get('telephone', 'A COMPLETER')
        siren = extracted_data.get('siren', 'A COMPLETER')
        siret = extracted_data.get('siret_complet', 'A COMPLETER')
        adresse_siege = extracted_data.get('adresse_siege', 'A COMPLETER')
        adresse_conso = extracted_data.get('adresse_consommation', 'A COMPLETER')
        ville_rcs = extracted_data.get('ville_rcs', 'A COMPLETER')
        forme_juridique = extracted_data.get('forme_juridique', 'A COMPLETER')
        code_ape = extracted_data.get('code_ape', 'A COMPLETER')
        commercial = extracted_data.get('commercial_ohm', 'A COMPLETER')

        # Dates
        date_debut = extracted_data.get('date_debut_livraison', '01/01/2026')
        date_fin = extracted_data.get('date_fin_livraison', '31/12/2026')
        date_signature = datetime.now().strftime('%d/%m/%Y')

        # PDL et puissance
        pdl = extracted_data.get('pdl_principal', 'A COMPLETER')
        puissance = extracted_data.get('puissance_principale', 'A COMPLETER')
        segment = extracted_data.get('segment', 'C4')

        # Volume
        volume_total = extracted_data.get('volume_total', 'A COMPLETER')

        # Données du formulaire
        capital_social = form_data.get('capital_social', 'A COMPLETER')
        fonction_signataire = form_data.get('fonction_signataire', 'A COMPLETER')
        garantie_montant = form_data.get('garantie_montant', '0')
        prix_abonnement = form_data.get('prix_abonnement', 'A COMPLETER')
        prix_capacite = form_data.get('prix_capacite', 'A COMPLETER')
        iban = form_data.get('iban', 'A COMPLETER')
        bic = form_data.get('bic', 'A COMPLETER')

        # Garanties d'origine
        go_souhaite = "☑ Souhaité" if form_data.get('go_souhaite') == 'oui' else "☑ Non souhaité"
        go_non_souhaite = "☐ Souhaité" if form_data.get('go_souhaite') == 'oui' else "☐ Non souhaité"
        go_percentage = form_data.get('go_percentage', '')

        # CEE
        cee_soumis = "☑ Soumis" if form_data.get('cee') == 'soumis' else "☐ Soumis"
        cee_non_soumis = "☐ Non Soumis" if form_data.get('cee') == 'soumis' else "☑ Non Soumis"

        replacements = {
            # En-tête
            'mallet': raison_sociale,
            'Nom du client : mallet': f'Nom du client : {raison_sociale}',

            # Partie client (ligne 20)
            'NOM CLIENT': raison_sociale,
            'Type société': forme_juridique,
            'Montant Capital Social': capital_social,
            'Adresse Siège Social Client': adresse_siege,
            'Ville RCS': ville_rcs,
            'Numéro RCS': siren,
            'Nom Prénom signataire': signataire_nom,
            'Fonction signataire': fonction_signataire,
            'Nom Client': raison_sociale,

            # Dates
            '01/01/2026ut_Fourniture': date_debut,
            'Date_Fin_ Fourniture': date_fin,
            'Date': date_signature,

            # Garantie
            'XXX €': f'{garantie_montant} €',

            # Facturation
            'A compléter par le Client': email,

            # Interlocuteurs
            'A compléter par le Client': signataire_nom,

            # Signature
            'Nom Signataire': signataire_nom,
            'Fonction Signataire': fonction_signataire,

            # Annexe 1 - PRM
            'PRM n°1': pdl,
            '…': segment,  # Cette ligne sera à affiner

            # Volume
            '… MWh': f'{volume_total} MWh',

            # Dates Annexe
            'jj/mm/aaaa': date_debut.replace('/', '/') if '/' in date_debut else date_debut,

            # IBAN/BIC
            '_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _': iban,
            '_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _': bic,

            # GO
            '☐ Souhaité': go_souhaite,
            '☐ Non souhaité': go_non_souhaite,

            # CEE
            '☐ Soumis': cee_soumis,
            '☐ Non Soumis': cee_non_soumis,
        }

        # Ajout des données de consommation selon le segment
        replacements.update(self._build_consumption_replacements(form_data, segment))

        # Ajout des prix selon le segment
        replacements.update(self._build_price_replacements(form_data, segment))

        return replacements

    def _build_consumption_replacements(self, form_data, segment):
        """Construit les remplacements pour les données de consommation"""
        replacements = {}

        if segment == 'C2':
            # 5 postes
            replacements['PTE'] = form_data.get('conso_pte', '0')
            replacements['HPH'] = form_data.get('conso_hph', '0')
            replacements['HCH'] = form_data.get('conso_hch', '0')
            replacements['HPE'] = form_data.get('conso_hpe', '0')
            replacements['HCE'] = form_data.get('conso_hce', '0')

        elif segment == 'C4':
            # 4 postes (pas de PTE)
            replacements['HPH'] = form_data.get('conso_hph', '0')
            replacements['HCH'] = form_data.get('conso_hch', '0')
            replacements['HPE'] = form_data.get('conso_hpe', '0')
            replacements['HCE'] = form_data.get('conso_hce', '0')

        elif segment == 'C5':
            option = form_data.get('c5_option', 'base')
            if option == 'base':
                replacements['PTE'] = form_data.get('conso_base', '0')
            elif option == 'hphc':
                replacements['HPH'] = form_data.get('conso_hp', '0')
                replacements['HCH'] = form_data.get('conso_hc', '0')

        return replacements

    def _build_price_replacements(self, form_data, segment):
        """Construit les remplacements pour les prix"""
        replacements = {}

        # Prix abonnement commun
        replacements['€/an'] = form_data.get('prix_abonnement', 'A COMPLETER')

        if segment == 'C2':
            # 5 postes
            replacements['Prix PTE'] = form_data.get('prix_pte', 'A COMPLETER')
            replacements['Prix HPH'] = form_data.get('prix_hph', 'A COMPLETER')
            replacements['Prix HCH'] = form_data.get('prix_hch', 'A COMPLETER')
            replacements['Prix HPE'] = form_data.get('prix_hpe', 'A COMPLETER')
            replacements['Prix HCE'] = form_data.get('prix_hce', 'A COMPLETER')

        elif segment == 'C4':
            # 4 postes
            replacements['Prix HPH'] = form_data.get('prix_hph', 'A COMPLETER')
            replacements['Prix HCH'] = form_data.get('prix_hch', 'A COMPLETER')
            replacements['Prix HPE'] = form_data.get('prix_hpe', 'A COMPLETER')
            replacements['Prix HCE'] = form_data.get('prix_hce', 'A COMPLETER')

        elif segment == 'C5':
            option = form_data.get('c5_option', 'base')
            if option == 'base':
                replacements['Prix BASE'] = form_data.get('prix_base', 'A COMPLETER')
            elif option == 'hphc':
                replacements['Prix HP'] = form_data.get('prix_hp', 'A COMPLETER')
                replacements['Prix HC'] = form_data.get('prix_hc', 'A COMPLETER')

        return replacements

    def save(self, content, output_path):
        """Sauvegarde le CPV généré"""
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✅ CPV généré: {output_path}")
