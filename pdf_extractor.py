"""
Module d'extraction des données depuis les PDFs
"""
import re
import pdfplumber
from pathlib import Path


class PDFExtractor:
    """Extracteur de données depuis les PDFs Fiche, RGPD et SIREN"""

    def __init__(self):
        self.data = {}

    def extract_fiche(self, pdf_path):
        """Extrait les données de la Fiche Contact"""
        print(f"📄 Extraction de la Fiche: {pdf_path}")

        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        # Extraction des données
        self.data['signataire_nom'] = self._extract_pattern(text, r'Nom/Prénom:\s*(.+?)(?:\n|$)')
        self.data['raison_sociale'] = self._extract_pattern(text, r'Raison Sociale:\s*(.+?)(?:\n|$)')
        self.data['email'] = self._extract_pattern(text, r'Email:\s*(.+?)(?:\n|$)')
        self.data['telephone'] = self._extract_pattern(text, r'Téléphone:\s*(.+?)(?:\n|$)')
        self.data['siren'] = self._extract_pattern(text, r'SIREN:\s*(\d+)')
        self.data['adresse_siege'] = self._extract_pattern(text, r'Adresse:\s*(.+?)(?:\n|Score)')
        self.data['score'] = self._extract_pattern(text, r'Score:\s*([0-9./]+)')

        # Informations commerciales
        self.data['commercial_ohm'] = self._extract_pattern(text, r'Commercial OHM:\s*(.+?)(?:\n|$)')
        self.data['courtier'] = self._extract_pattern(text, r'Courtier:\s*(.+?)(?:\n|$)')
        self.data['courtier_final'] = self._extract_pattern(text, r'Courtier Final:\s*(.+?)(?:\n|$)')

        # Points de livraison
        self.data['pdl_principal'] = self._extract_pattern(text, r'PDL Principal:\s*(\d+)')
        self.data['puissance_principale'] = self._extract_pattern(text, r'Puissance Principale:\s*(.+?)(?:\n|$)')
        self.data['nombre_pdl'] = self._extract_pattern(text, r'Nombre de Points de Livraison:\s*(\d+)')
        self.data['segment'] = self._extract_pattern(text, r'Segment Principal:\s*(\w+)')

        # Données énergétiques
        self.data['volume_total'] = self._extract_pattern(text, r'Volume Total:\s*([\d.,]+)\s*MWh')
        self.data['prix_pondere'] = self._extract_pattern(text, r'Prix Pondéré Moyen:\s*([\d.]+)')

        # Dates
        self.data['date_signature'] = self._extract_pattern(text, r'Date Signature:\s*(.+?)(?:\n|$)')
        self.data['date_debut_livraison'] = self._extract_pattern(text, r'Date Début Livraison:\s*(.+?)(?:\n|$)')
        self.data['date_fin_livraison'] = self._extract_pattern(text, r'Date Fin Livraison:\s*(.+?)(?:\n|$)')

        # Informations contractuelles
        self.data['type_contrat'] = self._extract_pattern(text, r'Type de Contrat:\s*(.+?)(?:\n|$)')
        self.data['typologie_contrat'] = self._extract_pattern(text, r'Typologie Contrat:\s*(.+?)(?:\n|$)')
        self.data['frais_abonnement'] = self._extract_pattern(text, r'Frais Abonnement:\s*(.+?)(?:\n|$)')
        self.data['depot_garantie'] = self._extract_pattern(text, r'Dépôt de Garantie:\s*(.+?)(?:\n|$)')
        self.data['cee'] = self._extract_pattern(text, r'CEE:\s*(.+?)(?:\n|$)')
        self.data['reference_vente'] = self._extract_pattern(text, r'Référence Vente:\s*(.+?)(?:\n|$)')

        print(f"✅ Fiche extraite - Segment: {self.data.get('segment')}, Score: {self.data.get('score')}")
        return self.data

    def extract_rgpd(self, pdf_path):
        """Extrait les données du document RGPD"""
        print(f"📄 Extraction du RGPD: {pdf_path}")

        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        # Extraction des données (mise à jour si différentes)
        rgpd_data = {}

        # Séparer le texte en sections pour éviter les confusions
        if 'INFORMATIONS GÉRANT' in text:
            section_societe = text.split('INFORMATIONS GÉRANT')[0]
            section_gerant = text.split('INFORMATIONS GÉRANT')[1]
        else:
            section_societe = text
            section_gerant = text

        # Informations Société (AVANT la section gérant)
        nom_societe = self._extract_pattern(section_societe, r'Nom\s*:\s*(.+?)(?:\n|$)')
        if nom_societe:
            rgpd_data['raison_sociale'] = nom_societe
            print(f"   → Nom société extrait: {nom_societe}")

        siren = self._extract_pattern(section_societe, r'SIREN\s*:\s*(\d{9})')
        if siren:
            rgpd_data['siren'] = siren
            print(f"   → SIREN extrait: {siren}")

        # BUG FIX 3: Extraire segment ET type de calendrier si présent (ex: "C5-BASE")
        segment_brut = self._extract_pattern(section_societe, r'Segment\s*:\s*([A-Z0-9-]+)')
        if segment_brut:
            # Si le format est "C5-BASE", séparer segment et calendrier
            if '-' in segment_brut:
                parts = segment_brut.split('-')
                rgpd_data['segment'] = parts[0].upper()  # "C5"
                rgpd_data['type_calendrier'] = parts[1].upper()  # "BASE"
                print(f"   → Segment extrait: {rgpd_data['segment']}, Type calendrier: {rgpd_data['type_calendrier']}")
            else:
                rgpd_data['segment'] = segment_brut.upper()
                print(f"   → Segment extrait: {segment_brut}")

        # PDL/PRM (pour pré-remplir la section 5 Sites)
        # BUG FIX 6: Gérer DEUX formats de PRM/PDL
        # Format 1 (multi-sites numéroté): "PDL/PCE (2 sites) : 1. 22122431219700  2. 22110564248878"
        # Format 2 (site unique): "PDL/PCE : 50051600701330"

        prm_list = []

        # FORMAT 1: Chercher d'abord les PRMs numérotés (multi-sites)
        # Pattern: "1. 22122431219700" ou "1. 22122431219700" avec espaces
        multi_match = re.findall(r'\d+\.\s*(\d{14})', text)
        if multi_match:
            prm_list = multi_match
            print(f"   → Format multi-sites détecté: {len(prm_list)} PRM(s) numéroté(s)")

        # FORMAT 2: Si aucun PRM numéroté, chercher format simple sur une ligne
        if not prm_list:
            # Pattern: "PDL/PCE : 50051600701330" ou "PRM : 50051600701330"
            simple_match = re.search(r'(?:PDL|PCE|PRM)[/\w\s]*:\s*(\d{14})', text, re.IGNORECASE)
            if simple_match:
                prm_list = [simple_match.group(1)]
                print(f"   → Format simple détecté: 1 PRM sur une ligne")

        # FALLBACK: Si toujours rien, chercher tous les nombres de 14 chiffres (sauf SIRET)
        if not prm_list:
            all_14 = re.findall(r'\b(\d{14})\b', text)
            if all_14:
                # Filtrer: exclure le SIRET si présent
                siren = rgpd_data.get('siren', '')
                if siren:
                    prm_list = [n for n in all_14 if not n.startswith(siren)]
                else:
                    prm_list = all_14
                print(f"   → Fallback générique: {len(prm_list)} nombre(s) de 14 chiffres trouvé(s)")

        # Stocker les résultats
        if prm_list:
            rgpd_data['prm_list'] = prm_list  # Liste complète
            rgpd_data['pdl'] = prm_list[0]    # Premier PRM (compatibilité)
            print(f"   → {len(prm_list)} PRM extrait(s): {', '.join(prm_list[:3])}{'...' if len(prm_list) > 3 else ''}")

            # Stocker aussi individuellement pour faciliter le pré-remplissage
            for idx, prm in enumerate(prm_list, 1):
                rgpd_data[f'prm_{idx}'] = prm
        else:
            print(f"   ⚠️ Aucun PRM/PDL trouvé dans le document")

        # Puissance souscrite (pour pré-remplir la section 5 Sites)
        # BUG FIX 3: Extraire la valeur numérique uniquement (sans "kVA")
        puissance_brut = self._extract_pattern(text, r'Puissance\s*:\s*([\d\s]+)\s*kVA')
        if puissance_brut:
            # Nettoyer: retirer espaces, garder seulement le nombre
            puissance_clean = puissance_brut.strip().replace(' ', '')
            rgpd_data['puissance'] = puissance_clean
            print(f"   → Puissance extraite: {puissance_clean} kVA")

        # BUG FIX 3: Adresse du site (avec plusieurs patterns pour C5/C4)
        # Format 1: "Adresse consommation : 4 RUE DE LATTRE..."
        adresse_consommation = self._extract_pattern(text, r'Adresse consommation\s*:\s*(.+?)(?:\n|$)')
        if not adresse_consommation:
            # Format 2: "Adresse du site : ..."
            adresse_consommation = self._extract_pattern(text, r'Adresse du site\s*:\s*(.+?)(?:\n|$)')
        if not adresse_consommation:
            # Format 3: juste "Adresse : ..." (dans la section société, pas gérant)
            adresse_consommation = self._extract_pattern(section_societe, r'Adresse\s*:\s*(.+?)(?:\n|$)')

        if adresse_consommation:
            rgpd_data['adresse_consommation'] = adresse_consommation.strip()
            rgpd_data['adresse_site'] = adresse_consommation.strip()  # Alias pour C5
            print(f"   → Adresse site extraite: {adresse_consommation[:60]}...")

        # Informations Gérant (section gérant uniquement)
        rgpd_data['civilite'] = self._extract_pattern(section_gerant, r'Civilité\s*:\s*(.+?)(?:\n|$)')
        nom_gerant = self._extract_pattern(section_gerant, r'Nom\s*:\s*(.+?)(?:\n|$)')
        if nom_gerant:
            rgpd_data['nom_gerant'] = nom_gerant
            print(f"   → Nom gérant extrait: {nom_gerant}")

        # Email (pour pré-remplir la section 2 Contact signataire)
        # BUG FIX 1: Chercher UNIQUEMENT dans la section GÉRANT (pas dans le header ECOGIES)
        email_match = re.search(r'INFORMATIONS GÉRANT.*?Email\s*:\s*(\S+@\S+)', text, re.DOTALL | re.IGNORECASE)
        if email_match:
            rgpd_data['email'] = email_match.group(1).strip()
            print(f"   → Email extrait (section GÉRANT): {rgpd_data['email']}")
        else:
            # Fallback: chercher n'importe où (pour les anciens formats)
            email = self._extract_pattern(text, r'Email\s*:\s*(.+?)(?:\n|$)')
            if email and '@' in email:
                rgpd_data['email'] = email.strip()
                print(f"   → Email extrait (fallback): {email}")

        # Téléphone (pour pré-remplir la section 2 Contact signataire)
        # BUG FIX 1: Chercher UNIQUEMENT dans la section GÉRANT
        tel_match = re.search(r'INFORMATIONS GÉRANT.*?Téléphone\s*:\s*([\d\s.]+)', text, re.DOTALL | re.IGNORECASE)
        if tel_match:
            rgpd_data['telephone'] = tel_match.group(1).strip().replace(' ', '').replace('.', '')
            print(f"   → Téléphone extrait (section GÉRANT): {rgpd_data['telephone']}")
        else:
            # Fallback
            telephone = self._extract_pattern(text, r'Téléphone\s*:\s*(.+?)(?:\n|$)')
            if telephone:
                rgpd_data['telephone'] = telephone.strip()
                print(f"   → Téléphone extrait (fallback): {telephone}")
        rgpd_data['date_validation_rgpd'] = self._extract_pattern(text, r'Date de validation.*?:\s*(.+?)(?:\n|$)')

        # Mise à jour des données principales
        self.data.update(rgpd_data)

        print(f"✅ RGPD extrait - Société: {rgpd_data.get('raison_sociale')}, SIREN: {rgpd_data.get('siren')}, PDL: {rgpd_data.get('pdl')}, Puissance: {rgpd_data.get('puissance')}, Email: {rgpd_data.get('email')}, Tél: {rgpd_data.get('telephone')}")
        return self.data

    def extract_siren(self, pdf_path):
        """Extrait les données de l'avis de situation SIREN ou extrait Pappers"""
        print(f"📄 Extraction de l'avis SIREN: {pdf_path}")

        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        # Extraction des données juridiques
        siren_data = {}

        # 1. SIRET (14 chiffres) - Extraction stricte uniquement
        # Tentative 1: SIRET direct (14 chiffres)
        siret_direct = self._extract_pattern(text, r'SIRET\s+(?:du siège)?\s*(\d{14})')
        if siret_direct:
            siren_data['siret_complet'] = siret_direct
            print(f"   → SIRET extrait (direct): {siret_direct}")
        else:
            # Tentative 2: Chercher n'importe quel numéro à 14 chiffres exact dans le texte
            siret_pattern = re.search(r'\b(\d{14})\b', text)
            if siret_pattern:
                siren_data['siret_complet'] = siret_pattern.group(1)
                print(f"   → SIRET trouvé (pattern général): {siren_data['siret_complet']}")
            else:
                # NE PAS inventer de SIRET - laisser vide si non trouvé
                print(f"   ⚠️ SIRET introuvable - champ laissé vide (mieux vaut vide qu'une valeur fausse)")

        # 2. Forme juridique - Nettoyer pour garder seulement le sigle (SAS, SARL, etc.)
        siren_data['categorie_juridique'] = self._extract_pattern(text, r'Catégorie juridique\s+(\d+\s*-\s*.+?)(?:\n|$)')
        forme_brute = self._extract_pattern(text, r'Catégorie juridique\s+\d+\s*-\s*(.+?)(?:\n|$)')
        if not forme_brute:
            # Format Pappers
            forme_brute = self._extract_pattern(text, r'Forme juridique\s+(.+?)(?:\n|$)')

        # Nettoyer la forme juridique pour garder uniquement le sigle (avant la virgule)
        if forme_brute:
            forme_nettoyee = forme_brute.split(',')[0].strip()
            siren_data['forme_juridique'] = forme_nettoyee
            if forme_brute != forme_nettoyee:
                print(f"   → Forme juridique nettoyée: '{forme_brute}' → '{forme_nettoyee}'")
            else:
                print(f"   → Forme juridique: {forme_nettoyee}")

        siren_data['code_ape'] = self._extract_pattern(text, r'Activité Principale.*?\(APE\)\s+(.+?)(?:\n|$)')

        # 3. Capital social - Extraction et nettoyage en nombre pur
        capital_brut = self._extract_pattern(text, r'Capital\s*social\s+([\d\s,\.]+)\s*Euros')
        if capital_brut:
            # Nettoyer: retirer espaces et remplacer virgule par point
            capital_nettoye = capital_brut.replace(' ', '').replace(',', '.')
            siren_data['capital_social'] = capital_nettoye
            print(f"   → Capital social: {capital_brut} € → {capital_nettoye}")

        # Code NAF (format Pappers : "96.01B")
        # BUG FIX 4: Nettoyer le point dans le NAF (ex: "47.11A" → "4711A")
        if not siren_data.get('code_ape'):
            naf_match = self._extract_pattern(text, r'code NAF\)?\s*(\d{2}\.\d{2}[A-Z])')
            if naf_match:
                # Retirer le point: "47.11A" → "4711A"
                naf_clean = naf_match.replace('.', '')
                siren_data['code_naf'] = naf_clean
                print(f"   → Code NAF extrait (Pappers): {naf_match} → {naf_clean}")

        # 4. Adresse siège (format Pappers)
        adresse_pappers = self._extract_pattern(text, r'Adresse de l.établissement\s+(.+?)(?:\n|Activité)', flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if adresse_pappers:
            # Nettoyer l'adresse (retirer les sauts de ligne superflus)
            siren_data['adresse_siege_pappers'] = ' '.join(adresse_pappers.split())
            print(f"   → Adresse siège Pappers: {siren_data['adresse_siege_pappers']}")

        # 5. Adresse établissement (pour site de consommation)
        # Chercher une adresse d'établissement différente du siège
        adresse_etablissement = self._extract_pattern(text, r'Adresse de l.établissement principal\s+(.+?)(?:\n|$)', flags=re.IGNORECASE | re.MULTILINE)
        if not adresse_etablissement:
            # Essayer un autre pattern
            adresse_etablissement = self._extract_pattern(text, r'Établissement principal.*?Adresse\s+(.+?)(?:\n|Activité)', flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if adresse_etablissement:
            siren_data['adresse_etablissement'] = ' '.join(adresse_etablissement.split())
            print(f"   → Adresse établissement: {siren_data['adresse_etablissement']}")

        # 6. Fonction du signataire (format Pappers : "Président", "Gérant")
        fonction_signataire = self._extract_pattern(text, r'(Président|Gérant|Directeur\s+[Gg]énéral|Co-gérant|Directrice\s+[Gg]énérale)')
        if fonction_signataire:
            siren_data['fonction_signataire'] = fonction_signataire
            print(f"   → Fonction signataire: {fonction_signataire}")

        # 7. Nom gérant - Nettoyer pour garder seulement NOM + Premier Prénom
        nom_complet_brut = self._extract_pattern(text, r'Nom,?\s*prénoms?\s+(.+?)(?:\n|$)')
        if nom_complet_brut:
            # Garder seulement les 2 premiers mots (NOM + Premier prénom)
            mots = nom_complet_brut.split()
            if len(mots) >= 2:
                nom_nettoye = f"{mots[0]} {mots[1]}"
                siren_data['nom_gerant_pappers'] = nom_nettoye
                if len(mots) > 2:
                    print(f"   → Nom gérant nettoyé: '{nom_complet_brut}' → '{nom_nettoye}'")
                else:
                    print(f"   → Nom gérant Pappers: {nom_nettoye}")
            else:
                siren_data['nom_gerant_pappers'] = nom_complet_brut
                print(f"   → Nom gérant Pappers: {nom_complet_brut}")

        # Date d'immatriculation
        date_immat = self._extract_pattern(text, r'Date d.immatriculation\s+(\d{2}/\d{2}/\d{4})')
        if date_immat:
            siren_data['date_immatriculation'] = date_immat
            print(f"   → Date immatriculation: {date_immat}")

        # 8. Ville RCS (format Pappers) - Déjà bien implémenté
        ville_rcs_match = self._extract_pattern(text, r'R\.C\.S\.\s+(\w+)')
        if ville_rcs_match:
            siren_data['ville_rcs'] = ville_rcs_match
            print(f"   → Ville RCS: {ville_rcs_match}")
        else:
            # Extraction de la ville du RCS depuis l'adresse (code postal)
            adresse = self.data.get('adresse_siege', '')
            code_postal = self._extract_pattern(adresse, r'(\d{5})')
            if code_postal:
                dept = code_postal[:2]
                # Mapping des départements aux tribunaux de commerce principaux
                ville_rcs_map = {
                    '38': 'Grenoble',
                    '69': 'Lyon',
                    '75': 'Paris',
                    '13': 'Marseille',
                    '44': 'Nantes',
                    '33': 'Bordeaux',
                    '34': 'Montpellier',
                    '59': 'Lille',
                    '31': 'Toulouse',
                }
                siren_data['ville_rcs'] = ville_rcs_map.get(dept, f'Département {dept}')
                if siren_data['ville_rcs']:
                    print(f"   → Ville RCS (déduite du code postal): {siren_data['ville_rcs']}")

        # Mise à jour des données principales
        self.data.update(siren_data)

        print(f"✅ SIREN extrait - Forme: {siren_data.get('forme_juridique')}, Capital: {siren_data.get('capital_social')} €")
        return self.data

    def extract_inpi(self, pdf_path):
        """Extrait les données de la fiche INPI (Registre National des Entreprises)
        Supporte aussi les extraits Pappers du RNE"""
        print(f"📄 Extraction de la fiche INPI: {pdf_path}")

        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        inpi_data = {}

        # 1. Dénomination → raison_sociale
        # Format INPI: "Dénomination : A.L.V. DUMAS"
        # Format Pappers: "Dénomination ou raison sociale A.L.V. DUMAS"
        denomination = self._extract_pattern(text, r'Dénomination\s*(?:ou raison sociale)?\s*:?\s*(.+?)(?:\n|$)')
        if denomination:
            inpi_data['raison_sociale'] = denomination
            print(f"   → Dénomination extraite: {denomination}")

        # 2. SIREN (9 chiffres, retirer espaces) → siren
        # Format INPI: "SIREN (siège) : 443 762 638"
        # Format Pappers: "numéro 443 762 638 R.C.S."
        siren_brut = self._extract_pattern(text, r'SIREN\s*(?:\(siège\))?\s*:?\s*([\d\s]+)')
        if not siren_brut:
            # Essayer format Pappers: "numéro 443 762 638 R.C.S."
            siren_brut = self._extract_pattern(text, r'numéro\s+([\d\s]+)\s+R\.C\.S\.')
        if siren_brut:
            # Retirer les espaces: "443 762 638" → "443762638"
            siren_nettoye = siren_brut.replace(' ', '')
            if len(siren_nettoye) == 9:
                inpi_data['siren'] = siren_nettoye
                print(f"   → SIREN extrait: {siren_brut} → {siren_nettoye}")
            else:
                print(f"   ⚠️ SIREN invalide (longueur != 9): {siren_nettoye}")

        # 3. Code APE/NAF (juste le code, pas le libellé) → code_naf
        # Ex: "2562B - Mécanique industrielle" → extraire "2562B"
        # Format INPI: "Code APE : 2562B - Mécanique industrielle"
        # Format alternatif: "APE/NAF : 5610A" ou juste "5610A" après "Activité"
        code_ape = self._extract_pattern(text, r'Code APE\s*:?\s*(\d{4}[A-Z])')
        if not code_ape:
            # Essayer format "APE/NAF" ou "APE-NAF"
            code_ape = self._extract_pattern(text, r'(?:APE|NAF)[/\s-]+:?\s*(\d{4}[A-Z])')
        if not code_ape:
            # Essayer recherche plus large: 4 chiffres + 1 lettre majuscule isolé
            code_ape = self._extract_pattern(text, r'\b(\d{4}[A-Z])\b')

        if code_ape:
            # Nettoyer le code NAF : retirer les points (ex: "47.11A" → "4711A")
            code_ape_clean = code_ape.replace('.', '').replace(' ', '')
            inpi_data['code_naf'] = code_ape_clean
            print(f"   → Code APE/NAF extrait: {code_ape} → {code_ape_clean}")
        else:
            # Pas de code APE dans ce document (normal pour certains extraits Pappers)
            print(f"   ⚠️ Code APE/NAF non trouvé dans le document")

        # 4. Forme juridique → forme_juridique
        # Format Pappers: "SAS, société par actions simplifiée" → extraire "SAS"
        forme_jur = self._extract_pattern(text, r'Forme juridique\s*:?\s*(.+?)(?:\n|$)')
        if forme_jur:
            # Si "Entrepreneur individuel" → capital_social = 0
            if 'Entrepreneur individuel' in forme_jur or 'Entreprise individuelle' in forme_jur:
                inpi_data['forme_juridique'] = 'Entrepreneur individuel'
                inpi_data['capital_social'] = '0'
                print(f"   → Forme juridique: Entrepreneur individuel (capital = 0)")
            else:
                # Nettoyer (retirer les virgules et texte après)
                forme_nettoyee = forme_jur.split(',')[0].strip()
                inpi_data['forme_juridique'] = forme_nettoyee
                print(f"   → Forme juridique: {forme_nettoyee}")

        # 5. Capital social (montant uniquement) → capital_social
        # Ex: "3750 EUR" → "3750"
        if not inpi_data.get('capital_social'):  # Ne pas écraser si EI
            capital_brut = self._extract_pattern(text, r'Capital social\s*:?\s*([\d\s,\.]+)\s*EUR')
            if capital_brut:
                # Nettoyer: retirer espaces et virgules
                capital_nettoye = capital_brut.replace(' ', '').replace(',', '.')
                inpi_data['capital_social'] = capital_nettoye
                print(f"   → Capital social: {capital_brut} EUR → {capital_nettoye}")

        # 6. Adresse du siège → adresse_siege_inpi
        adresse_siege = self._extract_pattern(text, r'Adresse du siège\s*:?\s*(.+?)(?:\n\n|$)', flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if adresse_siege:
            # Nettoyer (remplacer \n par espaces)
            adresse_nettoyee = ' '.join(adresse_siege.split())
            inpi_data['adresse_siege_inpi'] = adresse_nettoyee
            print(f"   → Adresse siège: {adresse_nettoyee[:60]}...")

        # 6.5. Adresse de l'établissement principal → adresse_etablissement_principal
        # Utile pour les sites C5 : l'adresse du site peut différer de l'adresse du siège
        # Ex: "Type d'établissement : Principal ... Adresse : 4 AVENUE DU MAL DE LATTRE DE TASSIGNY 77370 NANGIS"
        adresse_principal_match = re.search(
            r'Type\s*d.établissement\s*:\s*Principal.*?Adresse\s*:?\s*(.+?)(?:Données|Type|État|$)',
            text, re.DOTALL | re.IGNORECASE
        )
        if adresse_principal_match:
            adresse_principal = adresse_principal_match.group(1).strip()
            # Nettoyer (remplacer \n par espaces et retirer "FRANCE" à la fin)
            adresse_principal_clean = ' '.join(adresse_principal.split())
            adresse_principal_clean = re.sub(r'\s+FRANCE\s*$', '', adresse_principal_clean, flags=re.IGNORECASE).strip()
            inpi_data['adresse_etablissement_principal'] = adresse_principal_clean
            inpi_data['adresse_site'] = adresse_principal_clean  # Alias pour C5
            print(f"   → Adresse établissement principal: {adresse_principal_clean[:60]}...")

        # 7. Nom, Prénom(s) (page 2, section "Gestion et Direction") → nom_signataire_inpi
        # BUG FIX 5: Détecter si c'est une personne morale (société) ou personne physique
        # Format personne morale: "Dénomination : MJ INVEST"
        # Format personne physique: "Nom, Prénom(s) : REBATEL KEVIN"

        # D'abord chercher une Dénomination (personne morale)
        denomination_dirigeant = self._extract_pattern(text, r'(?:Gestion et Direction|DIRIGEANT).*?Dénomination\s*:?\s*(.+?)(?:\n|$)', flags=re.DOTALL | re.IGNORECASE)

        if denomination_dirigeant:
            # C'est une personne morale (société) comme dirigeant
            inpi_data['personne_morale_dirigeant'] = True
            inpi_data['denomination_dirigeant'] = denomination_dirigeant.strip()
            print(f"   → Personne morale détectée comme dirigeant: '{denomination_dirigeant}'")
            print(f"   → Le nom du signataire sera extrait du RGPD (pas de l'INPI)")
        else:
            # Personne physique : chercher "Nom, Prénom(s)"
            nom_complet = self._extract_pattern(text, r'Nom,\s*(?:Prénom\(s\)|prénoms)\s*:?\s*(.+?)(?:\n|$)')
            if nom_complet:
                inpi_data['personne_morale_dirigeant'] = False
                # Séparer par virgules et espaces
                mots = nom_complet.replace(',', ' ').split()
                # Garder les 2 premiers : NOM (tout en majuscules) + Prénom (capitalize)
                if len(mots) >= 2:
                    nom = mots[0].upper()  # NOM en majuscules
                    prenom = mots[1].capitalize()  # Prénom en capitalize
                    nom_nettoye = f"{nom} {prenom}"
                    inpi_data['nom_signataire_inpi'] = nom_nettoye
                    if len(mots) > 2:
                        print(f"   → Nom signataire nettoyé: '{nom_complet}' → '{nom_nettoye}'")
                    else:
                        print(f"   → Nom signataire: {nom_nettoye}")
                else:
                    inpi_data['nom_signataire_inpi'] = nom_complet.strip()
                    print(f"   → Nom signataire: {nom_complet}")

        # 8. Qualité (fonction signataire) → fonction_signataire
        # Ex: "Président de SAS" → extraire "Président"
        # Format INPI: "Qualité : Président de SAS"
        # Format Pappers: Juste "Président" sur une ligne
        qualite = self._extract_pattern(text, r'Qualité\s*:?\s*(.+?)(?:\n|$)')
        if not qualite:
            # Format Pappers: "Président" tout seul avant "Nom, prénoms"
            qualite = self._extract_pattern(text, r'(Président|Gérant|Directeur\s+général|Directrice\s+générale)\s*\n\s*Nom,\s*prénoms')
        if qualite:
            # Extraire juste le titre principal (Président, Gérant, etc.)
            if 'Président' in qualite:
                fonction = 'Président'
            elif 'Gérant' in qualite:
                fonction = 'Gérant'
            elif 'Directeur général' in qualite or 'Directrice générale' in qualite:
                fonction = 'Directeur général'
            else:
                fonction = qualite.strip()
            inpi_data['fonction_signataire'] = fonction
            print(f"   → Fonction signataire: {fonction}")

        # 9. SIRET (page 2, établissement actif) → siret_complet
        # BUG FIX 3: Chercher d'abord le SIRET de l'établissement Principal (pas Siège)
        # Ex: "Type d'établissement : Principal ... Siret : 91932688400021"
        siret_principal_match = re.search(r'Type d.établissement\s*:\s*Principal.*?Siret\s*:?\s*([\d\s]+)', text, re.DOTALL | re.IGNORECASE)
        if siret_principal_match:
            siret_brut = siret_principal_match.group(1)
            siret_nettoye = siret_brut.replace(' ', '').strip()
            if len(siret_nettoye) == 14:
                inpi_data['siret_complet'] = siret_nettoye
                inpi_data['siret_principal'] = siret_nettoye
                print(f"   → SIRET Principal extrait: {siret_brut} → {siret_nettoye}")
            else:
                print(f"   ⚠️ SIRET Principal invalide (longueur != 14): {siret_nettoye}")
        else:
            # Fallback: chercher n'importe quel SIRET
            siret_brut = self._extract_pattern(text, r'Siret\s*:?\s*([\d\s]+)')
            if siret_brut:
                siret_nettoye = siret_brut.replace(' ', '').strip()
                if len(siret_nettoye) == 14:
                    inpi_data['siret_complet'] = siret_nettoye
                    print(f"   → SIRET complet extrait (fallback): {siret_brut} → {siret_nettoye}")
                else:
                    print(f"   ⚠️ SIRET complet invalide (longueur != 14): {siret_nettoye}")

        # 10. Ville du RCS → ville_rcs
        # Ex: "Tribunal de Commerce de Grenoble en date du..." → "Grenoble"
        # Ou: "Registre du Commerce et des Sociétés de Grenoble" → "Grenoble"
        # Ou: "R.C.S. NANGIS" → "NANGIS"
        ville_rcs_match = re.search(r'(?:Tribunal de Commerce|Registre du Commerce et des Sociétés) de\s+([A-Z][\wÀ-ÿ-]+)', text, re.IGNORECASE)
        if not ville_rcs_match:
            # Format court: "R.C.S. NANGIS" ou "RCS NANGIS"
            ville_rcs_match = re.search(r'R\.?C\.?S\.?\s+([A-Z][\wÀ-ÿ-]+)', text, re.IGNORECASE)
        if not ville_rcs_match:
            # Format "immatriculée au RCS de NANGIS"
            ville_rcs_match = re.search(r'RCS\s+de\s+([A-Z][\wÀ-ÿ-]+)', text, re.IGNORECASE)

        if ville_rcs_match:
            inpi_data['ville_rcs'] = ville_rcs_match.group(1).strip().title()  # Capitalize pour uniformité
            print(f"   → Ville RCS extraite: {inpi_data['ville_rcs']}")
        else:
            # BUG FIX C5: Fallback - Extraire la ville depuis l'adresse de l'établissement principal
            # Format attendu: "4 AVENUE DU MAL DE LATTRE DE TASSIGNY 77370 NANGIS"
            # On cherche : code postal (5 chiffres) + virgule optionnelle + VILLE
            adresse_principal_match = re.search(
                r'Type\s*d.établissement\s*:\s*Principal.*?Adresse\s*:?\s*(.+?)(?:Données|Type|État|$)',
                text, re.DOTALL | re.IGNORECASE
            )
            if adresse_principal_match:
                adresse_text = adresse_principal_match.group(1).strip()
                # Chercher "77370 , NANGIS" ou "77370 NANGIS" dans l'adresse
                ville_from_addr = re.search(r'\b\d{5}\s*,?\s*([A-ZÀ-Ü][\w\s-]+?)(?:\s*FRANCE|Données|Type|État|\n|$)', adresse_text, re.IGNORECASE)
                if ville_from_addr:
                    ville_extraite = ville_from_addr.group(1).strip().title()
                    # Nettoyer : retirer les mots en trop (FRANCE, etc.)
                    ville_extraite = re.sub(r'\s+(FRANCE|FR)$', '', ville_extraite, flags=re.IGNORECASE).strip()
                    inpi_data['ville_rcs'] = ville_extraite
                    print(f"   → Ville RCS extraite depuis adresse établissement: {ville_extraite}")
                else:
                    print(f"   ⚠️ Ville RCS non trouvée (adresse principal: {adresse_text[:60]}...)")
            else:
                print(f"   ⚠️ Ville RCS non trouvée dans le document")

        # Mise à jour des données principales
        self.data.update(inpi_data)

        print(f"✅ INPI extrait - Forme: {inpi_data.get('forme_juridique')}, Capital: {inpi_data.get('capital_social')} €")
        return self.data

    def _extract_pattern(self, text, pattern, flags=None):
        """Extrait un pattern regex du texte"""
        if flags is None:
            flags = re.IGNORECASE | re.MULTILINE
        match = re.search(pattern, text, flags)
        if match:
            return match.group(1).strip()
        return ""

    def get_all_data(self):
        """Retourne toutes les données extraites"""
        return self.data


def _detect_pdf_type(pdf_path):
    """Détecte le type de PDF en analysant son contenu (pas le nom du fichier)

    Returns:
        str: 'fiche', 'rgpd', 'inpi', 'pappers', ou 'unknown'
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Lire les 2 premières pages pour détecter le type
            text = ""
            for i, page in enumerate(pdf.pages):
                if i >= 2:  # Limiter à 2 pages pour la détection
                    break
                text += page.extract_text() + "\n"

            text_lower = text.lower()

            # Détection INPI/RNE (fichier officiel du Registre National des Entreprises)
            if any(keyword in text_lower for keyword in [
                'attestation d\'immatriculation au registre national',
                'registre national des entreprises',
                'data.inpi.fr',
                'extrait du registre national',
                'extrait rne'
            ]):
                return 'inpi'

            # Détection Pappers (extrait Pappers du RNE)
            if 'extrait pappers' in text_lower or 'pappers.fr' in text_lower:
                return 'inpi'  # Les extraits Pappers sont traités comme des INPI

            # Détection RGPD
            if any(keyword in text_lower for keyword in [
                'document rgpd',
                'consentement données énergétiques',
                'informations gérant',
                'validation rgpd'
            ]):
                return 'rgpd'

            # Détection Fiche Contact
            if any(keyword in text_lower for keyword in [
                'fiche contact',
                'commercial ohm',
                'courtier final',
                'pdl principal'
            ]):
                return 'fiche'

            # Détection Avis SIREN (ancien format)
            if any(keyword in text_lower for keyword in [
                'avis de situation',
                'numéro siren',
                'catégorie juridique'
            ]):
                return 'pappers'

            return 'unknown'
    except Exception as e:
        print(f"   ⚠️ Erreur lors de la détection du type de {pdf_path.name}: {e}")
        return 'unknown'


def extract_all_pdfs(folder_path):
    """
    Extrait les données de tous les PDFs dans un dossier avec fusion intelligente

    Args:
        folder_path: Chemin vers le dossier contenant les PDFs

    Returns:
        dict: Toutes les données extraites et fusionnées selon les priorités
    """
    extractor = PDFExtractor()
    folder = Path(folder_path)

    # Recherche des fichiers PAR CONTENU (pas par nom de fichier)
    fiche_file = None
    rgpd_file = None
    siren_file = None
    inpi_file = None

    print("\n🔍 Détection automatique du type de PDFs...")
    for pdf_file in folder.glob("*.pdf"):
        pdf_type = _detect_pdf_type(pdf_file)
        print(f"   📄 {pdf_file.name[:50]:50s} → Type détecté: {pdf_type}")

        if pdf_type == 'fiche':
            fiche_file = pdf_file
        elif pdf_type == 'rgpd':
            rgpd_file = pdf_file
        elif pdf_type == 'inpi':
            inpi_file = pdf_file
        elif pdf_type == 'pappers':
            # Pappers est un fallback si pas d'INPI
            if not inpi_file:
                siren_file = pdf_file

    # Extraction
    if fiche_file:
        extractor.extract_fiche(fiche_file)
    else:
        print("⚠️ Fiche Contact non trouvée")

    if rgpd_file:
        extractor.extract_rgpd(rgpd_file)
    else:
        print("⚠️ Document RGPD non trouvé")

    # Priorité INPI sur Pappers (INPI = source officielle la plus récente)
    if inpi_file:
        extractor.extract_inpi(inpi_file)
    elif siren_file:
        extractor.extract_siren(siren_file)
    else:
        print("⚠️ Avis SIREN/INPI non trouvé")

    # Fusion intelligente avec priorités
    data = extractor.get_all_data()

    print("\n🔧 FUSION INTELLIGENTE DES DONNÉES:")

    # Priorité INPI/Pappers pour SIRET (plus fiable que RGPD)
    if data.get('siret_complet'):
        source = 'INPI' if inpi_file else 'Pappers'
        print(f"   ✓ SIRET ({source} prioritaire): {data['siret_complet']}")

    # Priorité Pappers pour Ville RCS (INPI n'a pas cette info)
    if data.get('ville_rcs'):
        print(f"   ✓ Ville RCS (Pappers): {data['ville_rcs']}")

    # Priorité INPI/Pappers pour Capital social
    if data.get('capital_social'):
        source = 'INPI' if inpi_file else 'Pappers'
        print(f"   ✓ Capital social ({source} prioritaire): {data['capital_social']} €")

    # Priorité INPI/Pappers pour Forme juridique
    if data.get('forme_juridique'):
        source = 'INPI' if inpi_file else 'Pappers'
        print(f"   ✓ Forme juridique ({source} prioritaire): {data['forme_juridique']}")

    # Priorité INPI/Pappers pour Fonction signataire
    if data.get('fonction_signataire'):
        source = 'INPI' if inpi_file else 'Pappers'
        print(f"   ✓ Fonction signataire ({source} prioritaire): {data['fonction_signataire']}")

    # Priorité INPI/Pappers pour Code NAF
    if data.get('code_naf'):
        source = 'INPI' if inpi_file else 'Pappers'
        print(f"   ✓ Code NAF ({source}): {data['code_naf']}")

    # Priorité INPI pour Adresse siège, sinon Pappers
    if data.get('adresse_siege_inpi'):
        data['adresse_siege'] = data['adresse_siege_inpi']
        print(f"   ✓ Adresse siège (INPI prioritaire): {data['adresse_siege'][:50]}...")
    elif data.get('adresse_siege_pappers'):
        data['adresse_siege'] = data['adresse_siege_pappers']
        print(f"   ✓ Adresse siège (Pappers): {data['adresse_siege'][:50]}...")

    # Priorité Pappers pour Adresse établissement (site de consommation)
    if data.get('adresse_etablissement'):
        print(f"   ✓ Adresse établissement (Pappers): {data['adresse_etablissement'][:50]}...")

    # Priorité INPI pour Nom signataire, sinon Pappers
    if data.get('nom_signataire_inpi'):
        data['nom_signataire'] = data['nom_signataire_inpi']
        print(f"   ✓ Nom signataire (INPI prioritaire): {data['nom_signataire']}")
    elif data.get('nom_gerant_pappers'):
        data['nom_gerant'] = data['nom_gerant_pappers']
        print(f"   ✓ Nom gérant (Pappers): {data['nom_gerant']}")

    # Nettoyer le segment (C5-BASE → C5, C4-H4 → C4, etc.)
    if data.get('segment'):
        segment_brut = data['segment']
        segment_clean = segment_brut.split('-')[0].upper()
        if segment_clean in ['C2', 'C3', 'C4', 'C5']:
            data['segment'] = segment_clean
            if segment_brut != segment_clean:
                print(f"   ✓ Segment nettoyé: '{segment_brut}' → '{segment_clean}'")

    # MODIFICATION 3: Mapper C3 → C2 (C3 est traité comme C2)
    if data.get('segment') == 'C3':
        data['segment'] = 'C2'
        print(f"   ✓ Segment C3 détecté → mappé vers C2")

    # Raison sociale : prendre la plus complète (plus de caractères)
    raison_sociale_candidates = []
    if data.get('raison_sociale'):
        raison_sociale_candidates.append(data['raison_sociale'])
    if data.get('nom_societe'):
        raison_sociale_candidates.append(data['nom_societe'])

    if raison_sociale_candidates:
        raison_sociale_finale = max(raison_sociale_candidates, key=len)
        data['raison_sociale'] = raison_sociale_finale
        print(f"   ✓ Raison sociale (la plus complète): {raison_sociale_finale}")

    print("   ✅ Fusion terminée\n")

    return data


if __name__ == "__main__":
    # Test
    test_folder = "/Users/strategyglobal/Desktop/cpv mint"
    data = extract_all_pdfs(test_folder)

    print("\n" + "="*60)
    print("DONNÉES EXTRAITES:")
    print("="*60)
    for key, value in sorted(data.items()):
        print(f"{key:30s}: {value}")
