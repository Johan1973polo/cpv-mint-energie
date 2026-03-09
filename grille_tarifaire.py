"""
Module de lecture et recherche dans les grilles tarifaires MINT
Support CSV (ancien format) et Excel (nouveau format 2026)
"""
import csv
from datetime import datetime
from pathlib import Path


class GrilleTarifaire:
    """Gestion des grilles tarifaires MINT C2, C4, C5"""

    def __init__(self, base_path="../", excel_parser=None):
        """
        Initialise avec soit des fichiers CSV, soit un parser Excel

        Args:
            base_path: Chemin vers les fichiers CSV (si pas d'Excel)
            excel_parser: Instance de ExcelGrilleParser (prioritaire si fourni)
        """
        self.base_path = Path(base_path)
        self.excel_parser = excel_parser
        self.grilles = {
            'C2': None,
            'C4': None,
            'C5': None
        }
        self.metadata = {}

        # Charger depuis Excel ou CSV
        if excel_parser:
            self._load_from_excel()
        else:
            self._load_grilles()

    def _load_from_excel(self):
        """Charge les grilles depuis un parser Excel"""
        if not self.excel_parser:
            return

        # Récupérer les grilles parsées
        self.grilles = self.excel_parser.grilles
        self.metadata = self.excel_parser.metadata

        for segment in ['C2', 'C4', 'C5']:
            if self.grilles[segment]:
                print(f"✅ Grille {segment} chargée depuis Excel: {len(self.grilles[segment])} lignes")

    def _load_grilles(self):
        """Charge les 3 grilles tarifaires depuis les CSV"""
        # Chercher les fichiers CSV
        grille_files = {
            'C2': None,
            'C4': None,
            'C5': None
        }

        for csv_file in self.base_path.glob("*.csv"):
            filename = csv_file.name
            # Identifier de manière précise par le nom exact
            if 'Grille_C5' in filename:
                grille_files['C5'] = csv_file
            elif 'Grille_C2' in filename:
                grille_files['C2'] = csv_file
            elif 'Grille_C4' in filename:
                grille_files['C4'] = csv_file

        # Charger chaque grille
        for segment, filepath in grille_files.items():
            if filepath:
                self.grilles[segment] = self._parse_csv(filepath, segment)
                print(f"✅ Grille {segment} chargée depuis CSV: {len(self.grilles[segment])} lignes")

    def _parse_csv(self, filepath, segment):
        """Parse un fichier CSV de grille tarifaire"""
        lignes = []

        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)

            # Trouver la ligne d'en-tête (contient "Date début")
            header_idx = None
            for i, row in enumerate(rows):
                if row and 'Date début' in row[0]:
                    header_idx = i
                    break

            if header_idx is None:
                return lignes

            # Parser les données
            for row in rows[header_idx + 1:]:
                if not row or not row[0]:  # Ligne vide
                    continue

                try:
                    ligne = {
                        'date_debut': row[0].strip() if len(row) > 0 else '',
                        'duree_mois': row[1].strip() if len(row) > 1 else '',
                        'date_fin': row[2].strip() if len(row) > 2 else '',
                    }

                    # Prix selon le segment
                    if segment == 'C2':
                        # Colonnes: Pointe, HPH, HCH, HPE, HCE, Prix Capacité 2025, 2026
                        ligne['prix_pte'] = row[3].replace(',', '.') if len(row) > 3 else '0'
                        ligne['prix_hph'] = row[4].replace(',', '.') if len(row) > 4 else '0'
                        ligne['prix_hch'] = row[5].replace(',', '.') if len(row) > 5 else '0'
                        ligne['prix_hpe'] = row[6].replace(',', '.') if len(row) > 6 else '0'
                        ligne['prix_hce'] = row[7].replace(',', '.') if len(row) > 7 else '0'
                        ligne['prix_capa_2025'] = row[8].replace(',', '.') if len(row) > 8 else '0'
                        ligne['prix_capa_2026'] = row[9].replace(',', '.') if len(row) > 9 else '0'

                    elif segment == 'C4':
                        # Colonnes: HPH, HCH, HPE, HCE, Prix Capacité 2025, 2026
                        ligne['prix_hph'] = row[3].replace(',', '.') if len(row) > 3 else '0'
                        ligne['prix_hch'] = row[4].replace(',', '.') if len(row) > 4 else '0'
                        ligne['prix_hpe'] = row[5].replace(',', '.') if len(row) > 5 else '0'
                        ligne['prix_hce'] = row[6].replace(',', '.') if len(row) > 6 else '0'
                        ligne['prix_capa_2025'] = row[7].replace(',', '.') if len(row) > 7 else '0'
                        ligne['prix_capa_2026'] = row[8].replace(',', '.') if len(row) > 8 else '0'

                    elif segment == 'C5':
                        # Colonnes: Base, HP, HC, Prix Capacité 2026
                        ligne['prix_base'] = row[3].replace(',', '.') if len(row) > 3 else '0'
                        ligne['prix_hp'] = row[4].replace(',', '.') if len(row) > 4 else '0'
                        ligne['prix_hc'] = row[5].replace(',', '.') if len(row) > 5 else '0'
                        ligne['prix_capa_2026'] = row[6].replace(',', '.') if len(row) > 6 else '0'

                    lignes.append(ligne)

                except Exception as e:
                    print(f"⚠️ Erreur parsing ligne: {e}")
                    continue

        return lignes

    def get_prix_p0(self, segment, date_debut, date_fin=None, duree_mois=None):
        """
        Récupère les prix P0 (sans marge) pour un segment et des dates/durée données

        Args:
            segment: 'C2', 'C4', ou 'C5'
            date_debut: Date de début (format DD/MM/YYYY ou datetime)
            date_fin: Date de fin (format DD/MM/YYYY ou datetime) - pour CSV
            duree_mois: Durée en mois (int) - pour Excel

        Returns:
            dict avec les prix P0 ou None si non trouvé
        """
        if segment not in self.grilles or not self.grilles[segment]:
            return None

        # Mode Excel : recherche par date_debut + duree_mois
        if duree_mois is not None:
            for ligne in self.grilles[segment]:
                if (ligne['date_debut'] == date_debut and
                    ligne.get('duree_mois') == duree_mois):
                    return ligne
            return None

        # Mode CSV : recherche par date_debut + date_fin
        if date_fin is not None:
            # Convertir les dates si nécessaire
            if isinstance(date_debut, str):
                date_debut = self._parse_date(date_debut)
            if isinstance(date_fin, str):
                date_fin = self._parse_date(date_fin)

            # NORMALISER LA DATE DE DÉBUT AU 1ER DU MOIS
            if date_debut:
                date_debut = date_debut.replace(day=1)

            # Chercher la ligne correspondante avec DEBUT ET FIN
            for ligne in self.grilles[segment]:
                ligne_date_debut = self._parse_date(ligne['date_debut'])
                ligne_date_fin = self._parse_date(ligne['date_fin'])

                if (ligne_date_debut and ligne_date_debut == date_debut and
                    ligne_date_fin and ligne_date_fin == date_fin):
                    return ligne

        return None

    def _parse_date(self, date_str):
        """Parse une date au format DD/MM/YYYY"""
        try:
            if '/' in date_str:
                return datetime.strptime(date_str, '%d/%m/%Y')
            elif '-' in date_str:
                return datetime.strptime(date_str, '%Y-%m-%d')
        except:
            return None

    def get_dates_disponibles(self, segment):
        """
        Récupère toutes les dates de début disponibles pour un segment

        Args:
            segment: 'C2', 'C4' ou 'C5'

        Returns:
            list: Liste des dates de début uniques (format DD/MM/YYYY)
        """
        if not self.grilles[segment]:
            return []

        dates = set()
        for ligne in self.grilles[segment]:
            dates.add(ligne['date_debut'])

        return sorted(list(dates))

    def get_durees_disponibles(self, segment, date_debut):
        """
        Récupère les durées disponibles pour un segment et une date de début

        Args:
            segment: 'C2', 'C4' ou 'C5'
            date_debut: Date de début (format DD/MM/YYYY)

        Returns:
            list: Liste des durées disponibles (en mois)
        """
        if not self.grilles[segment]:
            return []

        durees = []
        for ligne in self.grilles[segment]:
            if ligne['date_debut'] == date_debut:
                duree = ligne.get('duree_mois')
                if duree:
                    durees.append(duree)

        return sorted(durees) if durees else []

    def calculer_prix_avec_marge(self, prix_p0, marge_fournisseur=0, marge_courtier=0):
        """
        Calcule les prix finaux avec marges

        Args:
            prix_p0: dict des prix P0
            marge_fournisseur: Marge fournisseur (€/MWh, max 12.50) - optionnel
            marge_courtier: Marge courtier (€/MWh, min 6, max 25) - nouveau format

        Returns:
            dict avec prix finaux et commission courtier
        """
        # Nouveau format 2026 : marge courtier uniquement (6-25 €/MWh)
        if marge_courtier > 0:
            marge_courtier = max(6.0, min(25.0, float(marge_courtier)))
            marge_totale = marge_courtier
        else:
            # Ancien format : fournisseur + courtier
            marge_fournisseur = min(float(marge_fournisseur), 12.50)
            marge_courtier = min(float(marge_courtier), 12.50)
            marge_totale = marge_fournisseur + marge_courtier

            if marge_totale > 25.0:
                # Réduire proportionnellement
                ratio = 25.0 / marge_totale
                marge_fournisseur *= ratio
                marge_courtier *= ratio
                marge_totale = 25.0

        prix_finaux = {}
        commission = {}

        for key, value in prix_p0.items():
            if key.startswith('prix_') and not key.endswith(('_2025', '_2026', '_capa')):
                try:
                    p0 = float(value)
                    prix_final = p0 + marge_totale
                    prix_finaux[key] = round(prix_final, 2)

                    # Commission courtier sur ce poste
                    commission[key] = round(marge_courtier, 2)
                except:
                    prix_finaux[key] = value
                    commission[key] = 0
            else:
                prix_finaux[key] = value

        return {
            'prix_finaux': prix_finaux,
            'commission': commission,
            'marge_fournisseur': round(marge_fournisseur, 2) if marge_fournisseur else 0,
            'marge_courtier': round(marge_courtier, 2),
            'marge_totale': round(marge_totale, 2)
        }


if __name__ == '__main__':
    # Test
    grille = GrilleTarifaire()

    # Test C4
    prix = grille.get_prix_p0('C4', '21/11/2025', '31/12/2026')
    if prix:
        print(f"\n✅ Prix P0 trouvés pour C4 (21/11/2025 -> 31/12/2026):")
        print(f"  HPH: {prix['prix_hph']} €/MWh")
        print(f"  HCH: {prix['prix_hch']} €/MWh")
        print(f"  HPE: {prix['prix_hpe']} €/MWh")
        print(f"  HCE: {prix['prix_hce']} €/MWh")
        print(f"  Capacité 2026: {prix['prix_capa_2026']} €/MWh")

        # Test calcul avec marge
        result = grille.calculer_prix_avec_marge(prix, 10.0, 10.0)
        print(f"\n💰 Avec marge (10€ fournisseur + 10€ courtier):")
        print(f"  HPH: {result['prix_finaux']['prix_hph']} €/MWh")
        print(f"  Commission courtier: {result['commission']['prix_hph']} €/MWh")
