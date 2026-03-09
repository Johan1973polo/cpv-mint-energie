"""
Module de validations métier MINT Énergie 2026
Toutes les règles de validation pour les CPV
"""
import re
from datetime import datetime


class ValidationError(Exception):
    """Exception pour les erreurs de validation bloquantes"""
    pass


class ValidationWarning(Exception):
    """Exception pour les avertissements (non bloquants)"""
    pass


class ValidateurCPV:
    """Validateur pour les CPV MINT Énergie"""

    # Codes NAF exclus (santé/hébergement médicalisé)
    NAF_EXCLUS = ['8610Z', '8710Z']

    # FTA disponibles par segment
    FTA_PAR_SEGMENT = {
        'C2': ['HTLU', 'HTCU'],
        'C4': ['MU4', 'CU4', 'LU4'],
        'C5': ['BTINF']
    }

    # Limites par segment
    LIMITES = {
        'C2': {
            'car_max': 300,  # MWh
            'sites_max': None
        },
        'C4': {
            'car_max': 300,  # MWh
            'sites_max': None
        },
        'C5': {
            'car_max': 300,  # MWh
            'sites_max': 100
        }
    }

    def __init__(self):
        self.erreurs = []
        self.avertissements = []

    def reset(self):
        """Réinitialise les erreurs et avertissements"""
        self.erreurs = []
        self.avertissements = []

    def valider_score(self, score, segment=None):
        """
        Valide le score client selon le segment

        Règles:
        - C2 et C4 : Score < 5 génère un avertissement (non bloquant)
        - C5 : Score < 5 génère une erreur (bloquant)

        Args:
            score: Score au format "X/10" ou float
            segment: 'C2', 'C4' ou 'C5' (optionnel)

        Returns:
            bool: True si valide (peut avoir des avertissements)

        Raises:
            ValidationError: Si score < 5 pour C5
        """
        try:
            if isinstance(score, str):
                if '/' in score:
                    score_value = float(score.split('/')[0].strip())
                else:
                    score_value = float(score)
            else:
                score_value = float(score)

            if score_value < 5:
                # Pour C2 et C4 : Avertissement mais pas bloquant
                if segment in ['C2', 'C4']:
                    self.avertissements.append(
                        f"⚠️ Score < 5/10 ({score_value}/10) — Demande de dérogation obligatoire auprès de MINT Énergie avant envoi du contrat"
                    )
                    return True  # Non bloquant pour C2/C4
                else:
                    # Pour C5 (ou segment non spécifié) : Erreur bloquante
                    self.erreurs.append(
                        f"❌ Score insuffisant : {score_value}/10 - Minimum 5/10 OBLIGATOIRE"
                    )
                    return False

            return True

        except Exception as e:
            self.erreurs.append(f"❌ Score invalide : {score}")
            return False

    def valider_car_total(self, segment, car_total):
        """
        Valide le CAR total selon le segment

        Args:
            segment: 'C2', 'C4' ou 'C5'
            car_total: Consommation Annuelle de Référence en MWh

        Returns:
            bool: True si valide (peut avoir des avertissements)
        """
        car_total = float(car_total)
        car_max = self.LIMITES[segment]['car_max']

        if car_total > car_max:
            self.avertissements.append(
                f"⚠️ CAR total ({car_total:.2f} MWh) > {car_max} MWh → "
                f"Cotation sur-mesure requise (contacter devcommercial@mint.eco)"
            )

        return True

    def valider_nombre_sites(self, segment, nombre_sites):
        """
        Valide le nombre de sites selon le segment

        Args:
            segment: 'C2', 'C4' ou 'C5'
            nombre_sites: Nombre de PRM

        Returns:
            bool: True si valide

        Raises:
            ValidationError: Si limites dépassées
        """
        nombre_sites = int(nombre_sites)
        sites_max = self.LIMITES[segment].get('sites_max')

        if sites_max and nombre_sites > sites_max:
            self.erreurs.append(
                f"❌ Segment {segment} : Maximum {sites_max} sites autorisés (actuellement : {nombre_sites})"
            )
            return False

        return True

    def valider_prm(self, prm):
        """
        Valide le format d'un PRM (14 chiffres)

        Args:
            prm: Numéro de PRM

        Returns:
            bool: True si valide
        """
        prm_str = str(prm).strip()

        if not re.match(r'^\d{14}$', prm_str):
            self.erreurs.append(f"❌ PRM invalide : {prm} (doit être 14 chiffres)")
            return False

        return True

    def valider_siret(self, siret):
        """
        Valide le format d'un SIRET (14 chiffres)

        Args:
            siret: Numéro de SIRET

        Returns:
            bool: True si valide
        """
        siret_str = str(siret).strip().replace(' ', '')

        if not re.match(r'^\d{14}$', siret_str):
            self.erreurs.append(f"❌ SIRET invalide : {siret} (doit être 14 chiffres)")
            return False

        return True

    def valider_siren(self, siren):
        """
        Valide le format d'un SIREN (9 chiffres)

        Args:
            siren: Numéro de SIREN

        Returns:
            bool: True si valide
        """
        siren_str = str(siren).strip().replace(' ', '')

        if not re.match(r'^\d{9}$', siren_str):
            self.erreurs.append(f"❌ SIREN invalide : {siren} (doit être 9 chiffres)")
            return False

        return True

    def valider_naf(self, naf):
        """
        Valide le code NAF (format NNNNX) et vérifie les exclusions

        Args:
            naf: Code NAF

        Returns:
            bool: True si valide

        Raises:
            ValidationError: Si code NAF exclu
        """
        naf_str = str(naf).strip().upper()

        # Format : 4 chiffres + 1 lettre
        if not re.match(r'^\d{4}[A-Z]$', naf_str):
            self.erreurs.append(f"❌ Code NAF invalide : {naf} (format attendu : NNNNX, ex: 5610A)")
            return False

        # Vérifier les exclusions
        if naf_str in self.NAF_EXCLUS:
            self.erreurs.append(
                f"❌ Code NAF exclu : {naf_str} (santé/hébergement médicalisé non éligible)"
            )
            return False

        return True

    def valider_fta(self, segment, fta):
        """
        Valide la FTA selon le segment

        Args:
            segment: 'C2', 'C4' ou 'C5'
            fta: Formule Tarifaire d'Acheminement

        Returns:
            bool: True si valide
        """
        fta_str = str(fta).strip().upper()
        ftas_valides = self.FTA_PAR_SEGMENT.get(segment, [])

        if fta_str not in ftas_valides:
            self.erreurs.append(
                f"❌ FTA invalide pour {segment} : {fta} (autorisés : {', '.join(ftas_valides)})"
            )
            return False

        return True

    def valider_marge(self, marge):
        """
        Valide la marge courtier (6-25 €/MWh)

        Args:
            marge: Marge en €/MWh

        Returns:
            bool: True si valide
        """
        try:
            marge_value = float(marge)

            if marge_value < 6:
                self.erreurs.append(f"❌ Marge trop faible : {marge_value} €/MWh (minimum 6 €/MWh)")
                return False

            if marge_value > 25:
                self.erreurs.append(f"❌ Marge trop élevée : {marge_value} €/MWh (maximum 25 €/MWh)")
                return False

            return True

        except:
            self.erreurs.append(f"❌ Marge invalide : {marge}")
            return False

    def valider_duree(self, duree_mois):
        """
        Valide la durée du contrat (minimum 12 mois)

        Args:
            duree_mois: Durée en mois

        Returns:
            bool: True si valide
        """
        try:
            duree = int(duree_mois)

            if duree < 12:
                self.erreurs.append(f"❌ Durée insuffisante : {duree} mois (minimum 12 mois)")
                return False

            return True

        except:
            self.erreurs.append(f"❌ Dur��e invalide : {duree_mois}")
            return False

    def valider_dates(self, date_debut, date_fin):
        """
        Valide les dates de fourniture selon les règles MINT 2026-2029

        Règles (mail Adam MAURIN 03/11/2025):
        - Pas de limite de date fixe (la grille Excel définit les dates max)
        - Un contrat ne peut PAS couvrir uniquement l'année 2029
        - Il doit couvrir au minimum 2 années dont 2028 + 2029

        Args:
            date_debut: Date de début (format DD/MM/YYYY)
            date_fin: Date de fin (format DD/MM/YYYY)

        Returns:
            bool: True si valide
        """
        try:
            # Parser les dates
            if isinstance(date_debut, str):
                dt_debut = datetime.strptime(date_debut, '%d/%m/%Y')
            else:
                dt_debut = date_debut

            if isinstance(date_fin, str):
                dt_fin = datetime.strptime(date_fin, '%d/%m/%Y')
            else:
                dt_fin = date_fin

            # Date fin > Date début
            if dt_fin <= dt_debut:
                self.erreurs.append(
                    f"❌ Date de fin ({date_fin}) doit être postérieure à la date de début ({date_debut})"
                )
                return False

            # Nouvelle règle 2029 : Le contrat ne peut pas couvrir UNIQUEMENT 2029
            # Il doit couvrir au minimum 2 années (2028 + 2029 par exemple)
            annee_debut = dt_debut.year
            annee_fin = dt_fin.year

            # Si le contrat commence en 2029 ou après ET se termine en 2029
            # => Le contrat ne couvre que 2029, ce qui est interdit
            if annee_debut >= 2029 and annee_fin <= 2029:
                self.erreurs.append(
                    f"❌ Le contrat ne peut pas couvrir uniquement 2029 — "
                    f"il doit inclure au minimum 2028 + 2029 (date début : {date_debut}, date fin : {date_fin})"
                )
                return False

            return True

        except Exception as e:
            self.erreurs.append(f"❌ Dates invalides : {e}")
            return False

    def valider_site(self, site_data, segment):
        """
        Valide toutes les données d'un site

        Args:
            site_data: dict avec les données du site
            segment: 'C2', 'C4' ou 'C5'

        Returns:
            bool: True si toutes les validations passent
        """
        valide = True

        # PRM
        if 'prm' in site_data:
            valide &= self.valider_prm(site_data['prm'])

        # SIRET
        if 'siret' in site_data:
            valide &= self.valider_siret(site_data['siret'])

        # NAF
        if 'naf' in site_data:
            valide &= self.valider_naf(site_data['naf'])

        # FTA
        if 'fta' in site_data:
            valide &= self.valider_fta(segment, site_data['fta'])

        # Dates
        if 'date_debut' in site_data and 'date_fin' in site_data:
            valide &= self.valider_dates(site_data['date_debut'], site_data['date_fin'])

        return valide

    def valider_contrat_complet(self, data):
        """
        Valide un contrat complet

        Args:
            data: dict avec toutes les données du contrat

        Returns:
            dict: {'valide': bool, 'erreurs': list, 'avertissements': list}
        """
        self.reset()

        # Segment (extraire AVANT la validation du score)
        segment = data.get('segment', 'C4').upper()

        # Score client (accepter 'score' ou 'score_client')
        score_value = data.get('score_client') or data.get('score')
        if score_value:
            self.valider_score(score_value, segment=segment)

        # CAR total
        if 'car_total' in data:
            self.valider_car_total(segment, data['car_total'])

        # Nombre de sites
        if 'sites' in data:
            nombre_sites = len(data['sites'])
            self.valider_nombre_sites(segment, nombre_sites)

            # Valider chaque site
            for i, site in enumerate(data['sites']):
                self.valider_site(site, segment)

        # Marge courtier
        if 'marge_courtier' in data:
            self.valider_marge(data['marge_courtier'])

        # Durée
        if 'duree_mois' in data:
            self.valider_duree(data['duree_mois'])

        # Dates
        if 'date_debut' in data and 'date_fin' in data:
            self.valider_dates(data['date_debut'], data['date_fin'])

        # SIREN client
        if 'siren' in data:
            self.valider_siren(data['siren'])

        return {
            'valide': len(self.erreurs) == 0,
            'erreurs': self.erreurs,
            'avertissements': self.avertissements
        }


if __name__ == '__main__':
    # Tests unitaires
    validateur = ValidateurCPV()

    print("="*60)
    print("TESTS DE VALIDATION")
    print("="*60)

    # Test score
    print("\n✅ Test Score :")
    print(f"  Score 7/10 : {validateur.valider_score('7/10')}")
    validateur.reset()
    print(f"  Score 3/10 (C5) : {validateur.valider_score('3/10', segment='C5')}")
    print(f"  Erreurs : {validateur.erreurs}")
    validateur.reset()
    print(f"  Score 3/10 (C4) : {validateur.valider_score('3/10', segment='C4')}")
    print(f"  Avertissements : {validateur.avertissements}")
    validateur.reset()

    # Test CAR
    print("\n✅ Test CAR :")
    validateur.valider_car_total('C4', 250)
    print(f"  CAR 250 MWh : OK")
    validateur.valider_car_total('C4', 350)
    print(f"  CAR 350 MWh : {validateur.avertissements}")
    validateur.reset()

    # Test NAF
    print("\n✅ Test NAF :")
    print(f"  5610A (restaurant) : {validateur.valider_naf('5610A')}")
    validateur.reset()
    print(f"  8610Z (santé EXCLU) : {validateur.valider_naf('8610Z')}")
    print(f"  Erreurs : {validateur.erreurs}")
    validateur.reset()

    # Test marge
    print("\n✅ Test Marge :")
    print(f"  10 €/MWh : {validateur.valider_marge(10)}")
    validateur.reset()
    print(f"  3 €/MWh : {validateur.valider_marge(3)}")
    print(f"  Erreurs : {validateur.erreurs}")
    validateur.reset()

    # Test contrat complet
    print("\n✅ Test Contrat Complet :")
    contrat_test = {
        'score': '7/10',
        'segment': 'C4',
        'car_total': 180,
        'sites': [
            {
                'prm': '12345678901234',
                'siret': '44376263800023',
                'naf': '5610A',
                'fta': 'MU4',
                'date_debut': '01/01/2026',
                'date_fin': '31/12/2027'
            }
        ],
        'marge_courtier': 10,
        'duree_mois': 24,
        'date_debut': '01/01/2026',
        'date_fin': '31/12/2027',
        'siren': '443762638'
    }

    result = validateur.valider_contrat_complet(contrat_test)
    print(f"  Valide : {result['valide']}")
    print(f"  Erreurs : {result['erreurs']}")
    print(f"  Avertissements : {result['avertissements']}")
