#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de test pour le générateur DOCX CPV 2026
"""

import json
from docx_generator_2026 import CPVGenerator2026

# Données de test
extracted_data = {
    'raison_sociale': 'SARL TEST ENERGIE',
    'siren': '123456789',
    'adresse_siege': '10 rue de la Test, 75001 Paris',
    'forme_juridique': 'SARL',
    'capital_social': '10000',
    'ville_rcs': 'Paris',
    'nom_gerant': 'Jean DUPONT',
    'nom_signataire': 'Jean DUPONT',
    'fonction_signataire': 'Gérant',
    'email': 'jean.dupont@test.fr',
    'telephone': '0601020304',
    'adresse_facturation': '10 rue de la Test, 75001 Paris',
    'iban': 'FR76 1234 5678 9012 3456 7890 123',
    'bic': 'BNPAFRPPXXX'
}

form_data = {
    'date_debut': '01/03/2026',
    'date_fin': '28/02/2027',
    'duree_mois': '12',
    'segment': 'C4',
    'site_count': '2',

    # Site 1
    'site_1_prm': '12345678901234',
    'site_1_siret': '12345678900001',
    'site_1_naf': '4711F',
    'site_1_adresse': '10 rue de la Test, 75001 Paris',
    'site_1_segment': 'C4',
    'site_1_fta': 'C4',
    'site_1_puissance': '96',
    'site_1_cee': 'non_soumis',

    # Site 2
    'site_2_prm': '12345678901235',
    'site_2_siret': '12345678900002',
    'site_2_naf': '4711F',
    'site_2_adresse': '20 rue de la Test, 75001 Paris',
    'site_2_segment': 'C4',
    'site_2_fta': 'C4',
    'site_2_puissance': '120',
    'site_2_cee': 'non_soumis',

    # Consommations
    'conso_pointe': '10',
    'conso_hph': '20',
    'conso_hch': '15',
    'conso_hpe': '25',
    'conso_hce': '30',

    # Prix
    'prix_p0_data': json.dumps({
        'prix_finaux': {
            'pointe': 150.50,
            'hph': 120.25,
            'hch': 100.75,
            'hpe': 110.30,
            'hce': 95.60
        },
        'coefficient_alpha': 0.1234
    }),

    'marge_courtier': '10',
    'score_client': '8',
    'garantie_paiement': 'Non',
    'flexibilite_c4': 'Non'
}

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🧪 TEST GÉNÉRATEUR DOCX CPV 2026")
    print("="*60)

    try:
        generator = CPVGenerator2026('template_cpv_2026.docx')
        output_path = 'output/TEST_CPV_2026.docx'

        print("\n📝 Génération du CPV de test...")
        generated_file = generator.generate(output_path, extracted_data, form_data)

        print("\n" + "="*60)
        print(f"✅ CPV GÉNÉRÉ AVEC SUCCÈS !")
        print(f"📄 Fichier: {generated_file}")
        print("="*60)
        print("\nVous pouvez maintenant ouvrir le fichier pour vérifier le remplissage.")

    except Exception as e:
        print(f"\n❌ ERREUR: {e}")
        import traceback
        traceback.print_exc()
