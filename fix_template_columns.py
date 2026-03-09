#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script pour fixer les largeurs de colonnes NAF et FTA dans le template Word
Modifie directement le XML <w:tblGrid> du template
"""

from docx import Document
from lxml import etree
from docx.oxml.ns import qn


def fix_template_columns(template_path):
    """Modifie les largeurs de colonnes NAF et FTA dans le template"""

    print(f"📂 Ouverture du template: {template_path}")
    doc = Document(template_path)

    tables_modified = 0

    for idx, table in enumerate(doc.tables):
        # Récupérer le texte de l'en-tête pour identifier les tableaux de sites
        header_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells])

        if 'PRM' in header_text and 'NAF' in header_text:
            print(f"\n✅ Tableau {idx+1} trouvé (Sites): {header_text[:80]}...")

            # Lire le tblGrid actuel
            tblGrid = table._element.find(qn('w:tblGrid'))
            if tblGrid is not None:
                cols = tblGrid.findall(qn('w:gridCol'))
                print(f"   → {len(cols)} colonnes trouvées")

                # Afficher les largeurs actuelles
                print("   Largeurs AVANT modification:")
                for i, col in enumerate(cols):
                    w = col.get(qn('w:w'))
                    print(f"      Col {i}: {w} twips")

                # Identifier NAF (index 2) et FTA (index 6) et les élargir
                # Prendre de la largeur sur Adresse (index 3) qui est souvent trop large
                # NAF: mettre au moins 1200 twips (≈2.1cm)
                # FTA: mettre au moins 1500 twips (≈2.6cm)

                if len(cols) >= 7:
                    naf_col = cols[2]
                    addr_col = cols[3]
                    fta_col = cols[6] if len(cols) > 6 else None

                    # Récupérer les valeurs actuelles
                    naf_w = int(naf_col.get(qn('w:w'), 0))
                    addr_w = int(addr_col.get(qn('w:w'), 0))
                    fta_w = int(fta_col.get(qn('w:w'), 0)) if fta_col is not None else 0

                    # Calculer combien on doit ajouter
                    naf_target = max(naf_w, 1200)
                    fta_target = max(fta_w, 1500)

                    # Prendre la différence sur Adresse
                    diff = (naf_target - naf_w) + (fta_target - fta_w)
                    new_addr = addr_w - diff

                    # Appliquer les modifications
                    naf_col.set(qn('w:w'), str(naf_target))
                    addr_col.set(qn('w:w'), str(max(new_addr, 1500)))  # Adresse min 1500
                    if fta_col is not None:
                        fta_col.set(qn('w:w'), str(fta_target))

                    print(f"\n   Largeurs APRÈS modification:")
                    print(f"      NAF (col 2):     {naf_w} → {naf_target} twips")
                    print(f"      Adresse (col 3): {addr_w} → {max(new_addr, 1500)} twips")
                    if fta_col is not None:
                        print(f"      FTA (col 6):     {fta_w} → {fta_target} twips")

                    tables_modified += 1
                else:
                    print(f"   ⚠️  Pas assez de colonnes ({len(cols)} < 7)")
            else:
                print(f"   ⚠️  tblGrid introuvable")

    # Sauvegarder
    print(f"\n💾 Sauvegarde du template modifié...")
    doc.save(template_path)
    print(f"✅ {tables_modified} tableau(x) modifié(s) et sauvegardé(s) !")

    return tables_modified


if __name__ == '__main__':
    template_path = 'template_cpv_2026.docx'

    print("=" * 80)
    print("FIX DES LARGEURS DE COLONNES DU TEMPLATE WORD")
    print("=" * 80)

    count = fix_template_columns(template_path)

    print("\n" + "=" * 80)
    print(f"TERMINÉ - {count} tableau(x) modifié(s)")
    print("=" * 80)
