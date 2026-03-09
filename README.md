# 📄 Application CPV MINT ENERGIE

Application web locale pour le remplissage automatique des Conditions Particulières de Vente (CPV).

## 🚀 Démarrage Rapide

### Option 1 : Script automatique (recommandé)
```bash
./start.sh
```

### Option 2 : Manuelle

1. Créer l'environnement virtuel :
```bash
python3 -m venv venv
source venv/bin/activate
```

2. Installer les dépendances :
```bash
pip install -r requirements.txt
```

3. Lancer l'application :
```bash
python app.py
```

4. Ouvrir votre navigateur à l'adresse :
```
http://localhost:5000
```

## 📋 Documents requis

L'application nécessite les documents suivants :

1. **Fiche Contact** (PDF) - Contient les informations client et énergétiques
2. **Document RGPD** (PDF) - Consentement et données complémentaires
3. **Avis de situation SIREN** ou **Extrait d'immatriculation** (PDF) - Informations juridiques
4. **Template CPV** (TXT) - Le modèle de CPV à remplir

## ⚙️ Fonctionnalités

### Détection automatique du segment
- **C2** : 5 postes horosaisonniers (PTE, HPH, HCH, HPE, HCE)
- **C4** : 4 postes horosaisonniers (HPH, HCH, HPE, HCE)
- **C5** : 2 options (BASE ou HP/HC)

### Validations automatiques
- ✅ Score minimum 5/10 obligatoire
- ✅ C5 : Rappel minimum 5 SITES
- ✅ Consommation > 300 MWh → Alerte pricer personnalisé

### Formulaire adaptatif
Le formulaire s'adapte automatiquement selon :
- Le segment détecté (C2/C4/C5)
- L'option choisie pour C5 (BASE ou HP/HC)
- Validation en temps réel du total de consommation

### Données extraites automatiquement
- Raison sociale
- SIREN / SIRET
- Signataire et coordonnées
- Adresse siège et consommation
- PDL, puissance, segment
- Dates de livraison
- Volume total
- Forme juridique et code APE

### Données à saisir manuellement
- Capital social
- Fonction du signataire
- Répartition consommation par poste
- Prix de fourniture
- Prix abonnement et capacité
- Garanties d'origine
- CEE
- IBAN / BIC

## 📁 Structure du projet

```
cpv_app/
├── app.py                  # Application Flask principale
├── pdf_extractor.py        # Extraction des données PDFs
├── cpv_generator.py        # Génération du CPV
├── requirements.txt        # Dépendances Python
├── templates/
│   ├── index.html         # Page d'upload
│   └── form.html          # Formulaire de saisie
├── uploads/               # Fichiers uploadés (temporaire)
└── output/                # CPV générés
```

## 🔧 Utilisation

1. **Upload** : Déposer les 4 documents requis
2. **Validation** : L'application vérifie le score et la consommation
3. **Formulaire** : Remplir les champs manquants selon le segment
4. **Génération** : Télécharger le CPV pré-rempli

## 📝 Notes

- L'application fonctionne en local (pas de connexion internet requise)
- Les fichiers uploadés sont stockés temporairement
- Le CPV généré est au format TXT
- Tous les champs sont pré-remplis quand les données sont disponibles

## 🛠️ Support

Pour toute question ou problème, vérifiez :
- Que tous les PDFs sont bien présents
- Que le template CPV est au bon format
- Les logs dans la console de l'application
