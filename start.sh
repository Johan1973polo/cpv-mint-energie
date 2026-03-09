#!/bin/bash

echo "================================"
echo "🚀 Lancement de l'application CPV MINT ENERGIE"
echo "================================"

# Vérifier si l'environnement virtuel existe
if [ ! -d "venv" ]; then
    echo "📦 Création de l'environnement virtuel..."
    python3 -m venv venv
    echo "✅ Environnement créé"
fi

# Activer l'environnement virtuel
echo "🔧 Activation de l'environnement virtuel..."
source venv/bin/activate

# Installer les dépendances si nécessaire
echo "📥 Vérification des dépendances..."
pip install -q -r requirements.txt

echo ""
echo "================================"
echo "✅ Application prête !"
echo "================================"
echo "📍 Ouvrez votre navigateur à l'adresse :"
echo "   👉 http://localhost:5000"
echo "================================"
echo ""

# Lancer l'application
python app.py
