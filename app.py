"""
Application Flask pour le remplissage automatique des CPV
"""
import os
import json
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pathlib import Path
from pdf_extractor import extract_all_pdfs
from simple_docx_generator import SimpleDocxGenerator
from grille_tarifaire import GrilleTarifaire

app = Flask(__name__)

# Charger les grilles tarifaires au démarrage
grille_tarifaire = GrilleTarifaire()
app.config['SECRET_KEY'] = 'cpv-mint-energie-2025'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max

# Créer les dossiers nécessaires
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Session data (en production, utiliser une vraie base de données)
session_data = {}


@app.route('/')
def index():
    """Page d'accueil avec upload"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    """Upload et extraction des fichiers"""
    try:
        # Vérifier les fichiers
        if 'files[]' not in request.files:
            return jsonify({'error': 'Aucun fichier fourni'}), 400

        files = request.files.getlist('files[]')

        if len(files) == 0:
            return jsonify({'error': 'Aucun fichier sélectionné'}), 400

        # Créer un dossier temporaire pour cette session
        session_id = os.urandom(16).hex()
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)

        # Sauvegarder les fichiers
        uploaded_files = []
        template_file = None

        for file in files:
            if file.filename == '':
                continue

            filename = secure_filename(file.filename)
            filepath = os.path.join(session_folder, filename)
            file.save(filepath)
            uploaded_files.append(filename)

            # Identifier le template CPV
            if filename.endswith('.txt') or 'cpv' in filename.lower():
                template_file = filepath

        print(f"📁 Fichiers uploadés: {uploaded_files}")

        # Extraire les données des PDFs
        extracted_data = extract_all_pdfs(session_folder)

        # Vérifications de validation
        segment = extracted_data.get('segment', '').upper()
        score = extracted_data.get('score', '0/10')
        nombre_pdl = int(extracted_data.get('nombre_pdl', '1'))
        volume_total = float(extracted_data.get('volume_total', '0').replace(',', '.'))

        # Parser le score
        try:
            score_str = score.split('/')[0].strip()
            score_value = float(score_str) if score_str else 0
        except:
            score_value = 0

        # Validations
        warnings = []
        errors = []

        # Score minimum - DÉSACTIVÉ POUR TEST
        # if score_value == 0:
        #     errors.append(f'❌ SCORE NON RENSEIGNÉ - Veuillez renseigner le score dans la Fiche Contact (minimum 5/10 OBLIGATOIRE)')
        # elif score_value < 5:
        #     errors.append(f'❌ SCORE INSUFFISANT: {score_value}/10 - Minimum 5/10 OBLIGATOIRE pour tous les segments !')

        # Validation C5 - AVERTISSEMENT
        if segment == 'C5':
            if nombre_pdl < 5:
                errors.append(f'❌ C5 - MINIMUM 5 SITES OBLIGATOIRE ! (actuellement: {nombre_pdl} site{"s" if nombre_pdl > 1 else ""})')

        # Validation C4 - Score + Consommation
        if segment == 'C4':
            if volume_total < 300:
                warnings.append(f'✅ C4 - Score OK ({score}) et Consommation OK ({volume_total} MWh < 300 MWh)')
            else:
                warnings.append(f'⚠️ C4 - Consommation SUPÉRIEURE à 300 MWh ({volume_total} MWh) → PASSER SUR PRICER PERSONNALISÉ')

        # Validation C2 - Score + Consommation
        if segment == 'C2':
            if volume_total < 300:
                warnings.append(f'✅ C2 - Score OK ({score}) et Consommation OK ({volume_total} MWh < 300 MWh)')
            else:
                warnings.append(f'⚠️ C2 - Consommation SUPÉRIEURE à 300 MWh ({volume_total} MWh) → PASSER SUR PRICER PERSONNALISÉ')

        # Stocker les données de session
        session_data[session_id] = {
            'extracted_data': extracted_data,
            'template_file': template_file,
            'session_folder': session_folder
        }

        return jsonify({
            'success': True,
            'session_id': session_id,
            'data': extracted_data,
            'warnings': warnings,
            'errors': errors,
            'segment': segment,
            'score': score,
            'volume_total': volume_total,
            'nombre_pdl': nombre_pdl
        })

    except Exception as e:
        print(f"❌ Erreur lors de l'upload: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/form/<session_id>')
def show_form(session_id):
    """Affiche le formulaire de saisie adaptatif"""
    if session_id not in session_data:
        return "Session invalide", 404

    data = session_data[session_id]['extracted_data']
    segment = data.get('segment', 'C4').upper()

    return render_template('form_v2.html',
                         session_id=session_id,
                         data=data,
                         segment=segment)


@app.route('/generate/<session_id>', methods=['POST'])
def generate_cpv(session_id):
    """Génère le CPV pré-rempli"""
    try:
        if session_id not in session_data:
            return jsonify({'error': 'Session invalide'}), 404

        # Récupérer les données
        extracted_data = session_data[session_id]['extracted_data']

        # Récupérer les données du formulaire
        form_data = request.form.to_dict()

        # Utiliser le template DOCX
        template_docx = 'template_cpv.docx'

        # Générer le CPV
        docx_generator = SimpleDocxGenerator(template_docx)
        raison_sociale = extracted_data.get('raison_sociale', 'CLIENT').replace(' ', '_').replace('.', '_').replace('-', '_')
        output_filename = f'CPV_MINT_{raison_sociale}.docx'
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # Générer le fichier (DOCX puis conversion en PDF)
        generated_file = docx_generator.generate(output_path, extracted_data, form_data)
        output_filename = os.path.basename(generated_file)

        return jsonify({
            'success': True,
            'filename': output_filename,
            'download_url': f'/download/{output_filename}'
        })

    except Exception as e:
        print(f"❌ Erreur lors de la génération: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/download/<filename>')
def download_file(filename):
    """Télécharge le CPV généré"""
    try:
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        return f"Erreur: {str(e)}", 404


@app.route('/api/get_prix_p0', methods=['POST'])
def get_prix_p0():
    """API pour obtenir les prix P0 selon segment et dates"""
    try:
        data = request.json
        segment = data.get('segment', '').upper()
        date_debut = data.get('date_debut', '')
        date_fin = data.get('date_fin', '')

        if not all([segment, date_debut, date_fin]):
            return jsonify({'error': 'Paramètres manquants'}), 400

        # Récupérer les prix P0
        prix_p0 = grille_tarifaire.get_prix_p0(segment, date_debut, date_fin)

        if not prix_p0:
            return jsonify({'error': 'Aucun prix trouvé pour ces dates'}), 404

        return jsonify({
            'success': True,
            'prix_p0': prix_p0,
            'segment': segment,
            'date_debut': date_debut,
            'date_fin': date_fin
        })

    except Exception as e:
        print(f"❌ Erreur API get_prix_p0: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/calculer_prix_avec_marge', methods=['POST'])
def calculer_prix_avec_marge():
    """API pour calculer les prix finaux avec marges"""
    try:
        data = request.json
        prix_p0 = data.get('prix_p0', {})
        marge_fournisseur = float(data.get('marge_fournisseur', 0))
        marge_courtier = float(data.get('marge_courtier', 0))

        if not prix_p0:
            return jsonify({'error': 'Prix P0 manquants'}), 400

        # Calculer les prix avec marge
        result = grille_tarifaire.calculer_prix_avec_marge(
            prix_p0, marge_fournisseur, marge_courtier
        )

        return jsonify({
            'success': True,
            **result
        })

    except Exception as e:
        print(f"❌ Erreur API calculer_prix_avec_marge: {str(e)}")
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("\n" + "="*60)
    print("🚀 APPLICATION CPV MINT ENERGIE")
    print("="*60)
    print("📍 Accédez à l'application: http://localhost:5001")
    print("="*60 + "\n")

    app.run(debug=False, host='0.0.0.0', port=5001)
