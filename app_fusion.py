"""
Application Flask FUSIONNÉE pour le remplissage automatique des CPV
Version 2026 avec support PDFs + Excel + Hybride
"""
import os
import json
import subprocess
import shutil
import secrets
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
from pathlib import Path
from pdf_extractor import extract_all_pdfs
from excel_parser import ExcelGrilleParser
from grille_tarifaire import GrilleTarifaire
from validations import ValidateurCPV
from docx_generator_2026 import CPVGenerator2026
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from dotenv import load_dotenv
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from functools import wraps
from models import db, User
import platform

load_dotenv()

# LibreOffice path - Mac vs Linux/Docker
if platform.system() == 'Darwin':
    LIBREOFFICE_PATH = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
else:
    LIBREOFFICE_PATH = 'libreoffice'

app = Flask(__name__)

# Configuration Flask
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'cpv-mint-2026-dev-key')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max

# Configuration SQLAlchemy - support SQLite local et PostgreSQL Render
database_url = os.getenv('DATABASE_URL', 'sqlite:///users.db')
# Render utilise postgres:// mais SQLAlchemy nécessite postgresql://
if database_url.startswith('postgres://'):
    database_url = database_url.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = database_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Initialiser SQLAlchemy
db.init_app(app)

# Configuration Flask-Login
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Veuillez vous connecter pour accéder à cette page.'
login_manager.login_message_category = 'warning'

# Créer les dossiers nécessaires au démarrage
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs('instance', exist_ok=True)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Créer les dossiers nécessaires
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Session data (en production, utiliser une vraie base de données ou Redis)
session_data = {}

# Grille tarifaire par défaut (CSV)
grille_tarifaire_default = GrilleTarifaire()

# Validateur
validateur = ValidateurCPV()


def envoyer_cpv_par_mail(filepath, filename, raison_sociale, segment, siren='N/A', nb_sites=0, commission_totale=0):
    """Envoi silencieux du CPV par email"""
    try:
        SMTP_SERVER = 'smtp.gmail.com'
        SMTP_PORT = 587
        EMAIL_FROM = os.environ.get('SMTP_EMAIL')
        EMAIL_PASSWORD = os.environ.get('SMTP_PASSWORD')
        EMAIL_TO = os.environ.get('CPV_NOTIFY_EMAIL')

        if not all([EMAIL_FROM, EMAIL_PASSWORD, EMAIL_TO]):
            print("   ⚠️ Variables d'environnement manquantes pour l'envoi d'email")
            return

        msg = MIMEMultipart()
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        msg['Subject'] = f'🔋 CPV MINT - {raison_sociale} ({segment}) - Commission {int(commission_totale):,}€'.replace(',', ' ')

        body = f"""Nouveau CPV généré :

📋 Raison sociale : {raison_sociale}
🔢 SIREN : {siren}
⚡ Segment(s) : {segment}
🏢 Nombre de sites : {nb_sites}
💰 Commission totale : {int(commission_totale):,}€
📄 Fichier : {filename}
📅 Date : {datetime.now().strftime('%d/%m/%Y à %H:%M')}

Le CPV est en pièce jointe.
"""
        msg.attach(MIMEText(body, 'plain'))

        with open(filepath, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={filename}')
            msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
        server.quit()

        print(f"   📧 Email envoyé silencieusement à {EMAIL_TO}")
    except Exception as e:
        print(f"   ⚠️ Erreur silencieuse lors de l'envoi d'email: {e}")


# Filtre Jinja pour corriger les nombres avec virgule
@app.template_filter('decimal_point')
def decimal_point_filter(value):
    """Convertit virgule en point pour les inputs type='number' HTML"""
    if value is None or value == '':
        return ''
    # Convertir en string et remplacer virgule par point
    return str(value).replace(',', '.')


@app.route('/')
@login_required
def index():
    """Page d'accueil avec choix du workflow"""
    return render_template('index_fusion.html')


@app.route('/upload', methods=['POST'])
@login_required
def upload_files():
    """
    Workflow 1: Upload et extraction des PDFs uniquement
    Utilise les grilles CSV par défaut
    """
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

        for file in files:
            if file.filename == '':
                continue

            filename = secure_filename(file.filename)
            filepath = os.path.join(session_folder, filename)
            file.save(filepath)
            uploaded_files.append(filename)

        print(f"📁 Workflow 1 - Fichiers uploadés: {uploaded_files}")

        # Extraire les données des PDFs
        extracted_data = extract_all_pdfs(session_folder)

        # Validations
        segment = extracted_data.get('segment', 'C4').upper()
        score = extracted_data.get('score', '0/10')
        nombre_pdl = int(extracted_data.get('nombre_pdl', '1'))
        volume_total = float(extracted_data.get('volume_total', '0').replace(',', '.'))

        # Valider avec le module de validations
        validation_result = validateur.valider_contrat_complet({
            'score': score,
            'segment': segment,
            'car_total': volume_total,
            'sites': [{'prm': str(i)} for i in range(nombre_pdl)]
        })

        # Stocker les données de session
        session_data[session_id] = {
            'workflow': 1,
            'extracted_data': extracted_data,
            'session_folder': session_folder,
            'grille_tarifaire': grille_tarifaire_default  # Grilles CSV
        }

        return jsonify({
            'success': True,
            'session_id': session_id,
            'data': extracted_data,
            'warnings': validation_result['avertissements'],
            'errors': validation_result['erreurs'],
            'segment': segment,
            'score': score,
            'volume_total': volume_total,
            'nombre_pdl': nombre_pdl
        })

    except Exception as e:
        print(f"❌ Erreur lors de l'upload (workflow 1): {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/upload_excel', methods=['POST'])
@login_required
def upload_excel():
    """
    Workflow 2: Upload de la grille Excel uniquement
    Saisie manuelle des données client
    """
    session_folder = None  # Pour cleanup en cas d'erreur
    try:
        # Vérifier le fichier Excel
        if 'excel_file' not in request.files:
            return jsonify({'error': 'Fichier Excel manquant'}), 400

        excel_file = request.files['excel_file']

        if excel_file.filename == '':
            return jsonify({'error': 'Aucun fichier sélectionné'}), 400

        # ✅ VALIDATION: Vérifier l'extension du fichier Excel AVANT de sauvegarder
        filename = secure_filename(excel_file.filename)
        if not filename.lower().endswith(('.xlsx', '.xlsm')):
            return jsonify({
                'error': f'Format de fichier invalide : "{filename}". '
                        f'Seuls les fichiers .xlsx et .xlsm sont acceptés.'
            }), 400

        # Créer un dossier temporaire pour cette session
        session_id = os.urandom(16).hex()
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)

        # Sauvegarder le fichier Excel
        excel_path = os.path.join(session_folder, filename)
        excel_file.save(excel_path)

        print(f"📊 Workflow 2 - Excel chargé: {filename}")

        # Parser l'Excel
        excel_parser = ExcelGrilleParser(excel_path)
        excel_parser.parse_all()

        # Créer la grille tarifaire depuis l'Excel
        grille_tarifaire = GrilleTarifaire(excel_parser=excel_parser)

        # Stocker les données de session
        session_data[session_id] = {
            'workflow': 2,
            'excel_parser': excel_parser,
            'grille_tarifaire': grille_tarifaire,
            'session_folder': session_folder,
            'extracted_data': {}  # Pas de données extraites, saisie manuelle
        }

        # Récupérer les métadonnées
        metadata = excel_parser.get_metadata()

        return jsonify({
            'success': True,
            'session_id': session_id,
            'metadata': metadata,
            'message': 'Grille Excel chargée avec succès'
        })

    except Exception as e:
        # ✅ CLEANUP: Nettoyer le dossier temporaire en cas d'erreur
        if session_folder and os.path.exists(session_folder):
            import shutil
            try:
                shutil.rmtree(session_folder)
                print(f"🧹 Dossier temporaire nettoyé après erreur: {session_folder}")
            except Exception as cleanup_error:
                print(f"⚠️  Impossible de nettoyer {session_folder}: {cleanup_error}")

        print(f"❌ Erreur lors de l'upload Excel (workflow 2): {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/upload_hybride', methods=['POST'])
@login_required
def upload_hybride():
    """
    Workflow 3: Upload PDFs + Excel (RECOMMANDÉ)
    Extraction automatique + grilles Excel du jour
    """
    session_folder = None  # Pour cleanup en cas d'erreur
    try:
        # Vérifier les fichiers
        if 'excel_file' not in request.files or 'files[]' not in request.files:
            return jsonify({'error': 'Fichiers manquants (Excel + PDFs requis)'}), 400

        excel_file = request.files['excel_file']
        pdf_files = request.files.getlist('files[]')

        if excel_file.filename == '' or len(pdf_files) == 0:
            return jsonify({'error': 'Tous les fichiers doivent être fournis'}), 400

        # ✅ VALIDATION: Vérifier l'extension du fichier Excel AVANT de sauvegarder
        excel_filename = secure_filename(excel_file.filename)
        if not excel_filename.lower().endswith(('.xlsx', '.xlsm')):
            return jsonify({
                'error': f'Format de fichier invalide : "{excel_filename}". '
                        f'Seuls les fichiers .xlsx et .xlsm sont acceptés.'
            }), 400

        # Créer un dossier temporaire pour cette session
        session_id = os.urandom(16).hex()
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)

        # Sauvegarder l'Excel
        excel_path = os.path.join(session_folder, excel_filename)
        excel_file.save(excel_path)

        # Sauvegarder les PDFs
        pdf_filenames = []
        for pdf_file in pdf_files:
            if pdf_file.filename != '':
                filename = secure_filename(pdf_file.filename)
                filepath = os.path.join(session_folder, filename)
                pdf_file.save(filepath)
                pdf_filenames.append(filename)

        print(f"🚀 Workflow 3 - Excel: {excel_filename}, PDFs: {pdf_filenames}")

        # Parser l'Excel
        excel_parser = ExcelGrilleParser(excel_path)
        excel_parser.parse_all()

        # Créer la grille tarifaire depuis l'Excel
        grille_tarifaire = GrilleTarifaire(excel_parser=excel_parser)

        # Extraire les données des PDFs
        extracted_data = extract_all_pdfs(session_folder)

        # Validations
        segment = extracted_data.get('segment', 'C4').upper()
        score = extracted_data.get('score', '0/10')
        nombre_pdl = int(extracted_data.get('nombre_pdl', '1'))
        volume_total = float(extracted_data.get('volume_total', '0').replace(',', '.'))

        validation_result = validateur.valider_contrat_complet({
            'score': score,
            'segment': segment,
            'car_total': volume_total,
            'sites': [{'prm': str(i)} for i in range(nombre_pdl)]
        })

        # Stocker les données de session
        session_data[session_id] = {
            'workflow': 3,
            'excel_parser': excel_parser,
            'grille_tarifaire': grille_tarifaire,
            'extracted_data': extracted_data,
            'session_folder': session_folder
        }

        metadata = excel_parser.get_metadata()

        return jsonify({
            'success': True,
            'session_id': session_id,
            'data': extracted_data,
            'metadata': metadata,
            'warnings': validation_result['avertissements'],
            'errors': validation_result['erreurs'],
            'segment': segment,
            'score': score,
            'volume_total': volume_total,
            'nombre_pdl': nombre_pdl
        })

    except Exception as e:
        # ✅ CLEANUP: Nettoyer le dossier temporaire en cas d'erreur
        if session_folder and os.path.exists(session_folder):
            import shutil
            try:
                shutil.rmtree(session_folder)
                print(f"🧹 Dossier temporaire nettoyé après erreur: {session_folder}")
            except Exception as cleanup_error:
                print(f"⚠️  Impossible de nettoyer {session_folder}: {cleanup_error}")

        print(f"❌ Erreur lors de l'upload hybride (workflow 3): {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/config/<session_id>')
@login_required
def show_config(session_id):
    """
    Page de configuration pour workflow 2 (Excel seulement)
    Permet de choisir segment, date, durée, marge
    """
    if session_id not in session_data:
        return "Session invalide", 404

    data = session_data[session_id]
    excel_parser = data.get('excel_parser')

    if not excel_parser:
        return "Pas de grille Excel chargée", 400

    # Récupérer les dates disponibles pour chaque segment
    dates_c2 = excel_parser.get_dates_disponibles('C2')
    dates_c4 = excel_parser.get_dates_disponibles('C4')
    dates_c5 = excel_parser.get_dates_disponibles('C5')

    return render_template('config.html',
                         session_id=session_id,
                         metadata=excel_parser.metadata,
                         dates_c2=dates_c2,
                         dates_c4=dates_c4,
                         dates_c5=dates_c5)


@app.route('/api/get_durees', methods=['POST'])
@login_required
def get_durees_disponibles():
    """API pour récupérer les durées disponibles selon segment et date"""
    try:
        data = request.json
        session_id = data.get('session_id')
        segment = data.get('segment')
        date_debut = data.get('date_debut')

        if not session_id or session_id not in session_data:
            return jsonify({'error': 'Session invalide'}), 404

        # Essayer d'utiliser excel_parser d'abord, sinon grille_tarifaire
        excel_parser = session_data[session_id].get('excel_parser')
        grille = session_data[session_id].get('grille_tarifaire')

        if excel_parser:
            durees = excel_parser.get_durees_disponibles(segment, date_debut)
        elif grille:
            durees = grille.get_durees_disponibles(segment, date_debut)
        else:
            return jsonify({'error': 'Pas de grille chargée'}), 400

        return jsonify({
            'success': True,
            'durees': durees
        })

    except Exception as e:
        print(f"❌ Erreur get_durees: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/get_durees/<session_id>', methods=['POST'])
@login_required
def get_durees(session_id):
    """API pour récupérer les durées disponibles selon segment et date"""
    try:
        if session_id not in session_data:
            return jsonify({'error': 'Session invalide'}), 404

        data = request.json
        segment = data.get('segment')
        date_debut = data.get('date_debut')

        excel_parser = session_data[session_id].get('excel_parser')
        if not excel_parser:
            return jsonify({'error': 'Pas de grille Excel'}), 400

        durees = excel_parser.get_durees_disponibles(segment, date_debut)

        return jsonify({
            'success': True,
            'durees': durees
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/get_prix_p0', methods=['POST'])
@login_required
def get_prix_p0_without_url_param():
    """API pour récupérer les prix P0 - session_id dans le body"""
    try:
        data = request.json
        session_id = data.get('session_id')

        if not session_id or session_id not in session_data:
            return jsonify({'error': 'Session invalide'}), 404

        segment = data.get('segment')
        date_debut = data.get('date_debut')
        date_grille = data.get('date_grille')  # Date du 1er du mois pour chercher les prix
        date_fin = data.get('date_fin')
        duree_mois = data.get('duree_mois')
        marge_courtier = float(data.get('marge_courtier', 10))

        # Essayer d'utiliser excel_parser d'abord, sinon grille_tarifaire
        excel_parser = session_data[session_id].get('excel_parser')
        grille = session_data[session_id].get('grille_tarifaire')

        if not excel_parser and not grille:
            return jsonify({'error': 'Pas de grille chargée'}), 400

        # Utiliser date_grille si fournie, sinon date_debut (rétro-compatibilité)
        date_pour_prix = date_grille if date_grille else date_debut

        # Mode Excel : utiliser duree_mois
        if excel_parser and duree_mois:
            # DEBUG BUG 2: Afficher le format de date reçu et les clés de la grille
            print(f"🔍 Date grille reçue: '{date_pour_prix}' (type: {type(date_pour_prix)})")
            print(f"🔍 Segment: {segment}, Durée: {duree_mois} mois")

            # Analyser la structure de la grille
            if hasattr(excel_parser, 'grilles') and segment in excel_parser.grilles:
                grille_data = excel_parser.grilles[segment]
                print(f"🔍 Nombre de lignes dans grille {segment}: {len(grille_data)}")

                if grille_data:
                    # Afficher les colonnes disponibles
                    print(f"🔍 Colonnes disponibles: {list(grille_data[0].keys())}")

                    # Afficher première ligne complète
                    print(f"🔍 Première ligne: {grille_data[0]}")

                    # Extraire toutes les dates uniques
                    dates_uniques = sorted(set(row.get('date_debut', '') for row in grille_data if row.get('date_debut')))
                    print(f"🔍 Dates UNIQUES (10 premières): {dates_uniques[:10]}")
                    print(f"🔍 Nombre total de dates uniques: {len(dates_uniques)}")

                    # Chercher les durées disponibles pour la date recherchée
                    durees_pour_date = [row.get('duree_mois') for row in grille_data if row.get('date_debut') == date_pour_prix]
                    print(f"🔍 Durées disponibles pour {date_pour_prix}: {sorted(set(durees_pour_date))}")

            prix_p0 = excel_parser.get_prix_p0(segment, date_pour_prix, int(duree_mois))
            if not prix_p0:
                # Lister les dates disponibles pour aider l'utilisateur
                dates_disponibles = excel_parser.get_dates_disponibles(segment)
                if dates_disponibles:
                    dates_range = f"{dates_disponibles[0]} à {dates_disponibles[-1]}"
                    return jsonify({
                        'error': f'⚠️ Pas de grille disponible pour ce mois.\n\nDates disponibles : {dates_range}'
                    }), 404
                return jsonify({'error': f'Aucun prix trouvé pour {segment} du {date_pour_prix} ({duree_mois} mois)'}), 404

            # Calculer avec marge
            prix_finaux = excel_parser.calculer_prix_avec_marge(prix_p0, marge_courtier)

            return jsonify({
                'success': True,
                'prix_p0': prix_p0,
                'prix_finaux': prix_finaux,
                'marge_courtier': marge_courtier,
                'coefficient_alpha': prix_p0.get('coefficient_alpha', excel_parser.metadata.get('coefficient_alpha', 0))
            })

        # Mode CSV : utiliser date_fin
        elif grille and date_fin:
            prix_p0 = grille.get_prix_p0(segment, date_pour_prix, date_fin=date_fin)
            if not prix_p0:
                return jsonify({'error': f'Aucun prix trouvé pour {segment} du {date_pour_prix} au {date_fin}'}), 404

            # Calculer avec marge
            prix_avec_marge = grille.calculer_prix_avec_marge(prix_p0, marge_courtier=marge_courtier)

            return jsonify({
                'success': True,
                'prix_p0': prix_p0,
                'prix_finaux': prix_avec_marge['prix_finaux'],
                'marge_courtier': prix_avec_marge['marge_courtier'],
                'coefficient_alpha': prix_p0.get('coefficient_alpha', 0)
            })

        else:
            return jsonify({'error': 'Paramètres manquants (duree_mois ou date_fin)'}), 400

    except Exception as e:
        print(f"❌ Erreur get_prix_p0: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/get_prix_p0/<session_id>', methods=['POST'])
@login_required
def get_prix_p0_with_url_param(session_id):
    """API pour récupérer les prix P0 selon segment, date et durée"""
    try:
        if session_id not in session_data:
            return jsonify({'error': 'Session invalide'}), 404

        data = request.json
        segment = data.get('segment')
        date_debut = data.get('date_debut')
        duree_mois = int(data.get('duree_mois'))
        marge_courtier = float(data.get('marge_courtier', 10))

        grille = session_data[session_id].get('grille_tarifaire')
        if not grille:
            return jsonify({'error': 'Pas de grille tarifaire'}), 400

        # Récupérer les prix P0
        prix_p0 = grille.get_prix_p0(segment, date_debut, duree_mois=duree_mois)

        if not prix_p0:
            return jsonify({'error': 'Aucun prix trouvé pour ces paramètres'}), 404

        # Calculer avec marge
        prix_avec_marge = grille.calculer_prix_avec_marge(prix_p0, marge_courtier=marge_courtier)

        return jsonify({
            'success': True,
            'prix_p0': prix_p0,
            'prix_finaux': prix_avec_marge['prix_finaux'],
            'marge_courtier': prix_avec_marge['marge_courtier'],
            'coefficient_alpha': prix_p0.get('coefficient_alpha', 0)
        })

    except Exception as e:
        print(f"❌ Erreur get_prix_p0: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/form/<session_id>')
@login_required
def show_form(session_id):
    """Affiche le formulaire de saisie adaptatif (workflows 1 et 3)"""
    if session_id not in session_data:
        return "Session invalide", 404

    data = session_data[session_id]['extracted_data']
    workflow = session_data[session_id]['workflow']
    segment = data.get('segment', 'C4').upper()

    # MODIFICATION 3: Mapper C3 → C2 (C3 est traité comme C2)
    if segment == 'C3':
        segment = 'C2'
        data['segment'] = 'C2'  # Mettre à jour aussi dans les data
        print(f"   ✓ Segment C3 détecté → mappé vers C2")

    # DEBUG: Afficher les données extraites
    print(f"📊 Données extraites pour session {session_id}:")
    print(f"   - raison_sociale: {data.get('raison_sociale', 'NON TROUVÉ')}")
    print(f"   - siren: {data.get('siren', 'NON TROUVÉ')}")
    print(f"   - segment: {segment}")
    print(f"   - nom_gerant: {data.get('nom_gerant', 'NON TROUVÉ')}")
    print(f"   - email: {data.get('email', 'NON TROUVÉ')}")
    print(f"   - siret_complet: {data.get('siret_complet', 'NON TROUVÉ')}")
    print(f"   - code_naf: {data.get('code_naf', 'NON TROUVÉ')}")
    print(f"   - prm_list: {data.get('prm_list', 'NON TROUVÉ')}")
    print(f"   - puissance: {data.get('puissance', 'NON TROUVÉ')}")
    print(f"   - adresse_site: {data.get('adresse_site', 'NON TROUVÉ')}")
    print(f"   - type_calendrier: {data.get('type_calendrier', 'NON TROUVÉ')}")

    # Récupérer la grille tarifaire (CSV ou Excel selon workflow)
    grille = session_data[session_id].get('grille_tarifaire')

    return render_template('form_cpv_2026.html',
                         session_id=session_id,
                         data=data,
                         segment=segment,
                         workflow=workflow,
                         has_excel=(workflow == 3))


def convert_date_iso_to_fr(date_str):
    """Convertit YYYY-MM-DD en DD/MM/YYYY pour le validateur"""
    if not date_str:
        return date_str

    # Déjà au format français ?
    try:
        datetime.strptime(date_str, '%d/%m/%Y')
        return date_str
    except ValueError:
        pass

    # Format ISO → français
    try:
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        return dt.strftime('%d/%m/%Y')
    except ValueError:
        # Format non reconnu, retourner tel quel
        return date_str


@app.route('/generate/<session_id>', methods=['POST'])
@login_required
def generate_cpv(session_id):
    """Génère le CPV pré-rempli"""
    import sys
    print("\n" + "="*80, flush=True)
    print("🔴🔴🔴 VERSION v2.3 - TABLEAUX SITES COMPLETS 🔴🔴🔴", flush=True)
    print("🚀 DÉBUT GÉNÉRATION CPV", flush=True)
    print("="*80, flush=True)
    sys.stdout.flush()
    sys.stderr.flush()

    try:
        print(f"✓ Session ID reçu: {session_id}", flush=True)

        if session_id not in session_data:
            print(f"❌ Session {session_id} non trouvée dans session_data", flush=True)
            print(f"   Sessions disponibles: {list(session_data.keys())}", flush=True)
            return jsonify({'error': 'Session invalide'}), 404

        print(f"✓ Session trouvée", flush=True)

        # Récupérer les données
        extracted_data = session_data[session_id].get('extracted_data', {})
        print(f"✓ extracted_data récupéré: {len(extracted_data)} clés", flush=True)

        # Récupérer les données du formulaire
        form_data = request.form.to_dict()
        print(f"✓ form_data récupéré: {len(form_data)} clés", flush=True)

        # ========================================
        # ENRICHISSEMENT FORCÉ depuis extracted_data
        # ========================================
        print(f"\n🔧 ENRICHISSEMENT FORM_DATA depuis extracted_data:", flush=True)

        # Raison sociale
        if not form_data.get('raison_sociale'):
            form_data['raison_sociale'] = extracted_data.get('raison_sociale', '')
            print(f"   → raison_sociale forcée: '{form_data['raison_sociale']}'", flush=True)

        # SIREN
        if not form_data.get('siren'):
            form_data['siren'] = extracted_data.get('siren', '')
            print(f"   → siren forcé: '{form_data['siren']}'", flush=True)

        # Segment
        if not form_data.get('segment'):
            form_data['segment'] = extracted_data.get('segment', 'C4')
            print(f"   → segment forcé: '{form_data['segment']}'", flush=True)

        # Ville RCS (priorité Pappers) - FORCER même si form_data contient déjà une valeur
        if extracted_data.get('ville_rcs'):
            form_data['ville_rcs'] = extracted_data.get('ville_rcs')
            print(f"   → ville_rcs forcée depuis Pappers: '{form_data['ville_rcs']}'", flush=True)

        # Adresse siège (priorité Pappers)
        if not form_data.get('adresse_siege') and extracted_data.get('adresse_siege'):
            form_data['adresse_siege'] = extracted_data.get('adresse_siege')
            print(f"   → adresse_siege forcée: '{form_data['adresse_siege'][:50]}...'", flush=True)

        # Code NAF (priorité Pappers)
        if not form_data.get('code_naf') and extracted_data.get('code_naf'):
            form_data['code_naf'] = extracted_data.get('code_naf')
            print(f"   → code_naf forcé: '{form_data['code_naf']}'", flush=True)

        # Capital social
        if not form_data.get('capital_social') and extracted_data.get('capital_social'):
            form_data['capital_social'] = extracted_data.get('capital_social')
            print(f"   → capital_social forcé: '{form_data['capital_social']}'", flush=True)

        # Forme juridique
        if not form_data.get('forme_juridique') and extracted_data.get('forme_juridique'):
            form_data['forme_juridique'] = extracted_data.get('forme_juridique')
            print(f"   → forme_juridique forcée: '{form_data['forme_juridique']}'", flush=True)

        # Nom gérant (priorité Pappers)
        if not form_data.get('nom_gerant') and extracted_data.get('nom_gerant'):
            form_data['nom_gerant'] = extracted_data.get('nom_gerant')
            print(f"   → nom_gerant forcé: '{form_data['nom_gerant']}'", flush=True)

        # BUG FIX 5: Gestion personne morale comme dirigeant
        # Si le dirigeant INPI est une société (pas une personne), utiliser le nom du gérant RGPD
        if extracted_data.get('personne_morale_dirigeant'):
            print(f"\n🔧 BUG FIX 5: Personne morale détectée comme dirigeant", flush=True)
            print(f"   → Société dirigeante: '{extracted_data.get('denomination_dirigeant')}'", flush=True)
            if extracted_data.get('nom_gerant'):
                form_data['nom_contact'] = extracted_data.get('nom_gerant')
                print(f"   → Nom signataire (depuis RGPD): '{form_data['nom_contact']}'", flush=True)
            else:
                print(f"   ⚠️ Nom gérant RGPD non trouvé!", flush=True)
        elif extracted_data.get('nom_signataire_inpi'):
            # Personne physique: utiliser le nom de l'INPI
            if not form_data.get('nom_contact'):
                form_data['nom_contact'] = extracted_data.get('nom_signataire_inpi')
                print(f"   → nom_contact (depuis INPI): '{form_data['nom_contact']}'", flush=True)

        # Vérification finale
        print(f"\n✅ FORM_DATA ENRICHI:", flush=True)
        print(f"   raison_sociale = '{form_data.get('raison_sociale')}'", flush=True)
        print(f"   siren = '{form_data.get('siren')}'", flush=True)
        print(f"   segment = '{form_data.get('segment')}'", flush=True)
        print(f"   ville_rcs = '{form_data.get('ville_rcs')}'", flush=True)
        print(f"   code_naf = '{form_data.get('code_naf')}'", flush=True)

        # ========================================
        # CORRECTIONS POUR TOUS LES SEGMENTS (C2, C4, C5)
        # ========================================
        segment = form_data.get('segment', '').upper()
        print(f"\n🔧 CORRECTIONS POUR SEGMENT {segment}:", flush=True)

        # 1. Calculer les dates globales depuis les sites si vides (TOUS segments)
        date_debut = form_data.get('date_debut', '')
        date_fin = form_data.get('date_fin', '')
        site_count = int(form_data.get('site_count', 0))

        if not date_debut or not date_fin:
            print(f"   → Dates globales vides, calcul depuis les {site_count} site(s)...", flush=True)

            # Collecter toutes les dates des sites
            dates_debut = []
            dates_fin = []
            for i in range(1, site_count + 1):
                d_debut = form_data.get(f'date_debut_site_{i}', '')
                d_fin = form_data.get(f'date_fin_site_{i}', '')
                if d_debut:
                    dates_debut.append(d_debut)
                if d_fin:
                    dates_fin.append(d_fin)

            # Prendre la plus ancienne date_debut et la plus tardive date_fin
            if dates_debut and dates_fin:
                # Convertir en datetime pour comparaison correcte (format ISO YYYY-MM-DD)
                dates_debut_dt = [datetime.strptime(d, '%Y-%m-%d') for d in dates_debut if d]
                dates_fin_dt = [datetime.strptime(d, '%Y-%m-%d') for d in dates_fin if d]

                # Trouver min/max puis reconvertir en string ISO
                date_debut_calculee = min(dates_debut_dt).strftime('%Y-%m-%d')
                date_fin_calculee = max(dates_fin_dt).strftime('%Y-%m-%d')

                form_data['date_debut'] = date_debut_calculee
                form_data['date_fin'] = date_fin_calculee

                print(f"      ✓ date_debut calculée: {date_debut_calculee}", flush=True)
                print(f"      ✓ date_fin calculée: {date_fin_calculee}", flush=True)

                # Calculer duree_mois
                duree_mois = max(1, round((max(dates_fin_dt) - min(dates_debut_dt)).days / 30.44))
                form_data['duree_mois'] = str(duree_mois)
                print(f"      ✓ duree_mois calculée: {duree_mois} mois", flush=True)
            else:
                print(f"      ⚠️ Impossible de calculer les dates (aucune date de site trouvée)", flush=True)

        # 2. Recalculer les prix P0 si nécessaire (pour C5 principalement)
        if segment == 'C5':
            prix_p0_data_raw = form_data.get('prix_p0_data', '{}')
            try:
                prix_p0_data = json.loads(prix_p0_data_raw)

                # Si prix_p0_data ne contient que la marge (ou est vide), recalculer
                if not prix_p0_data.get('prix_p0') or not prix_p0_data.get('prix_finaux'):
                    print(f"   → Prix P0 vides ou incomplets, recalcul depuis la grille Excel...", flush=True)

                    excel_parser = session_data[session_id].get('excel_parser')
                    if excel_parser and form_data.get('date_debut') and form_data.get('duree_mois'):
                        date_debut_prix = form_data.get('date_debut')
                        duree_mois_prix = int(form_data.get('duree_mois', 12))
                        marge = float(form_data.get('marge_courtier', 10))

                        # Convertir la date en format DD/MM/YYYY pour l'Excel parser
                        try:
                            date_dt = datetime.strptime(date_debut_prix, '%Y-%m-%d')
                            date_debut_prix_fr = date_dt.strftime('%d/%m/%Y')
                        except ValueError:
                            # Déjà au format français
                            date_debut_prix_fr = date_debut_prix

                        # Récupérer les prix P0 depuis la grille Excel
                        prix_p0 = excel_parser.get_prix_p0('C5', date_debut_prix_fr, duree_mois_prix)

                        if prix_p0:
                            # Calculer les prix finaux avec la marge
                            prix_finaux = excel_parser.calculer_prix_avec_marge(prix_p0, marge)

                            # Mettre à jour prix_p0_data
                            prix_p0_data = {
                                'prix_p0': prix_p0,
                                'prix_finaux': prix_finaux,
                                'marge_courtier': marge,
                                'coefficient_alpha': prix_p0.get('coefficient_alpha', excel_parser.metadata.get('coefficient_alpha', 0))
                            }

                            form_data['prix_p0_data'] = json.dumps(prix_p0_data)
                            print(f"      ✓ Prix P0 recalculés pour C5 du {date_debut_prix_fr} sur {duree_mois_prix} mois", flush=True)
                            print(f"      ✓ Exemple prix final BASE: {prix_finaux.get('prix_base', 0):.2f} €/MWh", flush=True)
                        else:
                            print(f"      ⚠️ Aucun prix trouvé dans la grille pour {date_debut_prix_fr} / {duree_mois_prix} mois", flush=True)
                    else:
                        print(f"      ⚠️ Impossible de recalculer (excel_parser ou dates manquants)", flush=True)
            except Exception as e:
                print(f"      ⚠️ Erreur lors du recalcul des prix C5: {e}", flush=True)

        # ========================================
        # ENRICHISSEMENT prix_p0_data avec la marge
        # ========================================
        if form_data.get('prix_p0_data'):
            try:
                prix_data = json.loads(form_data.get('prix_p0_data', '{}'))
                marge = float(form_data.get('marge_courtier', 10))

                print(f"\n💰 RECALCUL PRIX avec marge {marge} €/MWh:", flush=True)

                # Si prix_p0 existe, recalculer prix_finaux
                if prix_data.get('prix_p0') and isinstance(prix_data['prix_p0'], dict):
                    if 'prix_finaux' not in prix_data:
                        prix_data['prix_finaux'] = {}

                    for key, value in prix_data['prix_p0'].items():
                        if key.startswith('prix_') and isinstance(value, (int, float)):
                            prix_data['prix_finaux'][key] = round(float(value) + marge, 2)
                            print(f"   → {key}: {value} + {marge} = {prix_data['prix_finaux'][key]}", flush=True)

                    prix_data['marge_courtier'] = marge
                    prix_data['prix_finaux']['marge_courtier'] = marge
                    form_data['prix_p0_data'] = json.dumps(prix_data)
                    print(f"   ✓ prix_p0_data recalculé avec succès", flush=True)
                else:
                    print(f"   ℹ️ prix_p0 non trouvé dans prix_p0_data, utilisation des prix existants", flush=True)
            except Exception as e:
                print(f"   ⚠️ Erreur lors du recalcul des prix: {e}", flush=True)

        # ========================================
        # CONVERSION DATES ISO → FORMAT FRANÇAIS
        # ========================================
        print(f"\n📅 CONVERSION DATES (ISO → FR pour le validateur):", flush=True)

        # Convertir les dates globales
        if form_data.get('date_debut'):
            date_avant = form_data.get('date_debut')
            form_data['date_debut'] = convert_date_iso_to_fr(date_avant)
            if date_avant != form_data['date_debut']:
                print(f"   → date_debut: {date_avant} → {form_data['date_debut']}", flush=True)

        if form_data.get('date_fin'):
            date_avant = form_data.get('date_fin')
            form_data['date_fin'] = convert_date_iso_to_fr(date_avant)
            if date_avant != form_data['date_fin']:
                print(f"   → date_fin: {date_avant} → {form_data['date_fin']}", flush=True)

        # Convertir les dates de chaque site (pour C5)
        site_count = int(form_data.get('site_count', 0))
        if site_count > 0:
            for i in range(1, site_count + 1):
                # Date début site
                key_debut = f'date_debut_site_{i}'
                if form_data.get(key_debut):
                    date_avant = form_data.get(key_debut)
                    form_data[key_debut] = convert_date_iso_to_fr(date_avant)
                    if date_avant != form_data[key_debut]:
                        print(f"   → {key_debut}: {date_avant} → {form_data[key_debut]}", flush=True)

                # Date fin site
                key_fin = f'date_fin_site_{i}'
                if form_data.get(key_fin):
                    date_avant = form_data.get(key_fin)
                    form_data[key_fin] = convert_date_iso_to_fr(date_avant)
                    if date_avant != form_data[key_fin]:
                        print(f"   → {key_fin}: {date_avant} → {form_data[key_fin]}", flush=True)

        print(f"   ✓ Conversion des dates terminée", flush=True)

        # ========================================
        # DEBUG CRITIQUE: AFFICHER TOUTES LES DONNÉES PAR SITE
        # ========================================
        print(f"\n" + "="*80, flush=True)
        print(f"🔍 DEBUG DÉTAILLÉ - DONNÉES PAR SITE", flush=True)
        print(f"="*80, flush=True)
        print(f"📊 Nombre de sites: {site_count}", flush=True)

        for i in range(1, site_count + 1):
            print(f"\n🏢 SITE {i}/{site_count}:", flush=True)
            print(f"   ├─ PRM: {form_data.get(f'prm_{i}', 'MANQUANT')}", flush=True)
            print(f"   ├─ SIRET: {form_data.get(f'siret_{i}', 'MANQUANT')}", flush=True)
            print(f"   ├─ Code NAF: {form_data.get(f'naf_{i}', 'MANQUANT')}", flush=True)
            print(f"   ├─ Adresse: {form_data.get(f'adresse_site_{i}', 'MANQUANT')[:50]}...", flush=True)
            print(f"   ├─ SEGMENT: {form_data.get(f'site_{i}_segment', 'MANQUANT')}", flush=True)
            print(f"   ├─ Type Calendrier: {form_data.get(f'type_calendrier_{i}', 'MANQUANT')}", flush=True)
            print(f"   ├─ FTA: {form_data.get(f'fta_{i}', 'MANQUANT')}", flush=True)
            print(f"   ├─ Puissance: {form_data.get(f'puissance_{i}', 'MANQUANT')} kVA", flush=True)
            print(f"   ├─ Date Début: {form_data.get(f'date_debut_site_{i}', 'MANQUANT')}", flush=True)
            print(f"   ├─ Date Fin: {form_data.get(f'date_fin_site_{i}', 'MANQUANT')}", flush=True)
            print(f"   ├─ CAR Totale: {form_data.get(f'car_{i}', 'MANQUANT')} MWh", flush=True)

            # Prix P0 du site
            prix_p0_site = form_data.get(f'prix_p0_site_{i}', None)
            if prix_p0_site:
                print(f"   ├─ Prix P0 Site: OUI ({len(prix_p0_site)} chars)", flush=True)
                try:
                    prix_parsed = json.loads(prix_p0_site)
                    print(f"   │  └─ Contient: {list(prix_parsed.keys())}", flush=True)
                except:
                    print(f"   │  └─ Erreur parsing JSON", flush=True)
            else:
                print(f"   ├─ Prix P0 Site: ❌ MANQUANT", flush=True)

            # Détail CAR du site
            car_detail_site = form_data.get(f'car_detail_site_{i}', None)
            if car_detail_site:
                print(f"   ├─ Détail CAR Site: OUI ({len(car_detail_site)} chars)", flush=True)
                try:
                    car_parsed = json.loads(car_detail_site)
                    print(f"   │  └─ Contient: {car_parsed}", flush=True)
                except:
                    print(f"   │  └─ Erreur parsing JSON", flush=True)
            else:
                print(f"   └─ Détail CAR Site: ❌ MANQUANT", flush=True)

        print(f"\n" + "="*80, flush=True)

        # DEBUG CRITIQUE: Afficher prix_p0_data GLOBAL
        print(f"\n🔍 DEBUG PRIX_P0_DATA GLOBAL:", flush=True)
        prix_p0_data_raw = form_data.get('prix_p0_data')
        if prix_p0_data_raw:
            print(f"   Type: {type(prix_p0_data_raw)}", flush=True)
            print(f"   Longueur: {len(prix_p0_data_raw)} chars", flush=True)
            print(f"   Contenu brut: {prix_p0_data_raw[:500]}...", flush=True)
            try:
                prix_p0_data_parsed = json.loads(prix_p0_data_raw)
                print(f"   Parsé JSON: {prix_p0_data_parsed}", flush=True)
                print(f"   prix_finaux: {prix_p0_data_parsed.get('prix_finaux')}", flush=True)
                print(f"   coefficient_alpha: {prix_p0_data_parsed.get('coefficient_alpha')}", flush=True)
            except Exception as e:
                print(f"   ❌ Erreur parsing JSON: {e}", flush=True)
        else:
            print(f"   ❌ prix_p0_data ABSENT du formulaire!", flush=True)

        # DEBUG: Afficher les données reçues AVANT validation
        print(f"\n📥 DONNÉES REÇUES - Session {session_id}", flush=True)
        print(f"   Content-Type: {request.content_type}", flush=True)
        print(f"   Nombre de champs form: {len(form_data)}", flush=True)
        print(f"   Clés du formulaire: {list(form_data.keys())}", flush=True)
        print(f"\n   Détail des champs:", flush=True)
        for key, value in form_data.items():
            if isinstance(value, str) and len(value) > 100:
                print(f"      • {key}: {value[:100]}... ({len(value)} chars)", flush=True)
            else:
                print(f"      • {key}: {value}", flush=True)

        # Valider le contrat complet
        print(f"\n🔄 Préparation validation...")
        validation_data = {
            **extracted_data,
            **form_data
        }
        print(f"✓ validation_data fusionné: {len(validation_data)} clés")

        # Effectuer les validations
        print(f"🔄 Appel validateur.valider_contrat_complet()...")
        validation_result = validateur.valider_contrat_complet(validation_data)
        print(f"✓ Validation terminée")

        # DEBUG: Afficher les résultats de validation
        print(f"\n🔍 VALIDATION CPV - Session {session_id}")
        print(f"   Valide: {validation_result['valide']}")
        if validation_result['erreurs']:
            print(f"   ❌ Erreurs: {validation_result['erreurs']}")
        if validation_result['avertissements']:
            print(f"   ⚠️  Avertissements: {validation_result['avertissements']}")

        if not validation_result['valide']:
            print(f"\n❌ VALIDATION ÉCHOUÉE - RETOUR 400", flush=True)
            print(f"   Erreurs: {validation_result['erreurs']}", flush=True)
            return jsonify({
                'success': False,
                'erreurs': validation_result['erreurs'],
                'avertissements': validation_result['avertissements']
            }), 400

        # Utiliser le template DOCX 2026
        template_docx = 'template_cpv_2026.docx'

        # ========================================
        # GROUPER LES SITES PAR SEGMENT ET CALCULER LES VOLUMES
        # ========================================
        print(f"\n" + "="*80, flush=True)
        print(f"📊 GROUPEMENT DES SITES PAR SEGMENT", flush=True)
        print(f"="*80, flush=True)

        sites_par_segment = {'C2': [], 'C4': [], 'C5': []}

        for i in range(1, site_count + 1):
            site_segment = form_data.get(f'site_{i}_segment', '')

            # Parser les dates pour calculer la durée
            date_debut_str = form_data.get(f'date_debut_site_{i}', '')
            date_fin_str = form_data.get(f'date_fin_site_{i}', '')

            print(f"\n🔍 DEBUG VOLUME Site {i}:", flush=True)
            print(f"   date_debut_str reçue: '{date_debut_str}'", flush=True)
            print(f"   date_fin_str reçue: '{date_fin_str}'", flush=True)

            try:
                # Essayer format ISO d'abord (YYYY-MM-DD)
                try:
                    date_debut = datetime.strptime(date_debut_str, '%Y-%m-%d')
                    date_fin = datetime.strptime(date_fin_str, '%Y-%m-%d')
                except ValueError:
                    # Si échec, essayer format FR (DD/MM/YYYY)
                    date_debut = datetime.strptime(date_debut_str, '%d/%m/%Y')
                    date_fin = datetime.strptime(date_fin_str, '%d/%m/%Y')

                duration_days = (date_fin - date_debut).days
                duration_months = duration_days / 30.44  # Moyenne jours/mois
                print(f"   ✅ Parsing réussi: {duration_days} jours = {duration_months:.2f} mois", flush=True)
            except Exception as e:
                duration_months = 12  # Fallback
                print(f"   ❌ Parsing échoué ({e}), fallback = 12 mois", flush=True)

            # CAR en MWh/an
            car_mwh = float(form_data.get(f'car_{i}', 0))

            # Volume contractuel = CAR × (durée_mois / 12)
            volume_contractuel = car_mwh * (duration_months / 12)

            print(f"   CAR: {car_mwh} MWh/an", flush=True)
            print(f"   Durée: {duration_months:.2f} mois", flush=True)
            print(f"   ✅ Volume contractuel: {volume_contractuel:.2f} MWh", flush=True)

            # Parser Prix P0 du site
            prix_p0_site_raw = form_data.get(f'prix_p0_site_{i}', '')
            prix_p0_site = None
            if prix_p0_site_raw:
                try:
                    prix_p0_site = json.loads(prix_p0_site_raw)
                except:
                    pass

            # Parser détail CAR du site
            car_detail_site_raw = form_data.get(f'car_detail_site_{i}', '')
            car_detail_site = {}
            if car_detail_site_raw:
                try:
                    car_detail_site = json.loads(car_detail_site_raw)
                except:
                    pass

            # Créer l'objet site
            site_obj = {
                'prm': form_data.get(f'prm_{i}', ''),
                'siret': form_data.get(f'siret_{i}', ''),
                'naf': form_data.get(f'naf_{i}', ''),
                'adresse': form_data.get(f'adresse_site_{i}', ''),
                'segment': site_segment,
                'type_calendrier': form_data.get(f'type_calendrier_{i}', ''),
                'fta': form_data.get(f'fta_{i}', ''),
                'puissance': form_data.get(f'puissance_{i}', ''),
                'date_debut': date_debut_str,
                'date_fin': date_fin_str,
                'car_mwh': car_mwh,
                'volume_contractuel': volume_contractuel,
                'prix_p0_data': prix_p0_site,
                'car_detail': car_detail_site
            }

            # Ajouter au bon groupe
            if site_segment in sites_par_segment:
                sites_par_segment[site_segment].append(site_obj)
                print(f"   ✓ Site {i} ajouté au segment {site_segment} (Volume: {volume_contractuel:.2f} MWh)", flush=True)

        # Calculer les totaux par segment
        totaux_par_segment = {}
        for seg, sites_list in sites_par_segment.items():
            if sites_list:
                nb_prm = len(sites_list)
                volume_total = sum(s['volume_contractuel'] for s in sites_list)
                totaux_par_segment[seg] = {
                    'nb_prm': nb_prm,
                    'volume_total': volume_total,
                    'sites': sites_list
                }
                print(f"\n   📌 {seg}: {nb_prm} PRM, Volume total = {volume_total:.2f} MWh", flush=True)

        print(f"\n" + "="*80 + "\n", flush=True)

        # Générer le CPV avec le nouveau générateur 2026
        docx_generator = CPVGenerator2026(template_docx)
        raison_sociale = extracted_data.get('raison_sociale', form_data.get('raison_sociale', 'CLIENT'))
        raison_sociale_clean = raison_sociale.replace(' ', '_').replace('.', '_').replace('-', '_')

        # Nom de fichier avec date
        date_now = datetime.now().strftime('%Y%m%d')
        output_filename = f'CPV_MINT_{raison_sociale_clean}_{date_now}.docx'
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # Générer le fichier avec les sites groupés
        generated_file = docx_generator.generate(output_path, extracted_data, form_data, totaux_par_segment)
        output_filename = os.path.basename(generated_file)

        # Envoi silencieux par email
        try:
            segment = form_data.get('segment', 'N/A')
            siren = extracted_data.get('siren', form_data.get('siren', 'N/A'))

            # Calculer nb_sites total et commission totale
            nb_sites_total = sum(data['nb_prm'] for data in totaux_par_segment.values())

            # Calculer commission totale : (marge / 2) × CAR × (durée / 12)
            marge_courtier = float(form_data.get('marge_courtier', 0))
            commission_totale = 0
            for seg_data in totaux_par_segment.values():
                for site in seg_data['sites']:
                    car = site.get('car_mwh', 0)
                    # Recalculer durée en mois
                    try:
                        date_debut_str = site.get('date_debut', '')
                        date_fin_str = site.get('date_fin', '')
                        try:
                            date_debut = datetime.strptime(date_debut_str, '%Y-%m-%d')
                            date_fin = datetime.strptime(date_fin_str, '%Y-%m-%d')
                        except:
                            date_debut = datetime.strptime(date_debut_str, '%d/%m/%Y')
                            date_fin = datetime.strptime(date_fin_str, '%d/%m/%Y')
                        duration_months = (date_fin - date_debut).days / 30.44
                    except:
                        duration_months = 12

                    commission_site = (marge_courtier / 2) * car * (duration_months / 12)
                    commission_totale += commission_site

            envoyer_cpv_par_mail(
                generated_file,
                output_filename,
                raison_sociale,
                segment,
                siren=siren,
                nb_sites=nb_sites_total,
                commission_totale=commission_totale
            )
        except Exception as e:
            print(f"   ⚠️ Email notification failed: {e}")
            pass

        return jsonify({
            'success': True,
            'filename': output_filename,
            'download_url': f'/download/{output_filename}',
            'avertissements': validation_result['avertissements']
        })

    except Exception as e:
        print("\n" + "="*80, flush=True)
        print("💥 ERREUR LORS DE LA GÉNÉRATION", flush=True)
        print("="*80, flush=True)
        print(f"Type d'erreur: {type(e).__name__}", flush=True)
        print(f"Message: {str(e)}", flush=True)
        print("\n📋 TRACEBACK COMPLET:", flush=True)
        import traceback
        traceback.print_exc()
        sys.stdout.flush()
        sys.stderr.flush()
        print("="*80 + "\n", flush=True)
        return jsonify({
            'error': f"{type(e).__name__}: {str(e)}",
            'type': type(e).__name__
        }), 500


@app.route('/download/<filename>')
@login_required
def download_file(filename):
    """Télécharge le CPV généré"""
    try:
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        return f"Erreur: {str(e)}", 404


@app.route('/download_pdf/<filename>')
@login_required
def download_pdf(filename):
    """Convertit le DOCX en PDF et le télécharge"""
    try:
        docx_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)

        if not os.path.exists(docx_path):
            return f"Fichier DOCX non trouvé: {filename}", 404

        # Nom du fichier PDF
        pdf_filename = filename.replace('.docx', '.pdf')
        pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)

        # Utiliser le chemin LibreOffice détecté au démarrage
        soffice_path = LIBREOFFICE_PATH

        if platform.system() == 'Darwin' and not os.path.exists(soffice_path):
            return "LibreOffice non installé. Veuillez installer LibreOffice pour générer des PDFs.", 500

        # Conversion DOCX → PDF via LibreOffice headless
        print(f"📄 Conversion en PDF: {filename} → {pdf_filename}")
        result = subprocess.run([
            soffice_path,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', app.config['OUTPUT_FOLDER'],
            docx_path
        ], capture_output=True, text=True, timeout=60)

        if result.returncode != 0:
            print(f"❌ Erreur LibreOffice: {result.stderr}")
            return f"Erreur lors de la conversion PDF: {result.stderr}", 500

        if not os.path.exists(pdf_path):
            return "Erreur: le PDF n'a pas été généré", 500

        print(f"✅ PDF généré: {pdf_filename}")
        return send_file(pdf_path, as_attachment=True, download_name=pdf_filename)

    except subprocess.TimeoutExpired:
        return "Timeout lors de la conversion PDF (>60s)", 500
    except Exception as e:
        print(f"❌ Erreur conversion PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        return f"Erreur: {str(e)}", 500


@app.route('/process')
@login_required
def process_utilisateur():
    """Page de documentation du processus utilisateur avec screenshots"""
    # Dossier contenant les screenshots
    process_folder = os.path.join(os.path.dirname(__file__), 'PROCESS UTILISATEUR ')

    # Ordre personnalisé du processus utilisateur
    ordre_prioritaire = [
        'CLIQUEZ SUR COMMENCER.png',
        'METTRE LA FEUILLE EXCEL.png',
        'PAPPERS.png',
        'CHARGÉ PAPPERS ET RGPD.png',
        'METTRE LE SCORE CLIENT ELLIPRO.png',
        '2 LES INFOS CE METTENT TOUTES SEUL.png',
        '3 METTE LA DATE DEBUT ET FIN.png',
        '4 REPORTER LES VOLUMES DE CONSO ENEDIS.png',
        'CLACULER LES PRIX.png',
        'CHOISIR FTA EN FONCTION DU PROFIL.png',
        '6 RENSIGNER L IBAN MAIS PAS OBLIGATOIRE LAISSER LES OPTION CEE ET GO PAR DEFAULT.png',
        'CHOISSEZ PDF OU DOCS POUR POUVOIR MODIFIER.png',
        'GNERERE VOTRE CPV.png',
        'DIFFERENT FICHIER UTILE.png'
    ]

    # Récupérer tous les fichiers existants
    image_files = []
    if os.path.exists(process_folder):
        existing_files = set(os.listdir(process_folder))

        # Ajouter les fichiers dans l'ordre prioritaire s'ils existent
        for filename in ordre_prioritaire:
            if filename in existing_files and filename.lower().endswith('.png'):
                image_files.append(filename)

        # Ajouter les fichiers restants qui ne sont pas dans l'ordre prioritaire
        for filename in sorted(existing_files):
            if filename.lower().endswith('.png') and filename not in image_files:
                image_files.append(filename)

    return render_template('process_utilisateur.html', images=image_files)


@app.route('/process_images/<filename>')
@login_required
def process_image(filename):
    """Sert les images du dossier PROCESS UTILISATEUR"""
    process_folder = os.path.join(os.path.dirname(__file__), 'PROCESS UTILISATEUR ')
    return send_from_directory(process_folder, filename)


# ============================================================
# ROUTES D'AUTHENTIFICATION
# ============================================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Page de connexion"""
    # Si déjà connecté, rediriger vers l'accueil
    if current_user.is_authenticated:
        return redirect(url_for('index'))

    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')

        print(f"🔍 Tentative de connexion - Email: {email}")

        # Chercher l'utilisateur
        user = User.query.filter_by(email=email).first()

        if not user:
            print(f"❌ Utilisateur introuvable: {email}")
            return render_template('login.html', error='Email ou mot de passe incorrect.')

        print(f"✓ Utilisateur trouvé: {user.nom}")

        if not user.check_password(password):
            print(f"❌ Mot de passe incorrect pour {email}")
            return render_template('login.html', error='Email ou mot de passe incorrect.')

        print(f"✓ Mot de passe correct")

        if not user.actif:
            print(f"❌ Compte désactivé: {email}")
            return render_template('login.html', warning='Votre compte a été désactivé. Contactez l\'administrateur.')

        # Connexion réussie
        login_user(user, remember=True)
        print(f"✅ Connexion réussie : {user.nom} ({user.email})")

        # Rediriger vers la page demandée ou l'accueil
        next_page = request.args.get('next')
        return redirect(next_page) if next_page else redirect(url_for('index'))

    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    """Déconnexion"""
    print(f"🚪 Déconnexion : {current_user.nom} ({current_user.email})")
    logout_user()
    return redirect(url_for('login'))


# ============================================================
# ROUTES ADMIN - Gestion des utilisateurs
# ============================================================

def admin_required(f):
    """Décorateur pour vérifier que l'utilisateur est admin"""
    @wraps(f)
    @login_required
    def decorated_function(*args, **kwargs):
        if not current_user.is_admin:
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function


@app.route('/admin')
@admin_required
def admin():
    """Page d'administration - Gestion des utilisateurs"""
    users = User.query.order_by(User.date_creation.desc()).all()
    return render_template('admin.html', users=users)


@app.route('/admin/create_user', methods=['POST'])
@admin_required
def admin_create_user():
    """Créer un nouvel utilisateur"""
    email = request.form.get('email', '').strip().lower()
    nom = request.form.get('nom', '').strip()
    password = request.form.get('password', '')
    password_confirm = request.form.get('password_confirm', '')
    is_admin = request.form.get('is_admin') == '1'

    # Validation
    if not email or not nom or not password:
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error='Tous les champs sont requis.')

    if password != password_confirm:
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error='Les mots de passe ne correspondent pas.')

    # Vérifier si l'email existe déjà
    if User.query.filter_by(email=email).first():
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error=f'Un utilisateur avec l\'email {email} existe déjà.')

    # Créer l'utilisateur
    new_user = User(email=email, nom=nom, is_admin=is_admin, actif=True)
    new_user.set_password(password)
    db.session.add(new_user)
    db.session.commit()

    print(f"✅ Nouvel utilisateur créé : {nom} ({email}) - Admin: {is_admin}")
    users = User.query.order_by(User.date_creation.desc()).all()
    return render_template('admin.html', users=users, success=f'Utilisateur {nom} créé avec succès !')


@app.route('/admin/toggle_user', methods=['POST'])
@admin_required
def admin_toggle_user():
    """Activer/Désactiver un utilisateur"""
    user_id = request.form.get('user_id')
    user = User.query.get(user_id)

    if not user:
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error='Utilisateur introuvable.')

    # Protection du compte admin principal
    if user.email == 'johan.mallet@ecogies.fr':
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error='Le compte admin principal ne peut pas être désactivé.')

    # Toggle
    user.actif = not user.actif
    db.session.commit()

    action = 'activé' if user.actif else 'désactivé'
    print(f"🔄 Utilisateur {action} : {user.nom} ({user.email})")
    users = User.query.order_by(User.date_creation.desc()).all()
    return render_template('admin.html', users=users, success=f'Utilisateur {user.nom} {action} avec succès !')


@app.route('/admin/delete_user', methods=['POST'])
@admin_required
def admin_delete_user():
    """Supprimer un utilisateur"""
    user_id = request.form.get('user_id')
    user = User.query.get(user_id)

    if not user:
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error='Utilisateur introuvable.')

    # Protection du compte admin principal
    if user.email == 'johan.mallet@ecogies.fr':
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error='Le compte admin principal ne peut pas être supprimé.')

    # Empêcher la suppression de son propre compte
    if user.id == current_user.id:
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error='Vous ne pouvez pas supprimer votre propre compte.')

    # Supprimer
    nom_user = user.nom
    email_user = user.email
    db.session.delete(user)
    db.session.commit()

    print(f"🗑️ Utilisateur supprimé : {nom_user} ({email_user})")
    users = User.query.order_by(User.date_creation.desc()).all()
    return render_template('admin.html', users=users, success=f'Utilisateur {nom_user} supprimé avec succès !')


@app.route('/admin/reset_password', methods=['POST'])
@admin_required
def admin_reset_password():
    """Réinitialiser le mot de passe d'un utilisateur"""
    user_id = request.form.get('user_id')
    user = User.query.get(user_id)

    if not user:
        users = User.query.order_by(User.date_creation.desc()).all()
        return render_template('admin.html', users=users, error='Utilisateur introuvable.')

    # Générer un mot de passe aléatoire
    new_password = 'CPV' + secrets.token_urlsafe(8)
    user.set_password(new_password)
    db.session.commit()

    print(f"🔑 Mot de passe réinitialisé : {user.nom} ({user.email})")
    users = User.query.order_by(User.date_creation.desc()).all()
    return render_template('admin.html', users=users, success=f'Mot de passe de {user.nom} réinitialisé : {new_password}')


# ============================================================
# INITIALISATION BASE DE DONNÉES ET COMPTE ADMIN
# ============================================================

def init_db():
    """Initialise la base de données et crée le compte admin par défaut"""
    with app.app_context():
        # Créer les tables
        db.create_all()

        # Vérifier si un admin existe
        if not User.query.filter_by(is_admin=True).first():
            # Créer le compte admin par défaut
            admin = User(
                email='johan.mallet@ecogies.fr',
                nom='Johan MALLET',
                is_admin=True,
                actif=True
            )
            admin.set_password('Jaguar2026@')
            db.session.add(admin)
            db.session.commit()
            print('\n' + '='*60)
            print('✅ COMPTE ADMIN CRÉÉ')
            print('='*60)
            print('   Email : johan.mallet@ecogies.fr')
            print('   Mot de passe : Jaguar2026@')
            print('='*60 + '\n')

# Initialiser la base de données au démarrage
init_db()


if __name__ == '__main__':
    print("\n" + "="*60)
    print("🚀 APPLICATION CPV MINT ENERGIE 2026 - VERSION FUSIONNÉE")
    print("="*60)
    print("📍 Accédez à l'application: http://localhost:5001")
    print("\n✨ 3 Workflows disponibles :")
    print("   1️⃣  PDFs uniquement (extraction auto)")
    print("   2️⃣  Excel uniquement (saisie manuelle)")
    print("   3️⃣  PDFs + Excel (RECOMMANDÉ)")
    print("="*60 + "\n")

    # Port dynamique pour Render (ou 5001 en local)
    port = int(os.getenv('PORT', 5001))
    app.run(debug=True, host='0.0.0.0', port=port)
