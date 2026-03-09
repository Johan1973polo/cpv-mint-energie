from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime

db = SQLAlchemy()

class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    nom = db.Column(db.String(100), nullable=False)
    is_admin = db.Column(db.Boolean, default=False, nullable=False)
    actif = db.Column(db.Boolean, default=True, nullable=False)
    date_creation = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    def set_password(self, password):
        """Hash le mot de passe et le stocke"""
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        """Vérifie si le mot de passe correspond au hash"""
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<User {self.email}>'
