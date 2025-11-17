from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_user, logout_user
from models.user import User

auth_bp = Blueprint('auth', __name__)


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username")
        senha = request.form.get("password")

        user = User.query.filter_by(username=username).first()

        if not user or not user.check_password(senha):
            flash("Usuário ou senha incorretos!", "danger")
            return render_template("login.html")

        login_user(user)
        return redirect(url_for("dashboard.dashboard"))

    return render_template("login.html")


@auth_bp.route("/logout")
def logout():
    logout_user()
    return redirect(url_for("auth.login"))
