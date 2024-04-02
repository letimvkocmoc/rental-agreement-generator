import os

from flask import Flask, render_template, request, session, abort, redirect, url_for
from utils import *
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

app.secret_key = os.getenv('SECRET_KEY')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        env_username = os.getenv('FLASK_USERNAME')
        env_password = os.getenv('PASSWORD')

        if username == str(env_username) and password == str(env_password):
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            return 'Неверный логин или пароль!'

    return render_template('login.html')


@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    return render_template('index.html', companies=companies, months=months)


@app.route('/generate_doc', methods=['POST'])
def generate_doc():
    if not session.get('logged_in'):
        abort(401)

    form_data = request.form
    company_info = companies[form_data['company']]

    context = prepare_context(form_data, company_info)

    return generate_and_return_doc(context)


if __name__ == '__main__':
    app.run(debug=True)
