from flask import Flask, render_template, request
from utils import *

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html', companies=companies, months=months)


@app.route('/generate_doc', methods=['POST'])
def generate_doc():
    form_data = request.form
    company_info = companies[form_data['company']]

    context = prepare_context(form_data, company_info)

    return generate_and_return_doc(context)


if __name__ == '__main__':
    app.run(debug=True)
