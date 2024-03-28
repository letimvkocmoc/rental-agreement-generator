import json
from docxtpl import DocxTemplate
from datetime import datetime
from flask import send_file

with open('companies.json', 'r', encoding='utf-8') as file:
    companies = json.load(file)

with open('month_translation.json', 'r', encoding='utf-8') as file:
    months = json.load(file)


def num_to_words(num):
    units = ['ноль', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']
    teens = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать']
    tens = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто']
    hundreds = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот', 'девятьсот']

    if num == 0:
        return 'ноль'

    result = ''
    if num // 100 > 0:
        result += hundreds[num // 100] + ' '
        num %= 100

    if num >= 20:
        result += tens[num // 10] + ' '
        num %= 10

    if 10 <= num < 20:
        result += teens[num - 10]
    elif num > 0:
        result += units[num]

    return result


def get_in_case_post(post):
    post_lower = post.lower()
    if post_lower.startswith("генеральный директор"):
        return "генерального директора"
    elif post_lower.startswith("директор"):
        return "директора"
    elif post_lower.startswith("исполняющий обязанности генерального директора"):
        return "исполняющего обязанности генерального директора"
    elif post_lower.startswith("исполняющий обязанности директора"):
        return "исполняющего обязанности директора"
    else:
        return post_lower


def get_in_case_doc(doc):
    doc_lower = doc.lower()
    if doc_lower.startswith("устав"):
        return "Устава"


def replace_quotes(text):
    parts = text.split('"')
    for i in range(1, len(parts), 2):
        parts[i] = '«' + parts[i] + '»'
    return ''.join(parts)


def prepare_context(form_data, company_info):
    docDate = datetime.strptime(form_data['docDate'], '%Y-%m-%d')
    rightsRegistrationDocIssuedDate = datetime.strptime(form_data['rightsRegistrationDocIssuedDate'], '%Y-%m-%d')
    rentalStartDate = datetime.strptime(form_data['rentalStartDate'], '%Y-%m-%d')

    orgLandlordShortname = get_shortname(company_info['orgLandlordFullname'])
    orgTenantShortname = get_shortname(form_data['orgTenantFullname'])
    orgTenant = replace_quotes(form_data['orgTenant'])

    context = {
        **company_info,
        'docID': form_data['docID'],
        'docDate': format_date(docDate),
        'orgLandlordPostInCase': get_in_case_post(company_info['orgLandlordPost']),
        'orgLandlordDocInCase': get_in_case_doc(company_info['orgLandlordDoc']),
        'orgTenant': orgTenant,
        'orgTenantPost': form_data['orgTenantPost'],
        'orgTenantDocInCase': get_in_case_doc(form_data['orgTenantDoc']),
        'orgTenantPostInCase': get_in_case_post(form_data['orgTenantPost']),
        'rentalObject': form_data['rentalObject'],
        'activityType': form_data['activityType'],
        'rightsRegistrationDoc': form_data['rightsRegistrationDoc'],
        'rightsRegistrationDocSeriesAndNumber': form_data['rightsRegistrationDocSeriesAndNumber'],
        'rightsRegistrationDocIssuedDate': format_date(rightsRegistrationDocIssuedDate),
        'orgTenantRegAddress': form_data['orgTenantRegAddress'],
        'orgTenantINN': form_data['orgTenantINN'],
        'orgTenantKPP': form_data['orgTenantKPP'],
        'orgTenantBank': form_data['orgTenantBank'],
        'orgTenantChA': form_data['orgTenantChA'],
        'orgTenantCorrA': form_data['orgTenantCorrA'],
        'orgTenantBIK': form_data['orgTenantBIK'],
        'orgTenantOGRN': form_data['orgTenantOGRN'],
        'orgTenantTel': form_data['orgTenantTel'],
        'orgTenantEmail': form_data['orgTenantEmail'],
        'rentalPrice': form_data['rentalPrice'],
        'orgTenantFullname': form_data['orgTenantFullname'],
        'orgLandlordShortname': orgLandlordShortname,
        'orgTenantShortname': orgTenantShortname,
        'validityTime': f"{form_data['validityTime']} ({num_to_words(int(form_data['validityTime']))}) {'год' if int(form_data['validityTime']) == 1 else 'года' if form_data['periodUnit'] == 'год' else 'лет'}",
        'rentalStartDate': format_date(rentalStartDate)}

    return context


def get_shortname(fullname):
    return ' '.join(
        [part if idx == 0 or len(part) <= 1 else f'{part[0]}.' for idx, part in enumerate(fullname.split())])


def format_date(date):
    return f"«{str(date.day).zfill(2)}» {months[date.strftime('%B')]} {date.year} г."


def generate_and_return_doc(context):
    doc = DocxTemplate("template.docx")
    doc.render(context)
    doc.save('output.docx')

    return send_file('output.docx', as_attachment=True)
