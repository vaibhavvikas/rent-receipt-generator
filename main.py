import calendar
import datetime
import os
from docx2pdf import convert
from glob import glob
from PyPDF4 import PdfFileMerger
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import jinja2

def get_month_period(year, month):
    _, num_days = calendar.monthrange(year, month)
    first_day = datetime.date(year, month, 1)
    last_day = datetime.date(year, month, num_days)
    return first_day.strftime('%d %B, %Y') + " to " + last_day.strftime('%d %B, %Y')

def rent_variables(month):
    tpl = DocxTemplate('rent_receipt_template.docx')
    output_file_path = calendar.month_name[month[0]] + ' ' + str(month[1]) + '.docx'
    
    context = {
        "RENT_MONTH": get_month_period(month[1], month[0]),
        "PAYEE_NAME": "Vaibhav Vikas",
        "PAID_AMOUNT": "26,000",
        "PAYMENT_MODE": "UPI",
        "PROPERTY_ADDRESS": "New Delhi",
        "LANDLORD_NAME": "Some Name",
        "LANDLORD_UPI": "123456789",
        "LANDLORD_PAN": "XXXXXXXXXX",
        "LANDLORD_SIGN": InlineImage(tpl, 'signature.png', height=Mm(20)),
    }

    jinja_env = jinja2.Environment(autoescape=True)
    tpl.render(context, jinja_env)
    tpl.save(output_file_path)

    convert_word_to_pdf(output_file_path)
    os.remove(output_file_path)

def convert_word_to_pdf(input):
    convert(input)

def pdf_merge():
    merger = PdfFileMerger()
    allpdfs = [a for a in glob("*.pdf") if a != "Rent_receipt.pdf"]
    allpdfs.sort(key = lambda x: datetime.datetime.strptime(x[:-4], '%B %Y'))
    [merger.append(pdf) for pdf in allpdfs]
    with open("Rent_receipt.pdf", "wb") as new_file:
        merger.write(new_file)
    merger.close()
    for each in allpdfs:
        os.remove(each)

def main():
    years = [2021, 2022]
    for i in range(11, 14):
        month = [i%12 or i, years[0] if i < 13 else years[1]]
        rent_variables(month)
    pdf_merge()


if __name__ == '__main__':
    main()
