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
    tpl = DocxTemplate('template/rent_receipt_template.docx')
    output_file_path = "output/" + calendar.month_name[month[0]] + ' ' + str(month[1]) + '.docx'

    context = {
        "RENT_MONTH": get_month_period(month[1], month[0]),
        "PAYEE_NAME": "Vaibhav Vikas",
        "PAID_AMOUNT": "26,000",
        "PAYMENT_MODE": "UPI",
        "PROPERTY_ADDRESS": "New Delhi",
        "LANDLORD_NAME": "First Last",
        "LANDLORD_PAN": "XXXXXXXXXX",
        "LANDLORD_SIGN": InlineImage(tpl, 'signature.png', height=Mm(8)),
    }

    jinja_env = jinja2.Environment(autoescape=True)
    tpl.render(context, jinja_env)
    tpl.save(output_file_path)

    convert(output_file_path)
    os.remove(output_file_path)


def pdf_merge():
    merger = PdfFileMerger()
    allpdfs = [a for a in glob("output/*.pdf") if not a.endswith("Rent_receipt.pdf")]
    allpdfs.sort(key=lambda x: datetime.datetime.strptime((x[7:]).rstrip(".pdf"), '%B %Y'))
    [merger.append(pdf) for pdf in allpdfs]
    with open("output/Rent_receipt.pdf", "wb") as new_file:
        merger.write(new_file)
    merger.close()
    for each in allpdfs:
        os.remove(each)


def main():
    years = [2021, 2022]
    for i in range(7, 14):
        month = [i % 12 or i, years[0] if i < 13 else years[1]]
        rent_variables(month)
    pdf_merge()


if __name__ == '__main__':
    main()
