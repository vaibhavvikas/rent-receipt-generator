from docx import Document
import calendar
from datetime import datetime
import os, time
from docx2pdf import convert
from glob import glob
from PyPDF4 import PdfFileMerger


def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)


def rent_variables(month):
    template_file_path = 'rent_receipt_template.docx'
    output_file_path = month + '.docx'
    
    variables = {
        "RENT_MONTH": month,
        "PAYEE_NAME": "Vaibhav Vikas",
        "PAID_AMOUNT": "26,000",
        "HOUSE_ADDRESS": "New Delhi",
        "LANDLORD_NAME": "Some Name",
        "LANDLORD_UPI": "123456789"
    }

    template_document = Document(template_file_path)
    for variable_key, variable_value in variables.items():
        for paragraph in template_document.paragraphs:
            replace_text_in_paragraph(paragraph, variable_key, variable_value)

        for table in template_document.tables:
            for col in table.columns:
                for cell in col.cells:
                    for paragraph in cell.paragraphs:
                        replace_text_in_paragraph(paragraph, variable_key, variable_value)

    template_document.save(output_file_path)
    convert_word_to_pdf(output_file_path)
    os.remove(output_file_path)

def convert_word_to_pdf(input):
    convert(input)

def pdf_merge():
    merger = PdfFileMerger()
    allpdfs = [a for a in glob("*.pdf") if a != "Rent_receipt.pdf"]
    allpdfs.sort(key = lambda x: datetime.strptime(x[:-4], '%B %Y'))
    [merger.append(pdf) for pdf in allpdfs]
    with open("Rent_receipt.pdf", "wb") as new_file:
        merger.write(new_file)
    merger.close()
    for each in allpdfs:
        os.remove(each)

def main():
    for i in range(7, 14):
        month = str(calendar.month_name[i%12 or 12]) + str(" 2021" if i < 13 else " 2022")
        rent_variables(month)
    pdf_merge()


if __name__ == '__main__':
    main()
