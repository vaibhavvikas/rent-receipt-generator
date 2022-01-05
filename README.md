# rent-receipt-generator
Generate Rent receipts for the months mentioned


## Libraries required:
    > pip install docx2pdf
    > pip install PyPDF4
    > pip install python-docx

## Other requirements:
    For converting docx to pdf you require Microsoft Office on windows 

## How to use

You only need to change the variables and the script will automatically generate recipts

```
variables = {
        "RENT_MONTH": month,
        "PAYEE_NAME": "Vaibhav Vikas",
        "PAID_AMOUNT": "26,000",
        "HOUSE_ADDRESS": "New Delhi",
        "LANDLORD_NAME": "Some Name",
        "LANDLORD_UPI": "123456789"
    }
```

to configure months for e.g. if you want to generate receipts only for months say Aug-Dec
just make changes in the line no 63 to (8, 12)

Similarly, for years you gotta make changes in the line no 64
