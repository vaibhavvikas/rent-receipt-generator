# rent-receipt-generator
Generate Rent receipts for the months mentioned


## Libraries required:
All the libraries are mentioned in the requirements.txt file. You can install all the libraries by
running the command

    > pip install -r requirements.txt


## Other requirements:

For converting docx to pdf you require Microsoft Office on windows 

## How to use

You only need to change the variables and the script will automatically generate recipts

```
   context = {
        "RENT_MONTH": get_month_period(month[1], month[0]),
        "PAYEE_NAME": "Vaibhav Vikas",
        "PAID_AMOUNT": "26,000",
        "PAYMENT_MODE": "UPI",
        "PROPERTY_ADDRESS": "New Delhi",
        "LANDLORD_NAME": "Some Name",
        "LANDLORD_PAN": "XXXXXXXXXX",
        "LANDLORD_SIGN": InlineImage(tpl, 'signature.png', height=Mm(20)),
    }
```

to configure months for e.g. if you want to generate receipts only for months say Aug-Dec
just make changes in the line no 57 to (8, 12)

Similarly, for years you gotta make changes in the line no 56
