from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfrw import PdfReader, PdfWriter
from io import StringIO
from tqdm import tqdm
from datetime import datetime


def convert_pdf_to_txt(path, index_of_page):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = {index_of_page}

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text


userInput = int(input("Choose any one option and press ENTER:\n1. Normal Service\n2. Amazon Service\n3. Postal "
                      "Service\n4. Merge without any ID "
                      "checks\n----------------------------------------------------\n--->"))


if userInput not in [1, 2, 3, 4]:
    raise ValueError("Invalid option")


def get_id(text_, type_):
    if type_ == 'label':
        if userInput == 1:
            return text_[text_.find('Shipment ref.') + len('Shipment ref.'):].strip().split('\n')[0]
        elif userInput == 2:
            return text_[text_.find('Parcel ref.') + len('Parcel ref.'):].strip().split(' ')[0]
        elif userInput == 3:
            return " ".join(text_[text_.find('M\nO\nR\nF\n') + len('M\nO\nR\nF\n'):].strip().split('\n')[0].lower().split())
    elif type_ == 'invoice':
        if userInput == 1:
            return text_.split('\n')[1].strip()
        elif userInput == 2:
            return text_.split('\n')[0].split(' ')[1]
        elif userInput == 3:
            return " ".join(text_.split('\n')[0].lower().split())


label = PdfReader('label.pdf')
invoice = PdfReader('invoice.pdf')

result = PdfWriter()

if userInput == 4:
    for i, j in zip(label.pages, invoice.pages):
        result.addpage(i)
        result.addpage(j)
    result.write(f'merged_{datetime.now().strftime("%Y_%m_%d-%I_%M_%p")}.pdf')
    exit()

invoiceKeys = {}
labelKeys = {}

print(f'Extracting {"orderIDs" if userInput != 3 else "Name and Surname"} from Labels:')
for index, page in tqdm(enumerate(label.pages)):
    text = convert_pdf_to_txt('label.pdf', index)
    labelId = get_id(text, 'label')
    labelKeys[labelId] = index

print(f'Extracting {"orderIDs" if userInput != 3 else "Name and Surname"} from Invoices:')
for index, page in tqdm(enumerate(invoice.pages)):
    text = convert_pdf_to_txt('invoice.pdf', index)
    invoiceId = get_id(text, 'invoice')
    invoiceKeys[invoiceId] = index

common = [i for i in labelKeys.keys() if i in invoiceKeys.keys()]

if not len(common) == len(labelKeys):
    print("skipping the following entities due to missing invoices\n----------")
    for i in labelKeys.keys():
        if i not in common:
            print(f'missing: {i}')
    print('----------')
    inp = input("enter y to continue:").lower()
    if inp != 'y':
        exit()
if not len(common) == len(invoiceKeys):
    print("skipping the following entities due to missing labels\n----------")
    for i in invoiceKeys.keys():
        if i not in common:
            print(f'missing: {i}')
    print('----------')
    inp = input("enter y to continue:").lower()
    if inp != 'y':
        exit()

for orderId in common:
    result.addpage(label.pages[labelKeys[orderId]])
    result.addpage(invoice.pages[invoiceKeys[orderId]])

result.write(f'merged_{datetime.now().strftime("%Y_%m_%d-%I_%M_%p")}.pdf')
