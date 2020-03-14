import pytesseract as ocr
from PIL import Image

from pdf2image import convert_from_path, convert_from_bytes

images = convert_from_path(r'C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\F0010 - Data Extraction from PDF files\Example File.pdf')

# ocr.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# TESSDATA_PREFIX = r'C:\Program Files\Tesseract-OCR'

# phrase = ocr.image_to_string(Image.open(r'C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\F0010 - Data Extraction from PDF files\Example File.pdf'))
# print(phrase)