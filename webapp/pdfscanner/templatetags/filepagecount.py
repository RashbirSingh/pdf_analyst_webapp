import os
from PyPDF2 import PdfFileReader
from django import template
register = template.Library()

@register.filter
def filepagecount(value):
    pdf = PdfFileReader(open(value.file.name, 'rb'))
    return pdf.getNumPages()