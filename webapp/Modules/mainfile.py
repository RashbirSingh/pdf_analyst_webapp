from Modules.PDFProcessing import PDFProcessing
from Modules.SpacyModel import SpacyModel
from Modules.Highlight import Highlight
import fitz

def mainrun(filepath):
    pp = PDFProcessing(filepath)
    sm = SpacyModel()
    hl = Highlight(filepath)
    result={}
    text = pp.extract_text_from_pdf(way='fitz')
    for i in range(len(text)):
        result.update(sm.labelMaker(text[i]))

    return result



def ok(wordlist, filepath ):
    Pdfdocument = fitz.open(filepath)
    for pagenumber in range(Pdfdocument.pageCount):
        page = Pdfdocument[pagenumber]
        for i in wordlist:
            print(i)
            text_instances = page.searchFor(i, quads=False, hit_max=32)
        for inst in text_instances:
            highlight = page.addHighlightAnnot(inst)
            highlight.setColors(colors='Red')

    Pdfdocument.save('output.pdf', garbage=4, deflate=True, clean=True)

