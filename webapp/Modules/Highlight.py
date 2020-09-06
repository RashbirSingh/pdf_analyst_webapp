import fitz
from Modules.HiColors import *

class Highlight:

    def __init__(self, path):
        self.path = path

    def HighlighWords(self,
                      textToSearch,
#                      annotType,
#                      Color,
#                      opacity,
                      outputFileName = "output.pdf"):

        Pdfdocument = fitz.open(self.path)
        for pagenumber in range(Pdfdocument.pageCount):
            page = Pdfdocument[pagenumber]
            text_instances = page.searchFor(textToSearch)
            for inst in text_instances:
                highlight = page.addHighlightAnnot(inst)

#            if annotType == "Highlight":
#                for inst in text_instances:
#                    annot = page.addHighlightAnnot(inst)
#                    annot = self.setHI(annot, Color, opacity)
#            elif annotType == "Underline":
#                for inst in text_instances:
#                    annot = page.addUnderlineAnnot(inst)
#                    annot = self.setHI(annot, Color, opacity)

#            elif annotType == "Rect":
#                for inst in text_instances:
#                    annot = page.addRectAnnot(inst)
#                    annot = self.setHI(annot, Color, opacity)
#            else:
#                for inst in text_instances:
#                    annot = page.addSquigglyAnnot(inst)
#                    annot = self.setHI(annot, Color, opacity)
        Pdfdocument.save(outputFileName, garbage=4, deflate=True, clean=True)

    def annotate(self, file: str, page: object, textToSearch: str, SearchType: str, ColorDict, debug: bool):
        Color = ColorDict[SearchType][0]
        annotType = ColorDict[SearchType][1]
        opacity = ColorDict[SearchType][2]
        HighlightThis = ColorDict[SearchType][3]
        areas = page.searchFor(textToSearch, quads=False, hit_max=32)
        ##TO DO: investigate if Newareas is adding value (or even working)
        newareas = joinAreas(areas, debug)
        AddYPointstoListofLineNumbers(newareas, file, page)

        if HighlightThis:
            if annotType == "Highlight":
                for area in newareas:
                    annot = page.addHighlightAnnot(area)
                    annot = self.setHI(annot, Color, opacity)
            elif annotType == "Underline":
                for area in newareas:
                    annot = page.addUnderlineAnnot(area)
                    annot = self.setHI(annot, Color, opacity)

            elif annotType == "Rect":
                for area in newareas:
                    annot = page.addRectAnnot(area)
                    annot = self.setHI(annot, Color, opacity)
            else:
                for area in newareas:
                    annot = page.addSquigglyAnnot(area)
                    annot = self.setHI(annot, Color, opacity)
        return (page)

    def setHI(annot, Color, opacity):
        annot.setColors(stroke=Color)
        annot.setOpacity(opacity)
        annot.update(fill_color=None)