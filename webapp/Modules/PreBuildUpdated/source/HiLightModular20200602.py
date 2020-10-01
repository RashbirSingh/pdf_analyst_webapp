
from __future__ import unicode_literals, print_function
import json
import fitz
import regex
import inspect
import spacy
import xlsxwriter
import datetime
import dateparser
import os
import re
from pprint import pprint

#Need to update this to link to the correct colorsDict
from Modules.PreBuildUpdated.source.HiColors import *
from Modules.PreBuildUpdated.source.HiRegexes import *

from typing import List
from mergedeep import merge, Strategy

import recognizers_suite as Recognizers
from recognizers_suite import Culture, ModelResult

from celery import shared_task
from celery_progress.backend import ProgressRecorder

now = datetime.datetime.now()
nlp = spacy.load("en_blackstone_proto")
nlp2 = spacy.load("en_core_web_sm")
# nlp = spacy.load("en_core_web_sm")
# nlp = en_core_web_sm.load()

## Preferences
SpaCyList = ["ORG", "TIME", "PERSON", "LAW", "INSTRUMENT", "PROVISION", "COURT", "MONEY", "JUDGE"]
DEFAULT_CULTURE = Culture.English
LOCALES = ['en-AU']
debug = False
verbose = 0
Y0 = {}
Y1 = {}
LineNumbers = {}


pageCounter = 0

#work out average width per 0##work out
averagecharactersize = 10

pattern, PageDict, DocDict, SentenceDict, regexGenericPageDict, regexGenericSentenceDict, InstancesDict, annotinfo = {}, {}, {}, {}, {}, {}, {}, {}
patternRegexSimpleDate = ""
sentences, areas = [], []
LabelsList = list([])
for key in InvColorDictLabelstoColors:
    LabelsList.append(key)
LabelsList = list(set(LabelsList))

annotType = "Highlight"
DEFAULT_CULTURE = Culture.English
Y0 = {}
Y1 = {}
LineNumbers = {}


patternRegexSimpleDate = regex.compile(regexSimpleDate)

for key in regexGeneric:
    regexGeneric[key] = regex.compile(regexGeneric[key],regex.MULTILINE)

@shared_task(bind=True)
def Highlight_Analyse(self, lst, ColorDict, savePDF, saveExcel, saveExcelUVO, label, debug):
    global LineNumbers
    import time
    d = {}
    DocDict = {}
    DocDictList = {}
    textPagesDict = {}
    textSentencesDict = {}
    timestr = time.strftime("%Y%m%d-%H%M%S")
    progress_recorder = ProgressRecorder(self)

    i = 0

    for input_pdf in lst:

        filename = "".join(os.path.split(input_pdf)[1])
        progress_recorder.set_progress(i, len(lst), f'processing PDF {input_pdf.split("/")[-1]}')
        i = i + 1

        pdl = PDF2DictList(input_pdf)
        pdl2 = PDF2DictList2(input_pdf)
        pdl2[1].update(pdl[1])
        pdl2[0][0].update(pdl[0][0])
        for k, v in list(pdl2[0][0].items()):
            if (k == "INSTRUMENT") or (k == "PROVISION"):
                if "LAW" in pdl2[0][0].keys():
                    pdl2[0][0][re.sub("the ", "", [k])] = pdl2[0][0].pop(k)
                    k = re.sub("the ", "", [k])
                    pdl2[0][0]["LAW"].update(pdl2[0][0][k])
                    pdl2[0][0].pop(k)
                    for key, v in pdl2[0][0]["LAW"].items():
                        pdl2[0][0]["LAW"][re.sub('the ', "", key.lower())] = pdl2[0][0]["LAW"].pop(key)
                        print(pdl2[0][0]["LAW"][key])

                else:
                    pdl2[0][0]["LAW"] = pdl2[0][0][k]
                    pdl2[0][0].pop(k)
                    for key, v in pdl2[0][0]["LAW"].items():
                        pdl2[0][0]["LAW"][re.sub('the ', "", key.lower())] = pdl2[0][0]["LAW"].pop(key)
                        print(pdl2[0][0]["LAW"][key])
            # if (k == "NUMBER"):
            #     for Numberkey, Numbervalue in pdl2[0][0]["NUMBER"].items():
            #         print(Numberkey)
            #         pdl2[0][0]["NUMBER"][" " + Numberkey + " "] = pdl2[0][0]["NUMBER"].pop(Numberkey)
        DocDictList[input_pdf] = pdl2[0]
        textSentencesDict[input_pdf] = pdl2[1]
        DocDict[input_pdf] = MergeList2Dict(DocDictList[input_pdf], debug)
        d = merge(d, DocDict[input_pdf], strategy=Strategy.ADDITIVE)

        if savePDF:
            
            doc = markup(input_pdf,DocDictList[input_pdf], ColorDict, debug)
            saveDocToPDF(doc,input_pdf,"output", timestr, debug)

    if saveExcel:

        saveDocDict2Excel(d, lst[0], timestr, debug)

    if saveExcelUVO:

        saveDict2ExcelUniqueValsOnly(d, lst[0], timestr, debug)

    return [d, DocDict, DocDictList, textSentencesDict, LineNumbers]

def isAcronym(n:str):
    if len(n)>3 and all(i.isupper() for i in n):
        return True
    else:
        return False

def isLegitName(n:str):
    if debug: print("n: ",n)
    if any(char.isdigit() for char in n):
        if debug: print("test1 failed")
        return False
    elif " -" in n:
        if debug: print("test2 failed")
        return False
    elif " -" in n:
        if debug: print("test3 failed")
        return False
    elif "Mm" in n: 
        if debug: print("test4 failed")
        return False
    elif any(n2[0].islower() for n2 in n.split()):
        if debug: print("test5 failed")
        return False
    elif all(isAcronym(n2) for n2 in n.split()):
        if debug: print("test6 failed")
        return False       
    else: return True

def PDF2DictList(input_pdf):
    global SpaCyList, LabelsList, regexGeneric, regexSimpleDate
    filename = "".join(os.path.split(input_pdf)[1])
    print(filename)
    doc = fitz.open(input_pdf)
    pageCounter = 0
    DocDictList=[]
    textPages = []
    textSentencesDict = {}
    for page in doc:
        pageCounter = pageCounter + 1

        sentenceCounter = 0

        text = page.getText("")
        textPages.append(text)
        PageDict = {}

        doc1 = nlp(text)

        textSentencesDict[pageCounter-1] = []

        for sentence in doc1.sents:
            textSentencesDict[pageCounter-1].append(str(sentence))
            SentenceDict = {}
            sentenceCounter = sentenceCounter + 1
            regexGenericSentenceDict = {}
            ents = sentence.ents
            EntsDict = (dict([(x.text, x.label_) for x in ents]))
            for k, v in EntsDict.items():
                if v in LabelsList and v in SpaCyList:
                    SentenceDict[v] = SentenceDict.get(v, [])

                    k.replace('\n','').replace('\u00b7','')

                    if v == "PERSON":
                        if isLegitName(k):
                            k.strip("'s")
                            k.strip("'")
                    newDict = {k:[(filename, pageCounter, sentenceCounter, str(sentence), input_pdf)]}

                    SentenceDict[v]=newDict

            for SearchType in regexGeneric:
                newDict={}
                matchlist = list(set(run_recognition(str(sentence), SearchType, [])))
                if matchlist != []:
                    regexGenericSentenceDict[SearchType] = regexGenericSentenceDict.get(SearchType, [])
                    for match in matchlist:
                        match.replace('\n','').replace('\u00b7','')
                        
                        newDict[match] = [(filename,pageCounter, sentenceCounter, str(sentence), input_pdf)]

                    regexGenericSentenceDict[SearchType] = newDict

            merge(SentenceDict, regexGenericSentenceDict, strategy=Strategy.ADDITIVE)
            merge(PageDict, SentenceDict, strategy=Strategy.ADDITIVE)

        DocDictList.append(PageDict)       

    return (DocDictList, textSentencesDict)


def PDF2DictList2(input_pdf):
    global SpaCyList, LabelsList, regexGeneric, regexSimpleDate
    filename = "".join(os.path.split(input_pdf)[1])
    print(filename)
    doc = fitz.open(input_pdf)
    pageCounter = 0
    DocDictList = []
    textPages = []
    textSentencesDict = {}
    for page in doc:
        pageCounter = pageCounter + 1
        sentenceCounter = 0
        text = page.getText("")
        textPages.append(text)
        PageDict = {}
        doc1 = nlp2(text)

        textSentencesDict[pageCounter - 1] = []

        for sentence in doc1.sents:
            textSentencesDict[pageCounter - 1].append(str(sentence))
            SentenceDict = {}
            sentenceCounter = sentenceCounter + 1
            regexGenericSentenceDict = {}
            ents = sentence.ents
            EntsDict = (dict([(x.text, x.label_) for x in ents]))
            for k, v in EntsDict.items():
                if v in LabelsList and v in SpaCyList:
                    SentenceDict[v] = SentenceDict.get(v, [])

                    k.replace('\n', '').replace('\u00b7', '')

                    if v == "PERSON":
                        if isLegitName(k):
                            k.strip("'s")
                            k.strip("'")
                    newDict = {k: [(filename, pageCounter, sentenceCounter, str(sentence), input_pdf)]}

                    SentenceDict[v] = newDict

            for SearchType in regexGeneric:
                newDict = {}
                matchlist = list(set(run_recognition(str(sentence), SearchType, [])))
                if matchlist != []:
                    regexGenericSentenceDict[SearchType] = regexGenericSentenceDict.get(SearchType, [])
                    for match in matchlist:
                        match.replace('\n', '').replace('\u00b7', '')

                        newDict[match] = [(filename, pageCounter, sentenceCounter, str(sentence), input_pdf)]

                    regexGenericSentenceDict[SearchType] = newDict

            merge(SentenceDict, regexGenericSentenceDict, strategy=Strategy.ADDITIVE)
            merge(PageDict, SentenceDict, strategy=Strategy.ADDITIVE)

        DocDictList.append(PageDict)

    return (DocDictList, textSentencesDict)


def run_recognition(string: str, SearchType:str, matchlist:list):

    while string != "":
        matchobj = regex.search(regexGeneric[SearchType], string)
        if matchobj:
            if not SearchType == "WORD":
                matchlist.append(matchobj.group(0).strip("\n,. "))
            else:
                if not matchobj.group(0)[:-1] == "": matchlist.append(matchobj.group(0)[:-1])
            string = string[matchobj.end():]
        else:
            break
    return matchlist

def MergeList2Dict(DictList: list, debug: bool):
    DocDict = {}
    for d in DictList:
        merge(DocDict, d, strategy=Strategy.ADDITIVE)
    return DocDict

def markup(input_pdf, DocDictListInstance, ColorDict, debug):
    Y0[input_pdf]={}
    Y1[input_pdf]={}
    LineNumbers[input_pdf] = {}
    pageCounter = 0

    doc = fitz.open(input_pdf)
    for page in doc:

        Y0[input_pdf][page]=set()
        Y1[input_pdf][page]=set()
        LineNumbers[input_pdf][page] = {}
        
        pageCounter = pageCounter + 1
        sentenceCounter = 0
        if debug: print(pageCounter)
        text = page.getText("")

        for key, value in DocDictListInstance[pageCounter-1].items():


            for k2 in value:
                SearchText = k2
                # TODO: Manage Number overlapping more accurately
                if SearchText.isdigit():
                    SearchText = " " + SearchText + " "
                for lst in value[k2]:

                    if lst[1] == pageCounter:

                        if ColorDict[key][3]: page=annotate(input_pdf, page, SearchText, key, "", ColorDict, debug)

        LineNumbers[input_pdf][page] = CompileListofLineNumbers(Y0, Y1, input_pdf, page)
    return(doc)

def setHI(annot, Color, opacity):
    annot.setColors(stroke=Color)
    annot.setOpacity(opacity)
    annot.update(fill_color=None)

def CompileListofLineNumbers(Y0, Y1, input_pdf, page):
    newdict={}
    try:
        for i in range(len(sorted(Y0[input_pdf][page]))):
            newdict[i] = (sorted(Y0[input_pdf][page])[i], sorted(Y1[input_pdf][page])[i])
    except:
        0
    return newdict
    
def AddYPointstoListofLineNumbers(areas, file, page):
    for area in areas:
        try:
            Y0[file][page].add(area.y0)
            Y1[file][page].add(area.y1)
        except:
            0

def annotate(file:str, page: object, SearchText: str, SearchType:str, sentence:str, ColorDict, debug:bool):

    Color = ColorDict[SearchType][0]
    annotType = ColorDict[SearchType][1]
    opacity = ColorDict[SearchType][2]
    HighlightThis = ColorDict[SearchType][3]
    # SearchText = " " + SearchText + " "
    areas = page.searchFor(SearchText, quads=False, hit_max = 32)
##TODO: investigate if Newareas is adding value (or even working)
    newareas = joinAreas(areas, debug)
    AddYPointstoListofLineNumbers(newareas, file, page)
    
    if HighlightThis:
        if annotType == "Highlight":
            for area in newareas:
                try:
                    annot = page.addHighlightAnnot(area)
                    annot = setHI(annot, Color, opacity)
                except:
                    ##if debug: print("failed")
                    0
        elif annotType == "Underline":
            for area in newareas:
                try:
                    annot = page.addUnderlineAnnot(area)
                    annot = setHI(annot, Color, opacity)
                except:
                    ##if debug: print("failed")
                    0

        elif annotType == "Rect":
            for area in newareas:
                try:
                    annot = page.addRectAnnot(area)
                    annot = setHI(annot, Color, opacity)
                except:
                    ##if debug: print("failed")
                    0
        else:
            for area in newareas:
                try:
                    annot = page.addSquigglyAnnot(area)
                    annot = setHI(annot, Color, opacity)
                except:
                    ##if debug: print("failed")
                    0
    return(page)      

def isAmbiguous(date_time):
    if int(date_time[8:9]) < 13 and int(date_time[5:6]) < 13:
        return True
    else:
        return False
    
def reverseDayandMonth(date1_time): 
##    return date1_time[0:4]+date1_time[8:9]+date1_time[4:6]
    return date1_time

## TODO use this in the front end
def guessCentury(twodigyr: str):
    if 2000 + int(twodigyr) - now.year < 10:
        return str(2000+int(twodigyr))
    else:
        return str(1900 + int(twodigyr))
    
def saveDocDict2Excel(DocDict:dict, input_pdf:str, timestr:str, file_sel_list:list, exportHyperlinks: bool, debug: bool):

    # workbook = xlsxwriter.Workbook(input_pdf+timestr+'DATA.xlsx',   )
    workbook = xlsxwriter.Workbook('DATA.xlsx')
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'valign': 'top'})
    date_time_format = workbook.add_format({'num_format': 'ddd d mmm yyyy hh:mm:ss', 'valign': 'top'})
    bold = workbook.add_format({'bold': True, 'valign': 'top'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    # TO DO: Add a number format for cells with money, e.g.  money_format = workbook.add_format({'num_format': '$#,##0'})
##   Probably best in a new column of parsed text ie stripped of all leading non-numerical characters.

    introText = "On each tab (sheet) of this workbook are listed instances of terms found in the file(s) analysed. \n \
Each category of terms is set out on a separate sheet of this workbook (see TABS below). \n \
For each term, each time that term was found in the file(s) analysed is indicated by a separate entry. \n \
For persons and organisations, the full name and partial matches are listed. \n \
Note that partial matches may be listed under more than one full name. \n \
For dates, the text found is given together with the automatically interpreted date or date range. \n \
Automatic interpretation can be unreliable for partial dates such as day and month or month only \n \
where no year is given. \n \
Some terms will be found in multiple categories, eg. ADDRESS and ORG or MONEY and NUMBER. \n \
All unique terms are listed in WORD, firstly capitalised terms then non-capitalised terms. "
    
    addIntroSheet(workbook, introText, bold, top, wrap, timestr, date_time_format, file_sel_list,exportHyperlinks )

    labelDict = {"Term": 25, "File":25, "Page": 7, "Sent.": 7, "Context": 80, "Hyperlink": 50}

    a = DocDict
    for key, value in sorted(DocDict.items()):
        row, col = 0,0

        URL_n = 0
    
##        cell_format.set_bold()
        worksheet = workbook.add_worksheet(name=key)
        if key == "PERSON" or key == "ORG":
            worksheet.write(row, col, 'Full Name', bold)
            worksheet.set_column(col, col, 30)
            col += 1
        if key == "DATE":
            worksheet.write(row, col, 'Date', bold)
            worksheet.set_column(col, col, 10)
            col += 1
            worksheet.write(row, col, 'End Date', bold)
            worksheet.set_column(col, col, 10)
            col += 1
        for label, width in labelDict.items():
            worksheet.write(row, col, label, bold)
            worksheet.set_column(col, col, width)
            col += 1

        row += 1

        for key2 in sorted(DocDict[key].keys()):

            if key == "PERSON" or key == "ORG":
                for key3 in sorted(DocDict[key][key2]):
                    for i in DocDict[key][key2][key3]:

                        col = 0
                        worksheet.write(row, col, key2, wrap)
                        col += 1
                        worksheet.write(row, col, key3, wrap)
                        col += 1
                        for j in i[
                                 :-1]:  # need to slice off last element which is the full file path (we don't export that)
                            worksheet.write(row, col, j, wrap)
                            col += 1
                        if exportHyperlinks == True:
                            if URL_n < 65529:
                                worksheet.write_url(row, col, i[-1] + "#page=" + str(i[-4]))
                            ##                              exception needed in case Excel reaches its limit of URLs on a single page (approx. 65K).
                            else:
                                worksheet.write(row, col, i[-1] + "#page=" + str(i[-4]))
                            URL_n += 1
                        row += 1


            else: # if key != "PERSON" or "ORG": 
                for i in DocDict[key][key2]:

                    col = 0

                    if key == "DATE":
                        fiddled = False
##                        d = eval(MS_Recognize(key))
                        ##if debug: print (key2)
                        ##if debug: print(MS_Recognize(key2))
                        if key2[:1] == "'":
                            key2 = guessCentury(key2[1:]) 
#                            key2 = "20"+key2[1:]
                            fiddled = True
                        d = MS_Recognize(key2)
                        dp = str(dateparser.parse(key2, locales=LOCALES))[0:10]
                        if fiddled == True:
                            key2 = "'"+key2[2:]
                        if debug: print(d)
                        if debug: print(dp)
                        if not d == None:
                            try:
                                d = eval(d)

                                ##if debug: print (d["values"])
                                ##if debug: print(d["values"][0])
                                if d["values"][0]["type"] == "daterange":
                                    date1 = d["values"][0]["start"]
                                    date1_time = datetime.datetime.strptime(date1, '%Y-%m-%d')
                                    worksheet.write_datetime(row,col, date1_time, date_format)
            ##                              Only the first date in a daterange is supported
                                    col += 1

                                    date2 = d["values"][0]["end"]
                                    date2_time = datetime.datetime.strptime(date2, '%Y-%m-%d')
                                    worksheet.write_datetime(row,col, date2_time, date_format)

                                    col +=1
                                    
                                elif d["values"][0]["type"] == "date":
    ##                                date1 = d["values"][0]["value"]
    ##                                date1_time = datetime.datetime.strptime(date1, '%Y-%m-%d')
    ##                                worksheet.write_datetime(row,col, date1_time, date_format)
    ## TODO try using this
                                    date1_time = datetime.datetime.strptime(dp, '%Y-%m-%d')
                                    worksheet.write_datetime(row,col, date1_time, date_format)

                                    col+=2
                                ##                        worksheet.write(row, col, date1_time)
                            except:
                                0
##                                if debug: print("518")
                    
                    worksheet.write(row,col, key2, wrap)        
##                    worksheet.write(row, col, key2)
                    col += 1
                    for j in i[:-1]: #last element is the full file path (used in building links))
    ##                    for j in i[:-1]: #need to slice off last element which is the full file path (we don't export that)
                        worksheet.write(row, col, j, wrap)
                        col += 1
                    if exportHyperlinks == True:
                        if URL_n < 65529:
                            worksheet.write_url(row, col, i[-1]+"#page="+str(i[-4]))
##                              exception needed in case Excel reaches its limit of URLs on a single page (approx. 65K).
                        else:
                            worksheet.write(row, col, i[-1]+"#page="+str(i[-4]))
                        URL_n += 1

                    row += 1
        worksheet.autofilter(0, 0, row-1, col)
    workbook.close()

def addIntroSheet(workbook: object, introText: str, bold, top, wrap, timestr, date_time_format, file_sel_list,exportHyperlinks):

    if True:
        
        worksheet = workbook.add_worksheet(name="AA_README")
        worksheet.set_column(0,0, 25)
        worksheet.set_column(1,1, 85)
        row, col = 0, 0
        worksheet.write(row, col, "Introduction", bold)
        col+=1
        worksheet.write(row, col, introText, wrap)

        row += 1
        col = 0
        worksheet.write(row, col, 'Date and time of analysis: ', bold)
        date1_time = datetime.datetime.strptime(timestr, '%Y%m%d-%H%M%S')
        col+=1
        worksheet.write(row, col, date1_time, date_time_format)
        col = 0
        row += 1
        worksheet.write(row, col, 'List of files analysed:', bold)
        col += 1
##        if True:
        if exportHyperlinks == True:
            for i in sorted(file_sel_list):
                worksheet.write_url(row, col, i)
                row +=1
        else:
            for i in sorted(file_sel_list):
                worksheet.write(row, col, i, top) 
                row +=1
            
        col = 0
        worksheet.write(row, col, 'Credits: ', bold)
        col+=1
        worksheet.write(row,col, "Files analysed by PDFAnalyst (c) 2020 by ComputerBrain Software\n \
No responsibility is taken for incorrect interpretation or categorisation of terms.", wrap)
        

def saveDict2ExcelUniqueValsOnly(DocDict:dict, input_pdf:str, timestr:str, file_sel_list:list, exportHyperlinks:bool, debug:bool):

    workbook = xlsxwriter.Workbook('DATAUVO.xlsx')
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
    date_time_format = workbook.add_format({'num_format': 'ddd d mmm yyyy hh:mm:ss'})
    bold = workbook.add_format({'bold': True, 'valign': 'top'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    cell_format = workbook.add_format()
    cell_format.set_text_wrap()

#    money = workbook.add_format({'num_format': '$#,##0'})
    introText = "On each tab (sheet) of this workbook are listed unique instances of terms found in the file(s) analysed. \n \
Each category of terms is set out on a separate sheet of this workbook (see TABS below). \n \
For each term, the number of times that term was found in the file(s) analysed is indicated. \n \
For persons and organisations, the full name and partial matches are listed. \n \
Note that partial matches may be listed under more than one full name. \n \
For dates, the text found is given together with the automatically interpreted date. \n \
Automatic interpretation can be unreliable for partial dates such as day and month or month only where \n \
no year is given. \n \
Some terms will be in multiple categories, eg. ADDRESS and ORG or MONEY and NUMBER. \n \
All unique terms are listed in WORD, firstly capitalised terms then non-capitalised terms. "
    
    addIntroSheet(workbook, introText, bold, top, wrap, timestr, date_time_format, file_sel_list,exportHyperlinks )
     
    for key in sorted(DocDict):

        worksheet = workbook.add_worksheet(name=key)
        if key == "DATE":
            worksheet.write('A1', 'Term', bold)
            worksheet.set_column(0,0, 20)                                
            worksheet.write('B1', 'Date', bold)
            worksheet.set_column('B:B', 10)
            worksheet.write('C1', 'EndDate', bold)
            worksheet.set_column('C:C', 10)
            worksheet.write('D1', 'Instances', bold)
            worksheet.set_column('D:D', 5)

        elif key == "PERSON" or key == "ORG":
            worksheet.write('A1', 'Full name', bold)
            worksheet.set_column(0,0, 30) 
            worksheet.write('B1', 'Term', bold)
            worksheet.set_column('B:B', 30)
            worksheet.write('C1', 'Instances', bold)
            worksheet.set_column('C:C', 5)

        else:
            worksheet.write('A1', 'Term', bold)
            worksheet.set_column(0,0, 30)    
            worksheet.write('B1', 'Instances', bold)
            worksheet.set_column('B:B', 5)

        row = 1

      
        for key2 in sorted(DocDict[key].keys()):
                
            if key == "PERSON" or key == "ORG":
                for key3 in sorted(DocDict[key][key2]):
                    col = 0
                    worksheet.write(row, col, key2, wrap)
                    col += 1
                    worksheet.write(row, col, key2, wrap)
                    col += 1
                    worksheet.write(row, col, len(DocDict[key][key2]))
                    
                    row += 1
                        
            elif key == "DATE":
                fiddled = False
                col = 0
                worksheet.write(row, col, key2, wrap)
                col+=1
                if key2[:1] == "'":
                    key2 = guessCentury(key2[1:]) 
##                    key2 = "20"+key2[1:]
                    fiddled = True
                d = MS_Recognize(key2)
                dp = str(dateparser.parse(key2, locales=LOCALES))[0:10]
                if fiddled == True:
                    key2 = "'"+key2[2:]
                if not d == None:
                    try:
                        d = eval(d)
                        if d["values"][0]["type"] == "daterange":
                            date1 = d["values"][0]["start"]
                            date1_time = datetime.datetime.strptime(date1, '%Y-%m-%d')
                            worksheet.write_datetime(row,col, date1_time, date_format)
                            date2 = d["values"][0]["end"]
                            date2_time = datetime.datetime.strptime(date2, '%Y-%m-%d')
                            worksheet.write_datetime(row,col+1, date2_time, date_format)
                        elif d["values"][0]["type"] == "date":
                            ##if debug: print ("425")

                            date1_time = str(dp).strip()
                            ##if debug: print("date1_time: ",date1_time)
                            date1_time = datetime.datetime.strptime(dp, '%Y-%m-%d')
                            worksheet.write_datetime(row,col, date1_time, date_format)
                    except:
                        0
                col = 3
                worksheet.write(row, col, len(DocDict[key][key2]))
                row += 1
            else:
##                elif: key != "DATE" and key != "PERSON" and key != "ORG":
                col = 0
                worksheet.write(row, col, key2, wrap)
                col +=1
                worksheet.write(row, col, len(DocDict[key][key2]))
                
                row += 1
        worksheet.autofilter(0, 0, row-1, col) 
    workbook.close()

def close (area0, area1):
##    ##if debug: print ("area0, area1: ", area0, area1)
    if area1.y0 == area0.y0 and ((area1.x0 - area0.x1) < averagecharactersize):
        return True
    else:
        return False

def joinAreas(areas:list, debug:bool):
## make recursive version
##    returnlist = []

    if len(areas) == 0:
        return []
    elif len(areas) == 1:
        return areas
    elif (areas[0].intersects(areas[1]) or close(areas[0], areas[1])):
        ##if debug: print ("join")
        newrect = areas[0]
        newrect.y1 = areas[1].y1
        newrect.x1 = areas[1].x1
        newlist = [newrect] + areas[2:]
##        newlist.append(areas[2:])
        return joinAreas(newlist, debug) 
    else:
##        ##if debug: print ("nojoin")
##        ##if debug: print ("areas[:1]",areas[:1])
        return areas[:1]+joinAreas(areas[1:],debug)    

##def saveDocToPDF(doc, input_pdf, outfldr, timestr, debug):
def saveDocToPDF(doc, p0, p1, outfldr, timestr, debug):
    global pathsep
    if debug: print("600")
    print('pathsep: ',pathsep)
    pathnew = p0+pathsep+outfldr+pathsep+p1+timestr+"HI.pdf"
    print("pathnew: ",pathnew)
    doc.save(path)
    doc.close()

def split(delimiters, string, maxsplit):
    regexPattern = '|'.join(map(regex.escape, delimiters))
    return regex.split(regexPattern, string, maxsplit)

def MS_Recognize(string:str):
  
    if string not in ['', 'exit']:
            # Retrieve all the ModelResult recognized from the user input

            #if simpledate then reverse day and month
            
            matchobj = regex.search(patternRegexSimpleDate, string)
            if matchobj:
                date_list = split("[\/\.\-\|]", matchobj.group(0), 0)
##                date_list = matchobj.group(0).split("[\/\.\-\|]")
##                ##if debug: print(date_list)
##                string = date_list[1] + "/" + date_list[0] + "/" + date_list[2]
            else:
                0
            results = parse_all(string, DEFAULT_CULTURE)

            inspect.getmembers(results)

            # Flatten results
            results = [item for sublist in results for item in sublist]

            for result in results:
                values = {}
                resultDict = result.__dict__
                resolutionDict = resultDict['resolution']
                values = resolutionDict['values']
                return(
                    json.dumps(
##                        result,
                        resolutionDict,
                        default=lambda o: o.__dict__,
                        indent='\t',
                        ensure_ascii=False))
##                if verbose >= 0.5:
##                    ##if debug: print (resolutionDict)
##                    ##if debug: print (values)
##

def parse_all(user_input: str, culture: str) -> List[List[ModelResult]]:
    return [
        # Number recognizer - This function will find any number from the input
        # E.g "I have two apples" will return "2".
       # Recognizers.recognize_number(user_input, culture),

        # Ordinal number recognizer - This function will find any ordinal number
        # E.g "eleventh" will return "11".
       # Recognizers.recognize_ordinal(user_input, culture),

        # Percentage recognizer - This function will find any number presented as percentage
        # E.g "one hundred percents" will return "100%"
     #   Recognizers.recognize_percentage(user_input, culture),

        # Age recognizer - This function will find any age number presented
        # E.g "After ninety five years of age, perspectives change" will return
        # "95 Year"
     #   Recognizers.recognize_age(user_input, culture),

        # Currency recognizer - This function will find any currency presented
        # E.g "Interest expense in the 1988 third quarter was $ 75.3 million"
        # will return "75300000 Dollar"
        #Recognizers.recognize_currency(user_input, culture),

        # Dimension recognizer - This function will find any dimension presented E.g "The six-mile trip to my airport
        # hotel that had taken 20 minutes earlier in the day took more than
        # three hours." will return "6 Mile"
    #    Recognizers.recognize_dimension(user_input, culture),

        # Temperature recognizer - This function will find any temperature presented
        # E.g "Set the temperature to 30 degrees celsius" will return "30 C"
    #    Recognizers.recognize_temperature(user_input, culture),

        # DateTime recognizer - This function will find any Date even if its write in colloquial language -
        # E.g "I'll go back 8pm today" will return "2017-10-04 20:00:00"
        Recognizers.recognize_datetime(user_input, culture),

        # PhoneNumber recognizer will find any phone number presented
        # E.g "My phone number is ( 19 ) 38294427."
        #Recognizers.recognize_phone_number(user_input, culture),

        # Email recognizer will find any phone number presented
        # E.g "Please write to me at Dave@abc.com for more information on task
        # #A1"
        Recognizers.recognize_email(user_input, culture),
    ]

