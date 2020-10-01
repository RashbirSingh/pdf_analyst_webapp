import copy
from celery.result import AsyncResult
from copy import deepcopy
import pandas
import itertools
import time

from time import sleep

from Modules.PreBuildUpdated.source.HiLightModular20200602 import *
from Modules.PreBuildUpdated.source.HiColors import *
from Modules.PreBuildUpdated.source.arrangeAliases6b import arrangeAliases


globaldebug = False
debug = False or globaldebug
debug2 = False or globaldebug
debug3 = False or globaldebug
debug4 = False or globaldebug
debug5 = False or globaldebug
debug6 = False or globaldebug
debug7 = True or globaldebug

# Preferences
# TODO preferences screen

exportHyperlinks = True
DEFAULT_CULTURE = Culture.English
LOCALES = ['en-AU']  # affects parsing of Dates
pathsep = "/"  # OS dependent

outfldr = ""  # optional output subfolder
ModesDict = {
    "Highlight": "Highlight",
    "Underline": "Underline",
    "Rectangle": "Rect",
    "Squiggly underline": "Squiggly"
}

result = {}
results = {}
docs = {}
d = {}
DocDict = {}
DocDictList = {}
##textPagesDict = {}
textSentencesDict = {}
edited_result = {}
file = ""
file_list = []
file_sel_list = []
## pageinFocus is the same as the page number, which is offset by +1 from the zero-based index of textPagesDict
pageinFocus = 1
fileinFocus = ""
root = None
textWidget = None

defaultColorDict = copy.deepcopy(InvColorDictLabelstoColors)
checkVar = {}
radioVar = {}



def analyse_file(eff=None, listbox=object, listbox2=object, label=object, label2=object, label3=object, label4=object,
                 tree=object, root=object, scalePage=object, button=object, button2=object):
    global file, docs, results, debug, file_list, file_sel_list, d, DocDict, DocDictList, textSentencesDict, InvColorDictLabelstoColors
    lst = []

    sel = listbox.curselection()
    for i in sel:
        lst.append(file_list[i])

    label3.config(text="Analysing " + str(lst))
    if not tuple(lst) in results:
        results[tuple(lst)] = Highlight_Analyse(lst, InvColorDictLabelstoColors, False, False, False, label3, False)
    d = results[tuple(lst)][0]
    DocDict = results[tuple(lst)][1]
    DocDictList = results[tuple(lst)][2]
    textSentencesDict = results[tuple(lst)][3]

    file_sel_list = lst
    return {'d': d,
            'DocDict': DocDict,
            'DocDictList': DocDictList,
            'textSentencesDict': textSentencesDict}


def ExportDicttoExcelUVO(eff=None, label=object):
    global file, time, debug, results, file_sel_list, exportHyperlinks
    timestr = time.strftime("%Y%m%d-%H%M%S")
    d = results[tuple(file_sel_list)][0]
    pth = file_sel_list[0]
    dry = os.path.split(pth)[0]
    saveDict2ExcelUniqueValsOnly(d, dry + pathsep + outfldr + pathsep + "Analysis_", timestr, file_sel_list,
                                 exportHyperlinks, debug7)

def ExportDetailstoExcel(eff=None, label=object):
    global file, time, debug7, results, file_sel_list, exportHyperlinks
    timestr = time.strftime("%Y%m%d-%H%M%S")
    d = results[tuple(file_sel_list)][0]
    pth = file_sel_list[0]
    dry = os.path.split(pth)[0]
    if True:
        saveDocDict2Excel(d, dry + pathsep + outfldr + pathsep + "Analysis_", timestr, file_sel_list, exportHyperlinks,
                          debug7)

def ExporttoPDF(overlap, prioritydict):
    global time, debug6, results, file_sel_list, outfldr
    timestr = time.strftime("%Y%m%d-%H%M%S")
    pathnew, p0, p1 = "", "", ""

    for file in file_sel_list:
        d2 = results[tuple(file_sel_list)][2]

        if overlap == 0:
#### REMOVE OVERLAPPING HIGHLIGHTS
            for pageno in range(len(d2[file])):
                for com in itertools.combinations(prioritydict.keys(), 2):
                    if (com[0] in d2[file][pageno]) and (com[1] in d2[file][pageno]):
                        for k in list(d2[file][pageno][com[0]].keys()):
                            if k in list(d2[file][pageno][com[1]].keys()):
                                # noinspection PyPackageRequirements
                                if prioritydict[com[0]] > prioritydict[com[1]]:
                                    try: d2[file][pageno][com[0]].pop(k)
                                    except: pass
                                elif prioritydict[com[0]] < prioritydict[com[1]]:
                                    try: d2[file][pageno][com[1]].pop(k)
                                    except: pass

                            if k in list(d2[file][pageno][com[0]].keys()):
                                if prioritydict[com[0]] > prioritydict[com[1]]:
                                    try: d2[file][pageno][com[0]].pop(k)
                                    except: pass
                                elif prioritydict[com[0]] < prioritydict[com[1]]:
                                    try: d2[file][pageno][com[1]].pop(k)
                                    except: pass

        p0 = os.path.split(file)[0]
        p1 = os.path.split(file)[1]
        doc = markup(file, d2[file], InvColorDictLabelstoColors, debug6)
        pathnew = p0 + pathsep + outfldr + pathsep + "/highlight/" + p1[:-4] + ".pdf"
        doc.save(pathnew) ##TODO: Add user specific path
        doc.close()


def ExtractHighlights(absolutedocumentlist):
    global time, file_sel_list, debug
    timestr = time.strftime("%Y%m%d-%H%M%S")
    successText, failureText, failure2Text = "", "", ""
    ##    try:
    for file in absolutedocumentlist:
        outputFileName = file + timestr[9:] + ".xlsx"

        res = pdfannot2df(file, outputFileName, debug)
        if res[0] != "": successText += ", " + res[0]
        if res[1] != "": failureText += "\n" + res[1]
        if res[2] != "": failure2Text += "\n" + res[2]


def DeleteHighlights(absolutedocumentlist, eff=None, listbox=object, label=object):
    global time, file_sel_list, debug7
    HLsFound = False
    failureMsg = ""

    file_sel_list = absolutedocumentlist

    timestr = time.strftime("%Y%m%d-%H%M%S")
    successText, failureText, failure2Text = "", "", ""
    for file in file_sel_list:
        outputFileName = file + timestr[9:] + ".pdf"
        p0 = os.path.split(file)[0]
        p1 = os.path.split(file)[1]
        res = deleteHLs(file, debug)
        doc = res[0]
        HLsFound = (HLsFound) or (res[1])
        if HLsFound:
            pathnew = p0 + pathsep + outfldr + pathsep + "/cleanpdf/" + p1[:-4] + " " + "CLN.pdf"
            doc.save(pathnew)
            doc.close()
        else:
            failureMsg = failureMsg + ", " + res[2]
    msg = ""
    if HLsFound:
        msg = "Cleaned files saved with filename(s) ending in '_CLN.pdf'."
    if failureMsg != "":
        msg2 = "No HLs found in these files:" + failureMsg + "."


def deleteHLs(input_pdf, debug):
    HLsFound = False
    failureMsg = ""
    pdf = fitz.open(input_pdf)
    l, order = [], 0
    context = ""
    contentDict = {}
    for ixpage, page in enumerate(pdf):
        annot = page.firstAnnot

        while annot:
            HLsFound = True
            annot = page.deleteAnnot(annot)

    if HLsFound == False: failureMsg = os.path.split(input_pdf)[1]
    return pdf, HLsFound, failureMsg


def rgb2hex(rgb):
    return "#%02x%02x%02x" % rgb1to255(rgb)


def rgb1to255(tpl):
    return (round(tpl[0] * 255), round(tpl[1] * 255), round(tpl[2] * 255))



def _extract_word_from_highlight(annot, words, debug=False):
    """Extracts words behind a highlight
    :param annot: a highlight annotation to extract words from. Warning : if it is followed by others
    highlight annotations which labels are named specifically ('same_label'+ -/- 'next_integer'),
     their words will be extracted too.
     :param words: the words of the page containing the annot.
     :param debug: debug param
    :return: the list of words extracted and the annotation which may have changed.
    """

    mywords = []

    #  If the highlight spans on multiple text boxes (possibly multiple lines
    if len(annot.vertices) > 4:

        for k in range(len(annot.vertices) // 4):
            rectangle = fitz.Rect(annot.vertices[k * 4][0], annot.vertices[k * 4][1],
                                  annot.vertices[k * 4 + 3][0], annot.vertices[k * 4 + 3][1])

            for w in words:
                r = deepcopy(rectangle)
                area_intersect = r.intersect(fitz.Rect(w[:4])).getRectArea()
                area_word = fitz.Rect(w[:4]).getRectArea()

                if area_intersect / area_word > 0.6:
                    mywords.append(w)

    else:
        mywords += [w for w in words if
                    annot.rect.intersect(fitz.Rect(w[:4])).getRectArea() / fitz.Rect(w[:4]).getRectArea() > 0.6]

        if annot.next and annot.next.info['content']:
            label_next = annot.next.info['content']


            while annot.next and annot.next.info['content'] and label_next[(len(label_next) - 4):(
                    len(label_next) - 1)] == '-/-' and int(label_next[len(label_next) - 1]) > 1:
                annot = annot.next
                mywords += [w for w in words if
                            annot.rect.intersect(fitz.Rect(w[:4])).getRectArea() / fitz.Rect(w[:4]).getRectArea() > 0.6]
                if annot.next and annot.next.info['content']:
                    label_next = annot.next.info['content']

    return mywords, annot


def pdfannot2df(input_pdf, outputFileName, debug):
    """Takes an annotated pdf as an input and transforms it into a dlf
    :param input_pdf: path to the pdf.
    :return:the adf corresponding to the pdf's annotations
    """

    pdf = fitz.open(input_pdf)
    l, order = [], 0
    context = ""
    contentDict = {}
    for ixpage, page in enumerate(pdf):
        tmp = {'page': ixpage + 1, 'pdf_path': input_pdf, 'page_width': page.rect[2], 'page_height': page.rect[3]}
        words = page.getTextWords()
        annot = page.firstAnnot

        while annot:
            mywords = []
            date1, date2, content, context = "", "", "", ""

            if annot.type[1] == 'Highlight':
                mywords, annot = _extract_word_from_highlight(annot, words)

            # need to build functionality here that would concatenate results from closely adjacent annotations of the same type - trouble is that 'closely adjacent' will depend on the text size

            elif annot.type[1] == 'Square':
                mywords = [w for w in words if
                           annot.rect.intersect(fitz.Rect(w[:4])).getRectArea() / fitz.Rect(w[:4]).getRectArea() > 0.6]

            else:
                print('encountered an annotation different from "Square" and "Highlights".') if debug else 0

            annot_text = " ".join(w[4] for w in mywords)

            print(mywords) if debug else 0

            order += 1
            print('order : ', order) if debug else 0

            # unpack dict in annot.info['content'] field

            content = annot.info['content']
            print("content: " + content) if debug else 0

            ##            json_acceptable_content = content.replace("'", "\"")
            ##            contentDict = json.loads(json_acceptable_content)

            if content != "":
                print("in content loop") if debug else 0
                try:
                    contentDict = eval(content)
                    print("contentDict: ", contentDict) if debug else 0
                    if contentDict["date"]["values"][0]["type"] == "daterange":
                        date1 = contentDict["date"]["values"][0]["start"]
                        date2 = contentDict["date"]["values"][0]["end"]
                    elif contentDict["date"]["values"][0]["type"] == "date":
                        date1 = contentDict["date"]["values"][0]["value"]
                    context = contentDict["context"]
                except:
                    print("exception") if debug else 0

            ##  are there any other types of dates in the MS_Recognizer output?

            # get RGB triple of annot color

            try:
                lst = (round(annot.colors['stroke'][0], 2), round(annot.colors['stroke'][1], 2),
                   round(annot.colors['stroke'][2], 2))
            except:
                lst = 0
            # update tmp dictionary

            if (lst) in InvColorDicttoLabels:
                tmp.update({'x': int(annot.rect[0]), 'y': int(annot.rect[1]),
                            # Those might be wrong for multi line highlights as the rect only
                            # correspond to the one of the last line
                            'w': int(annot.rect[2] - annot.rect[0]), 'h': int(annot.rect[3] - annot.rect[1]),
                            'type': annot.type[1], 'label': annot.info['content'], \
                            'color': annot.colors['stroke'],
                            'colorName': InvColorDicttoNames[lst], 'colorKey': InvColorDicttoLabels[lst],
                            'order': order, 'text': annot_text, 'date1': date1, 'date2': date2, 'context': context})
            else:
                tmp.update({'x': int(annot.rect[0]), 'y': int(annot.rect[1]), 'w': int(annot.rect[2] - annot.rect[0]),
                            'h': int(annot.rect[3] - annot.rect[1]), 'type': annot.type[1], \
                            'label': annot.info['content'], 'color': annot.colors['stroke'],
                            'colorName': "Unknown", 'colorKey': "Unknown", 'order': order,
                            'text': annot_text, 'date1': date1, 'date2': date2, 'context': context})
            print(tmp) if debug else 0
            l.append(deepcopy(tmp))

            annot = annot.next

    adf = pandas.DataFrame(l)
    print('adf : ', adf) if debug else 0
    successText, failureText, failure2Text = "", "", ""
    if adf.empty:
        print(f'WARNING : the document {input_pdf} does not contain any annotations, the returned dataframe is empty.')
        failureText = input_pdf

    elif adf[adf['type'].isnull()].shape[0]:
        raise Exception(f'Missing {adf[adf["type"].isnull()].shape[0]} type annotation(s) in {input_pdf}')
        failure2Text = input_pdf

    else:
        final_columns = ['order', 'page', 'x', 'y', 'w', 'h', 'type', 'label', 'color', 'colorName', 'colorKey',
                         'page_height', 'page_width', 'pdf_path', 'text', 'date1', 'date2', 'context']
        adf = adf[final_columns]

        ##Export each group of highlights to a separate sheet
        colorKeys = adf['colorKey'].unique().tolist()
        if debug: print("165")
        writer = pandas.ExcelWriter("AvaliableHL.xlsx", engine='xlsxwriter')
        for color in colorKeys:
            mydf = adf.loc[adf.colorKey == color]
            mydf.to_excel(writer, sheet_name=color)

        writer.save()
        successText = outputFileName

    return successText, failureText, failure2Text


@shared_task(bind=True)
def analyse_file_webapp(self, lst, overlap, prioritydict):
    global file, docs, results, debug, file_list, file_sel_list, d, DocDict, DocDictList, textSentencesDict, InvColorDictLabelstoColors


    try:
        label3 = Label()
    except:
        pass
    for value in prioritydict.values():
        if value >= 1:
            results = {}
            break
    #if not tuple(lst) in results:
    returndata = Highlight_Analyse.delay(lst, InvColorDictLabelstoColors, False, False, False, False, False)
    # returndata = Highlight_Analyse(lst, InvColorDictLabelstoColors, False, False, False, False, False)
    results[tuple(lst)] = returndata.get()
    d = results[tuple(lst)][0]


    #### to manage overlapping of highlights
    if overlap == 0: # If user do not want to overlap(Checkbox unticked)
        for com in itertools.combinations(prioritydict.keys(), 2):
            if (com[0] in d) and (com[1] in d):
                for k in list(d[com[0]].keys()):
                    if k in list(d[com[1]].keys()):
                        if prioritydict[com[0]] > prioritydict[com[1]]:
                            try: d[com[0]].pop(k)
                            except: pass
                        elif prioritydict[com[0]] < prioritydict[com[1]]:
                            try: d[com[1]].pop(k)
                            except: pass

                    if k in list(d[com[0]].keys()):
                        if prioritydict[com[0]] > prioritydict[com[1]]:
                            try: d[com[0]].pop(k)
                            except: pass
                        elif prioritydict[com[0]] < prioritydict[com[1]]:
                            try: d[com[1]].pop(k)
                            except: pass

    DocDict = results[tuple(lst)][1]
    DocDictList = results[tuple(lst)][2]
    textSentencesDict = results[tuple(lst)][3]

    file_sel_list = lst


    return {'d': d,
            'DocDict': DocDict,
            'DocDictList': DocDictList,
            'textSentencesDict': textSentencesDict,
            'returndata': returndata}


def ExcelFileCreator(d, fileName, outputpath, exportHyperlinks=True):
    pth = file_sel_list[0]
    dry = os.path.split(pth)[0]
    timestr = time.strftime("%Y%m%d-%H%M%S")
    workbook = xlsxwriter.Workbook(dry + pathsep + outfldr + pathsep + "Analysis_"+timestr+'DATA.xlsx')
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'valign': 'top'})
    date_time_format = workbook.add_format({'num_format': 'ddd d mmm yyyy hh:mm:ss', 'valign': 'top'})
    bold = workbook.add_format({'bold': True, 'valign': 'top'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    cell_format = workbook.add_format()
    cell_format.set_text_wrap()

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
    workbook.close()

@shared_task(bind=True)
def test(self, lst, InvColorDictLabelstoColors):
    result = Highlight_Analyse.delay(lst, InvColorDictLabelstoColors, False, False, False, False, False)
    return result


def analyse_file_webapp_shared_task(lst, overlap, prioritydict, task_id):
    global file, docs, results, debug, file_list, file_sel_list, d, DocDict, DocDictList, textSentencesDict, InvColorDictLabelstoColors

    for value in prioritydict.values():
        if value >= 1:
            results = {}
            break
    taskresultdata = AsyncResult(task_id).get()
    results[tuple(lst)] = taskresultdata
    d = taskresultdata[0]


    #### to manage overlapping of highlights
    if overlap == 0: # If user do not want to overlap(Checkbox unticked)
        for com in itertools.combinations(prioritydict.keys(), 2):
            if com in d:
                for k in list(d[com[0]].keys()):
                    if k in list(d[com[1]].keys()):
                        if prioritydict[com[0]] > prioritydict[com[1]]:
                            try: d[com[0]].pop(k)
                            except: pass
                        elif prioritydict[com[0]] < prioritydict[com[1]]:
                            try: d[com[1]].pop(k)
                            except: pass

                    if k in list(d[com[0]].keys()):
                        if prioritydict[com[0]] > prioritydict[com[1]]:
                            try: d[com[0]].pop(k)
                            except: pass
                        elif prioritydict[com[0]] < prioritydict[com[1]]:
                            try: d[com[1]].pop(k)
                            except: pass

    DocDict = results[tuple(lst)][1]
    DocDictList = results[tuple(lst)][2]
    textSentencesDict = results[tuple(lst)][3]

    file_sel_list = lst


    return {'d': d,
            'DocDict': DocDict,
            'DocDictList': DocDictList,
            'textSentencesDict': textSentencesDict}
