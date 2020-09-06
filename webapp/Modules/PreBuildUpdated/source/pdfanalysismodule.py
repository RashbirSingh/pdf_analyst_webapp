import Modules.PreBuildUpdated.source.GUI20200621 as g
from Modules.PreBuildUpdated.source.HiLightModular20200602 import *
from tkinter import Label
import xlsxwriter
import os

import time

d = g.d
file = g.file
docs = g.docs
results = g.results
file_list = g.results
file_sel_list = g.file_sel_list
DocDict = g.DocDict
DocDictList = g.DocDictList
textSentencesDict = g.textSentencesDict
InvColorDictLabelstoColors = g.InvColorDictLabelstoColors
exportHyperlinks = g.exportHyperlinks
pathsep = "/"  # OS dependent - can be changed using os.path
outfldr = ""  # optional output subfolder


def analyse_file(listbox=['test.pdf', 'test1.pdf']):
    """
    listbox - (list) path to pdf

    return:
    dict of d, DocDict, DocDictList, textSentencesDict
    """
    global file, docs, results, file_list, file_sel_list, d, DocDict, DocDictList, textSentencesDict, InvColorDictLabelstoColors

    label3 = Label(fg='red')
    if not tuple(listbox) in results:
        print("ok")
        g.results[tuple(listbox)] = Highlight_Analyse(listbox, InvColorDictLabelstoColors, False, False, False, label3,
                                                    False)
    d = results[tuple(listbox)][0]
    # print(d)
    DocDict = results[tuple(listbox)][1]
    # print(DocDict)
    DocDictList = results[tuple(listbox)][2]
    # print(DocDictList)
    textSentencesDict = results[tuple(listbox)][3]
    # print(textSentencesDict)
    file_sel_list = listbox

    return {'d': d,
            'DocDict': DocDict,
            'DocDictList': DocDictList,
            'textSentencesDict': textSentencesDict}


# ---------------------------------------

def ExportDetailstoExcel(eff=None, label=object):
    global file, time, results, file_sel_list, exportHyperlinks
    timestr = time.strftime("%Y%m%d-%H%M%S")
    d = analyse_file()['d']
    pth = file_sel_list[0]
    dry = os.path.split(pth)[0]
    ##    try:
    if True:
        saveDocDict2Excel(d, dry + pathsep + outfldr + pathsep + "Analysis_", timestr, file_sel_list, exportHyperlinks,
                          False)


def ExportDetailstoExcel(eff=None, label=object):
    global file, time, debug7, results, file_sel_list, exportHyperlinks
    timestr = time.strftime("%Y%m%d-%H%M%S")
    d = results[tuple(file_sel_list)][0]
    pth = file_sel_list[0]
    dry = os.path.split(pth)[0]
    ##    try:
    if True:
        saveDocDict2Excel(d, dry + pathsep + outfldr + pathsep + "Analysis_", timestr, file_sel_list, exportHyperlinks,
                          debug7)
        label.config(text="Saved as " + outfldr + os.path.sep + 'Analysis_' + timestr + 'DATA.xlsx')


