# Creating an empty dictionary
## TODO add the juge and others here too
ColorDict = {}
  
# Adding list as value
# Parse: Color tuple, KEY, Style, Opacity, CheckVar)
ColorDict["Grey"] = ((0.5,0.5,0.5), "WORD", "Highlight", 0.25, 0)
ColorDict["Black"] = ((1.0,1.0,1.0), "REDACTED", "Highlight", 0.25, 1)
ColorDict["Red"] = ((1,0,0), "NEGATIVE_POINT", "Highlight", 0.25, 1)
ColorDict["Orange"] = ((1,0.5,0), "KEY_POINT", "Highlight", 0.25, 1)
ColorDict["Yellow"] = ((1,1,0), "HIGHLIGHT", "Highlight", 0.25, 1)
ColorDict["Green"] = ((0.5,1,0), "DATE", "Highlight", 0.25, 1)

ColorDict["AdobeDkGreen"] = ((0,0.6,0), "Unused1", "Highlight", 0.25, 1)
ColorDict["AdobeLtGreen"] = ((0,1,0.4), "Unused2", "Highlight", 0.25, 1)
ColorDict["AdobeDkBlue"] = ((0,0.2, 1), "Unused3", "Highlight", 0.25, 1)
ColorDict["AdobePurple"] = ((0.6,0, 1), "Unused4", "Highlight", 0.25, 1)
ColorDict["AdobeDkGrey"] = ((0.6,0.6, 0.6), "Unused5", "Highlight", 0.25, 1)

ColorDict["MS_Yellow"] = ((1,0.9,0), "Unused6", "Highlight", 0.25, 1)
##[1.0, 0.901960015296936, 0.0]
ColorDict["MS_LtGreen"] = ((0.1,0.9, 0.0), "Unused8", "Highlight", 0.25, 1)
##[0.1490200012922287, 0.901960015296936, 0.0]
ColorDict["MS_Cyan"] = ((0.3,0.8,1), "Unused7", "Highlight", 0.25, 1)
##[0.26666998863220215, 0.784309983253479, 0.9607800245285034]
ColorDict["MS_Pink"] = ((0.9,0, 0.6), "Unused9", "Highlight", 0.25, 1)
##[0.9254900217056274, 0.0, 0.549019992351532]

ColorDict["Lime"] = ((0,1,0), "POSITIVE_POINT", "Highlight", 0.25, 1)
ColorDict["Aqua"] = ((0,1,0.5), "TIME", "Highlight", 0.25, 1)
ColorDict["Cyan"] = ((0,1,1), "PERSON", "Highlight", 0.25, 1)
ColorDict["Mid_Blue"] = ((0.2,0.7,1), "ORG", "Highlight", 0.25, 1)
ColorDict["Dark_Blue"] = ((0,0,1), "Unused10", "Highlight", 0.25, 1)
ColorDict["Purple"] = ((0.6,0.3,1), "PHONE", "Highlight", 0.25, 1)
ColorDict["Magenta"] = ((1,0,1), "EMAIL", "Highlight", 0.25, 1)
ColorDict["HotPink"] = ((1,0,0.5), "ADDRESS", "Highlight", 0.25, 1)
ColorDict["Brick"] = ((1,0.4,0.4), "LAW", "Highlight", 0.25, 1)
ColorDict["Brick2"] = ((0.8,0.4,0.4), "CASELAW", "Highlight", 0.25, 1)
ColorDict["Salmon"] = ((1,0.6,0.6), "BARCODE", "Highlight", 0.25, 1)
ColorDict["PalePink"] = ((1,0.8,0.8), "NUMBER", "Highlight", 0.25, 1)
ColorDict["Light_Purple"] = ((0.8,0.6,1), "QUOTE", "Highlight", 0.25, 1)
ColorDict["LtYellow"] = ((1,1,0.6), "MONEY", "Highlight", 0.25, 1)

# ColorDict["Tan"] = ((0.82, 0.7, 0.54), "CASENAME", "Highlight", 0.25, 1)
# ColorDict["Dark_Green"] = ((0, 0.39, 0), "CITATION", "Highlight", 0.25, 1)
ColorDict["Light_Salmon"] = ((1, 0.6, 0.47), "INSTRUMENT", "Highlight", 0.25, 1)
ColorDict["Tomato"] = ((1, 0.3, 0.27), "PROVISION", "Highlight", 0.25, 1)
ColorDict["Light_Coral"] = ((0.9, 0.5, 0.5), "COURT", "Highlight", 0.25, 1)
ColorDict["Amethyst"] = ((0.6, 0.4, 0.8), "JUDGE", "Highlight", 0.25, 1)


InvColorDictLabelstoColors = {}
for item in ColorDict.items():
    InvColorDictLabelstoColors[item[1][1]] = [item[1][0], item[1][2], item[1][3], item[1][4]] 

InvColorDicttoNames = {}
for item in ColorDict.items():
    InvColorDicttoNames[item[1][0]] = (item[0])

#for item in InvColorDicttoNames.items():
#    print(item)

InvColorDicttoLabels = {}
for item in ColorDict.items():
    InvColorDicttoLabels[item[1][0]] = (item[1][1])

#for item in InvColorDicttoLabels.items():
#    print(item)
    
    


