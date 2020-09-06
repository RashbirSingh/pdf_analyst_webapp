regexSimpleDate = "((\d{1,2})[\/\.\-\|]\d{1,2}[\/\.\-\|]\d{2,4})"

regexGeneric = {}
regexGeneric["NUMBER"] = " (\d{1,3}[,\.\']*)+\d*[ \?\!\.]"
regexGeneric["WORD"] = "[A-Za-z]+\W"
regexGeneric["DATE"] = "(?m)(?<!\.)(?<!\.0)((\d{1,2})[\/\.\-\|]\d{1,2}[\/\.\-\|]\d{2,4})|(((\d{1,2}((st)|(nd)|(rd)|(th))?(\sday)?((\sof)|(\sin))?\W((January)|(February)|(March)|(April)|(May)|(June)|(July)|(August)|(September)|(October)|(November)|(December))\W+((\d{4}|('\d{2})))?)))|((January)|(February)|(March)|(April)|(May)|(June)|(July)|(August)|(September)|(October)|(November)|(December))(\sthe)?\W+\d{1,2}((st)|(nd)|(rd)|(th))\W+(\d{4}|('\d{2}))|((\January)|(February)|(March)|(April)|(\WMay\W)|(June)|(July)|(August)|(September)|(October)|(November)|(December))|(?<!\n])(?<!\/)(?<!(NSW ))(?<![\.\/])(((19)|(20)|')\d{2})"
regexGeneric["TIME"] = "(?<=\W)(\d{1,2} o'clock)|(\d{1,2}[.:]\d{2})\W?(([Aa]|[Pp])\.?[Mm]\.?)?(?=\W)"
## need to fix caselaw Regex which is performing poorly
# regexGeneric["CASELAW"]= "(?m)([A-Z][A-Za-z]*\W{1,2})(([A-Z]+[a-z]*\W{1,2})|(for\W{1,2})|(and\W{1,2})|(the\W{1,2})|(v\W{1,2}))*(\(No\.?\W{0,2}\d{1,2}\)\W{0,2}|\(in\Wliq\.?\)\W{1,2})?((\d+)?\W{0,2})([\[\(]\d{4}[\]\)]\W{0,2})((\d+)?\W{0,2})([A-Z]{2,5}\W{0,2})\d{1,5}"
regexGeneric["CASELAW"]= r"([A-Z][A-Za-z]*\W{1,2})(([A-Z]+[a-z]*\W{1,2})|(for\W{1,2})|(and\W{1,2})|(the\W{1,2})|(v\W{1,2}))*(\(No\.?\W{0,2}\d{1,2}\)\W{0,2}|\(in\Wliq\.?\)\W{1,2})?((\d+)?\W{0,2})([\[\(]\d{4}[\]\)]\W{0,2})((\d+)?\W{0,2})([A-Z]{2,5}\W{0,2})\d{1,5}"
regexGeneric["BARCODE"] = "(?<=\W)(\w{1,4}\W{1,2}\d{3,4}\W{1,2}\d{3,4}(\W{1,2}\d{3,4})?)"
regexGeneric["EMAIL"] = "(\w+(\.\w+)*@\w*(\.\w+)*)"

regexGeneric["MONEY"] = "(\$\d+((,|\.|\s)\d{3})*\.*\d*)(\W*(hundred|thousand|million|billion))*"
regexGeneric["ADDRESS"] = "(?p)(?m)((\d+((st)|(nd)|(rd)|(th))?)?,?\s?(((Shop)|(Room)|(Level)|(Suite)|(Floor)|(floor)|\/|(Block)|(Unit))\W*)?(\d+\-?\d*\W*))?(\d+\-?\d*\W*)?(([A-Z][a-z]*\W*)?([A-Z][a-z]*\W*)?((Alley)|(Ally)|(Arcade)|(Arc)|(Avenue)|(Ave)|(Boulevard)|(Bvd)|(Bypass)|(Bypa)|(Circuit)|(Cct)|(Close)|(Cl)|(Corner)|(Crn)|(Court)|(Ct)|(Crescent)|(Cres)|(Cul-de-Sac)|(Cds)|(Drive)|(Dr)|(Esplanade)|(Esp)|(Freepost)|(Green)|(Grn)|(Grove)|(Gr)|(Highway)|(Hwy)|(Junction)|(Jnc)|(Lane)|(Lane)|(Link)|(Link)|(Mews)|(Mews)|(Parade)|(G?\.?P?\.?O?\W*Box \d+)|(Pde)|(Place)|(Pl)|(Ridge)|(Rdge)|(Road)|(Rd)|(Square)|(Sq)|(State Route)|(Motorway)|(Mwy)|(Tollway)|(Tlwy)|(Street)|(St)|(Terrace)|(Tce))\W*)([A-Z][a-z]+,?\s?)+((\d{4}),?\s?)?(((Victoria)|(VIC)|(New South Wales)|(NSW)|(South Australia)|(SA)|(Northern Territory)|(NT)|(Western Australia)|(WA)|(Tasmania)|(TAS)|(ACT)|(Queensland)|(QLD))\.?\s?)?((\d{4}),?\s?)?"

##Currently unused Regexes
##    regexGeneric["ORG"] = "(?p)(?m)([A-Za-z1-9])+\W((Proprietary|Pty|Private|Pte)\W+)?(Ltd|Limited|LLC|LLP|plc|PLC|Inc.|Partners(hip)*)"
##    regexGeneric["PHONE"] = "(?<=\W)[\(\+]?\d([\d \(\)\+\-]{9,15})"
regexGeneric["QUOTE"] = r"(?m)(\W[\"“”].*(?:\n|\r\n?)*.+[\"“”]\W|\W[\'‘’].*(?:\n|\r\n?)*.*[\'‘’]\W)" 
##    regexGeneric["EMAIL"] = "(?<=\W)(\w+(\.\w+)*@\w*(\.\w+)*)"
##regexGeneric["LAW"]= "(?m)((([A-Z][A-Za-z]* )((([a-z]*) )?([A-Z][a-z]* ))*v ([A-Za-z]* )+|((Re ([A-Za-z]* )+))))((\(No\.? (\d{1,2}\) )?\) )|\(in liq\.?\) )?((\d+)? ?)([\[\(]\d{4}[\]\)] )((\d+)? ?)([A-Z]{2,5} )\d{1,5}"
#   The following RegEx does a pretty good job of identifying full names in the case of most people but some (esp women and juniors) are left out.
##    regexGeneric["PERSON"] = "(?<=(?i:(Mr|Ms|Mr|Dr|Prof|Mrs|Sir|Dame|Lord|Justice|Chief Justice|Judge)) )(([A-Z]+[a-z]*)( *([A-Z]+[a-z]*)*))"
##regexGeneric["WORD"] = "\p{L}*" - very slow! multilingual
