from celery.result import AsyncResult
from django.core.files import File
from django.shortcuts import render
from webappproject import settings
from pdfscanner.signup import newnserform
from pdfscanner.forms import UserForm, UserProfileInfoForm, documentform, usersettingsform
from django.urls import reverse
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseRedirect, HttpResponse, JsonResponse
from django.contrib.auth import authenticate, login, logout
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode
from django.utils.encoding import force_bytes
from django.contrib.auth import get_user_model
from django.contrib.auth.models import User
from django.contrib.auth.tokens import default_token_generator
from django.contrib.sites.shortcuts import get_current_site
from django.template.loader import render_to_string
from django.core.mail import EmailMessage
from .tests import go_to_sleep
from .models import FileDocument, UserSettingsDocument
import probablepeople as pp
import json
from PIL import ImageColor
import rpy2.robjects as robjects
from datetime import datetime
import re
import fitz
from PyPDF2 import PdfFileReader
import pandas as pd
from Modules import BinaryPdfForensics as BPF
from celery import shared_task
from celery_progress.backend import ProgressRecorder

from Modules.PreBuildUpdated.source import GUI20200621 as gui
from Modules.PreBuildUpdated.source import HiColors
import shutil
import os
import pprint
from django.http import FileResponse, Http404

TEMPLATE_DIR_PDFSCANNER = os.path.join(settings.TEMPLATES_DIR, 'pdfscanner')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UserModel = get_user_model()
from pdfscanner.tokens import account_activation_token

r = robjects.r


def index(request):
    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, 'index.html'))


def aboutus(request):
    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "aboutus.html"))


def analysepdf(request):
    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "analysepdf.html"))


def signup(request):
    form = newnserform

    if request.method == 'POST':
        form = newnserform(request.POST)

        if form.is_valid():
            form.save(commit=True)
            return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "registeredsuccess.html"))
        else:
            return HttpResponse("Form Invalid")

    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "signup.html"), {'form': form})


@login_required
def multisaver(request, form, files, count):
    obj = form.save(commit=False)
    obj.user = request.user
    obj.file_field = files
    obj.save()


@login_required
def uploaddocument(request):
    if request.method == 'POST':
        form = documentform(request.POST, request.FILES)
        files = request.FILES.getlist('file_field')
        if form.is_valid():
            for f, i in zip(files, range(len(files))):
                form = documentform(request.POST, request.FILES)
                multisaver(request, form, f, i)

            return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "uploaddocument.html"),
                          {'documents': FileDocument.objects.filter(user=request.user)})
    else:
        form = documentform()
    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "uploaddocument.html"), {
        'form': form
    })


def register(request):
    registered = False
    if request.method == 'POST':
        user_form = UserForm(data=request.POST)
        profile_form = UserProfileInfoForm(data=request.POST)

        if user_form.is_valid() and profile_form.is_valid():
            user = user_form.save(commit=False)
            user.is_active = False
            user.set_password(user.password)
            user.save()

            profile = profile_form.save(commit=False)
            profile.user = user

            if 'profile_pic' in request.FILES:
                profile.profile_pic = request.FILES['profile_pic']

            profile.save()
            registered = True
            current_site = get_current_site(request)
            mail_subject = 'Activate your account.'
            message = render_to_string(os.path.join(TEMPLATE_DIR_PDFSCANNER, 'acc_active_email.html'), {
                'user': user,
                'domain': current_site.domain,
                'uid': urlsafe_base64_encode(force_bytes(user.pk)),
                'token': default_token_generator.make_token(user),
            })
            to_email = user_form.cleaned_data.get('email')
            email = EmailMessage(
                mail_subject, message, to=[to_email]
            )
            email.send()
            return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "registeredsuccess.html"),
                          {'info': 'Please confirm your email address to complete the registration'})
        else:
            print(user_form.errors, profile_form.errors)
    else:
        print("NO POST")
        user_form = UserForm()
        profile_form = UserProfileInfoForm()

    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "register.html"),
                  {'user_form': user_form,
                   'profile_form': profile_form,
                   'registered': registered})


def user_login(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')

        user = authenticate(username=username, password=password)

        if user:
            if user.is_active:
                login(request, user)
                return HttpResponseRedirect(reverse('index'))

            else:
                return HttpResponse("Account Not Active")
        else:
            print('someone tried login and failed!')
            print("username {} and password {}".format(username, password))
            return HttpResponse("Invalid login details were supplied")
    else:
        return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "login.html"), {})


def activate(request, uidb64, token):
    try:
        uid = urlsafe_base64_decode(uidb64).decode()
        user = UserModel._default_manager.get(pk=uid)
    except(TypeError, ValueError, OverflowError, User.DoesNotExist):
        user = None
    if user is not None and default_token_generator.check_token(user, token):
        user.is_active = True
        user.save()
        return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "registeredsuccess.html"),
                      {'info': 'Thank you for your email confirmation. Now you can login your account.'})
    else:
        return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "registeredsuccess.html"),
                      {'info': 'Activation link is invalid!'})


@login_required
def user_logout(request):
    logout(request)
    return HttpResponseRedirect('index')



@login_required
def documentsview(request):

    def checkboxcheck(value):
        if value == 'on':
            return 1
        else:
            return 0

    if request.method == 'POST':


        ## check if user want highligh overlapping
        overlap = request.POST.getlist('overlap')
        if len(overlap) > 0:
            overlap = checkboxcheck(overlap[0])
        else:
            overlap = 0
        request.session['overlap'] = overlap  ## Setting session value for the overlap

        ## check if user wants to use valid name detection
        if 'filtername' in request.POST.dict():
            filtername = checkboxcheck(request.POST.getlist('filtername')[0])
        else:
            filtername = 0

        ## check if user wants sort the data or not
        sortdata = 1
        # if 'sortdata' in request.POST.dict():
        #     sortdata = checkboxcheck(request.POST.getlist('sortdata')[0])
        # else:
        #     sortdata = 0

        prioritydict = request.session['prioritydict']  ## Storing the priority list into the request session

        # if 'view_settings' in request.POST:
        #
        #     # if 'user_settings' in request.FILES:
        #     #     usersettings = request.FILES['user_settings'].read()
        #     #     data = json.loads(usersettings)
        #
        #     # elif 'usersettings' in request.POST:
        #     #     with open((settings.MEDIA_DIR + request.POST.getlist('usersettings')[0]).strip()) as fp:
        #     #         data = json.load(fp)
        #
        #     # else:
        #     #     return HttpResponse("Please select a setting to view")
        #
        #     for key, value in data.items():
        #         if (value[0][0] < 1) and (value[0][0] < 1) and (value[0][0] < 1):
        #             data[key][0] = (value[0][0] * 255,
        #                             value[0][1] * 255,
        #                             value[0][2] * 255)
        #
        #     return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "viewsettings.html"),
        #                   {'settingsdata': data})

        # if 'usersettings' in request.POST:
        #     with open((settings.MEDIA_DIR + request.POST.getlist('usersettings')[0]).strip()) as fp:
        #         data = json.load(fp)
        #         gui.InvColorDictLabelstoColors = data

        if 'user_settings' in request.FILES:
            usersettings = request.FILES['user_settings'].read()
            jsonfile = json.loads(usersettings)
            gui.InvColorDictLabelstoColors = jsonfile

        usersettings = request.session['usersettings']
        gui.InvColorDictLabelstoColors = usersettings


        request.session['InvColorDictLabelstoColors'] = gui.InvColorDictLabelstoColors
        ## Getting the list of documents to be analysed
        documentslist = request.POST.getlist('doc')
        if len(documentslist) == 0:
            return HttpResponse("Please select atleast one document to analyse")
        absolutedocumentlist = [settings.BASE_DIR + s for s in documentslist]

        # resultDict = gui.analyse_file_webapp(absolutedocumentlist, overlap, prioritydict)
        resulttask = gui.Highlight_Analyse.delay(absolutedocumentlist, gui.InvColorDictLabelstoColors, False, False,
                                                 False, False, False)

        request.session['sortdata'] = sortdata
        request.session['filtername'] = filtername
        request.session['overlap'] = overlap
        request.session['prioritydict'] = prioritydict
        request.session['documentslist'] = documentslist
        request.session['absolutedocumentlist'] = absolutedocumentlist


        return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "progressbar.html"),
                      context={'task_id': resulttask.task_id,
                               'resulttask': resulttask})

    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "analysedocument.html"),
                  {'documents': FileDocument.objects.filter(user=request.user)})

def analysisresult(request):
    if request.method == 'POST':
        resulttask = request.POST.getlist('resulttask')[0]
        sortdata = request.session['sortdata']
        filtername = request.session['filtername']
        overlap = request.session['overlap']
        prioritydict = request.session['prioritydict']
        documentslist = request.session['documentslist']
        absolutedocumentlist = request.session['absolutedocumentlist']
        datekeeper = {}
        request.session['datekeeper'] = {}

        resultDict = gui.analyse_file_webapp_shared_task(absolutedocumentlist, overlap, prioritydict, resulttask)

        if sortdata == 1:  # Checking if sortdata is set to true
            resultDict['d'] = dict(sorted(resultDict['d'].items(), key=lambda x: x[0]))
            for key, value in list(resultDict['d'].items()):
                if (key != 'DATE') and (key != 'NUMBER'):
                    resultDict['d'][key] = dict(sorted(resultDict['d'][key].items(), key=lambda x: x[0]))

                elif key == 'NUMBER':
                    resultDict['d'][key] = dict(
                        sorted(resultDict['d'][key].items(), key=lambda x: float(re.sub('\D+', '', x[0]))))

                ##TODO: Manage highlighting of date
                elif key == 'DATE':
                    for k, v in list(resultDict['d']['DATE'].items()):
                        monthlist = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august',
                                     'september', 'october', 'november', 'december']
                        newdate = k
                        for month in monthlist:
                            if k.lower() in month:
                                newdate = k + " 1 2020"
                            elif (bool(re.match(r"[0-9]{1,2} [A-z]* ", k)) == False) and \
                                    (bool(re.match(r"[0-9]{1,2} [A-z]*", k)) and \
                                     ('\n' not in k)):
                                newdate = k + " 2020"
                            elif bool(re.match(r"'[1-9]{1,2}", k)):
                                newdate = re.sub("'", "19", k)
                            elif bool(re.match(r"'[0-9]{1,2}", k)):
                                newdate = re.sub("'", "20", k)

                        r('library("lubridate")')
                        new_key = r(
                            'parse_date_time("' + newdate + '", orders = c("ymd", "dmy", "mdy", "bdy", "bY", "b", "Yb", "Y"))[1]')[
                            0]
                        new_key = datetime.utcfromtimestamp(new_key).strftime('%d-%m-%Y')
                        datekeeper[new_key] = k
                        resultDict['d']['DATE'][new_key] = resultDict['d']['DATE'].pop(k)
                    request.session['datekeeper'] = datekeeper
                    resultDict['d']['DATE'] = dict(
                        sorted(resultDict['d']['DATE'].items(), key=lambda x: datetime.strptime(x[0], '%d-%m-%Y')))

        ## TODO: Move this to backend too
        if filtername == 1:  # Checking if filtername is set to true
            for key, value in resultDict['d'].items():
                if key == 'PERSON':
                    for name in list(resultDict['d'][key]):
                        if re.sub(" ", "", name).isalpha() == False:
                            resultDict['d'][key].pop(name)
                        elif pp.tag(name)[1] != 'Person':
                            resultDict['d'][key].pop(name)
                        elif len(name) < 4:
                            resultDict['d'][key].pop(name)
        gui.ExporttoPDF(overlap, prioritydict)

        return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "analysisresult.html"),
                      {'resultdict': resultDict['d'],
                       'documentslist': documentslist,
                       'listofkeys': list(resultDict['d'].keys())})


@login_required
def deletedocument(request, pk):
    doc = FileDocument.objects.get(pk=pk)
    if str(request.user) == str(doc.user):
        doc.delete()
    else:
        return HttpResponse("User not authorised to perform this action")
    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "analysedocument.html"),
                  {'documents': FileDocument.objects.filter(user=request.user)})

@login_required
def downloadfilemeta(request):
    pk = request.GET['key']
    doc = FileDocument.objects.get(pk=pk)
    path = doc.file_field.path
    infodf = {}
    bpf = BPF.BinaryPdfForensics(path)
    file_stats = bpf.file_stats()
    file_hashes = bpf.file_hashes()
    infodf["Path"] = file_stats[0]
    infodf["File Size"] = file_stats[1]
    infodf["Most Recent Access"] = file_stats[2]
    infodf["Most Recent Content Change"] = file_stats[3]
    infodf["Most Recent Metadata Change"] = file_stats[4]

    infodf["MD5 Hash"] = file_hashes[0]
    infodf["SHA1 Hash"] = file_hashes[1]
    infodf["SHA224 Hash"] = file_hashes[2]
    infodf["SHA256 Hash"] = file_hashes[3]
    infodf["SHA384 Hash"] = file_hashes[4]
    infodf["SHA512 Hash"] = file_hashes[5]

    infodf = pd.DataFrame([infodf])
    infodf = infodf.T.reset_index().T
    infodf = infodf.T
    infodf.columns = ['Parameter', 'Value']
    infodf.to_csv("infodf.csv", index=False)
    with open('infodf.csv') as myfile:
        response = HttpResponse(myfile, content_type='text/csv')
        response['Content-Disposition'] = 'attachment; filename=infodf.csv'
        return response


@login_required
def pdftoimage(request):
    pk = request.GET['key']
    doc = FileDocument.objects.get(pk=pk)
    path = doc.file_field.path
    bpf = BPF.BinaryPdfForensics(path)
    bpf.pdftoimage()
    shutil.make_archive("pdftoimage", 'zip', "pdftoimage")
    response = HttpResponse(open(settings.BASE_DIR + "/pdftoimage.zip", 'rb'), content_type='application/zip')
    response['Content-Disposition'] = 'attachment; filename=pdftoimage.zip'
    try:
        shutil.rmtree("pdftoimage")
        os.remove("pdftoimage.zip")
    except:
        pass
    return response

@login_required
def extractimages(request):
    pk = request.GET['key']
    doc = FileDocument.objects.get(pk=pk)
    path = doc.file_field.path
    bpf = BPF.BinaryPdfForensics(path)
    bpf.get_image()
    shutil.make_archive("extractedImages", 'zip', "extractedImages")
    response = HttpResponse(open(settings.BASE_DIR + "/extractedImages.zip", 'rb'), content_type='application/zip')
    response['Content-Disposition'] = 'attachment; filename=extractedImages.zip'
    try:
        shutil.rmtree("extractedImages")
        os.remove("extractedImages.zip")
    except:
        pass
    return response


@login_required
def exportdetailstoexcel(request):
    if request.method == 'POST':
        path = os.path.join(settings.BASE_DIR, 'DATA.xlsx')
        try:
            os.remove(path)
        except:
            pass
        documentslist = request.POST.getlist('doc')
        absolutedocumentlist = [settings.BASE_DIR + s for s in documentslist]
        overlap = request.session['overlap']
        prioritydict = request.session['prioritydict']
        resultDict = gui.analyse_file_webapp(absolutedocumentlist, overlap, prioritydict)
        gui.arrangeAliases(resultDict['d'], False)
        gui.ExportDetailstoExcel()

        path = settings.BASE_DIR + "/DATA.xlsx"
        if os.path.exists(path):
            with open(path, "rb") as excel:
                data = excel.read()
        response = HttpResponse(data, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=DATA.xlsx'
        try:
            os.remove(path)
        except:
            pass
        return response


@login_required
def exporttopdf(request):
    if request.method == 'POST':
        path = os.path.join(settings.MEDIA_DIR, 'documents', 'highlight')
        try:
            shutil.rmtree(path)
        except:
            pass
        os.mkdir(path)
        documentslist = request.POST.getlist('doc')
        absolutedocumentlist = [settings.BASE_DIR + s for s in documentslist]
        overlap = request.session['overlap']
        prioritydict = request.session['prioritydict']
        resultDict = gui.analyse_file_webapp(absolutedocumentlist, overlap, prioritydict)
        gui.InvColorDictLabelstoColors = request.session['InvColorDictLabelstoColors']
        gui.ExporttoPDF(overlap, prioritydict)

        shutil.make_archive("HighlightedPDFs", 'zip', path)
        response = HttpResponse(open(settings.BASE_DIR + "/HighlightedPDFs.zip", 'rb'), content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename=HighlightedPDFs.zip'
        # try:
        #     shutil.rmtree(path)
        # except:
        #     pass
        return response


@login_required
def exportdicttoexceluvo(request):
    if request.method == 'POST':
        path = os.path.join(settings.BASE_DIR, 'DATAUVO.xlsx')
        try:
            os.remove(path)
        except:
            pass
        documentslist = request.POST.getlist('doc')
        absolutedocumentlist = [settings.BASE_DIR + s for s in documentslist]
        overlap = request.session['overlap']
        prioritydict = request.session['prioritydict']
        resultDict = gui.analyse_file_webapp(absolutedocumentlist, overlap, prioritydict)
        gui.arrangeAliases(resultDict['d'], False)
        gui.ExportDicttoExcelUVO()

        path = settings.BASE_DIR + "/DATAUVO.xlsx"
        if os.path.exists(path):
            with open(path, "rb") as excel:
                data = excel.read()
        response = HttpResponse(data, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=DATAUVO.xlsx'
        try:
            os.remove(path)
        except:
            pass
        return response


@login_required
def deletehighlights(request):
    if request.method == 'POST':
        path = os.path.join(settings.MEDIA_DIR, 'documents', 'cleanpdf')
        try:
            shutil.rmtree(path)
        except:
            pass
        os.mkdir(path)
        documentslist = request.POST.getlist('doc')
        absolutedocumentlist = [settings.BASE_DIR + s for s in documentslist]
        overlap = request.session['overlap']
        prioritydict = request.session['prioritydict']
        # resultDict = gui.analyse_file_webapp(absolutedocumentlist, overlap, prioritydict)
        gui.DeleteHighlights(absolutedocumentlist)

        shutil.make_archive("CleanedPDFs", 'zip', path)
        response = HttpResponse(open(settings.BASE_DIR + "/CleanedPDFs.zip", 'rb'), content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename=CleanedPDFs.zip'
        try:
            shutil.rmtree(path)
        except:
            pass
        return response


@login_required
def extracthighlights(request):
    if request.method == 'POST':
        path = os.path.join(settings.BASE_DIR, 'AvaliableHL.xlsx')
        try:
            os.remove(path)
        except:
            pass
        documentslist = request.POST.getlist('doc')
        absolutedocumentlist = [settings.BASE_DIR + s for s in documentslist]
        overlap = request.session['overlap']
        prioritydict = request.session['prioritydict']
        resultDict = gui.analyse_file_webapp(absolutedocumentlist, overlap, prioritydict)
        gui.ExtractHighlights(absolutedocumentlist)

        if os.path.exists(path):
            with open(path, "rb") as excel:
                data = excel.read()
        response = HttpResponse(data, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=AvaliableHL.xlsx'
        try:
            os.remove(path)
        except:
            pass
        return response


@login_required
def exporttocsv(response):
    return HttpResponse("EXPORTING TO CSV")


def highlightreview(request):
    filenamepath = request.GET['filepath']
    page = int(request.GET['page']) - 1
    word = request.GET['word']

    datekeeper = request.session['datekeeper']

    if word in datekeeper.keys():
        word = datekeeper[word]

    doc = fitz.open(filenamepath)
    docpage = doc[page]
    text_instances = docpage.searchFor(word)
    for inst in text_instances:
        highlight = docpage.addRectAnnot(inst)
    pix = docpage.getPixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))
    pix.writePNG('output.png')
    doc.close()

    with open(settings.BASE_DIR + "/output.png", 'rb') as img:
        response = HttpResponse(img, content_type='image/png')
        response['Content-Disposition'] = 'filename=some_file.pdf'
        return response

def settingspage(request):
    prioritydict = {}
    highlightStyleDict = {'addresshighlightstyle': 'ADDRESS',
                          'barhighlightstyle': 'BARCODE',
                          'caselawhighlightstyle': 'CASELAW',
                          'citationhighlightstyle': 'CITATION',
                          'courthighlightstyle': 'COURT',
                          'datehighlightstyle': 'DATE',
                          'emailhighlightstyle': 'EMAIL',
                          'instrumenthighlightstyle': 'INSTRUMENT',
                          'judgehighlightstyle': 'JUDGE',
                          'keypointhighlightstyle': 'KEY_POINT',
                          'lawhighlightstyle': 'LAW',
                          'moneyhighlightstyle': 'MONEY',
                          'negativepointhighlightstyle': 'NEGATIVE_POINT',
                          'numhighlightstyle': 'NUMBER',
                          'orghighlightstyle': 'ORG',
                          'personhighlightstyle': 'PERSON',
                          'positivepointprior': 'POSITIVE_POINT',
                          'provisionprior': 'PROVISION',
                          'quotehighlightstyle': 'QUOTE',
                          'redactedprior': 'REDACTED',
                          'timehighlightstyle': 'TIME',
                          'wordhighlightstyle': 'WORD'
                          }

    rangeDict = {'addressrange': 'ADDRESS',
                 'barrange': 'BARCODE',
                 'caselawrange': 'CASELAW',
                 'citationrange': 'CITATION',
                 'courtrange': 'COURT',
                 'daterange': 'DATE',
                 'emailrange': 'EMAIL',
                 'instrumentrange': 'INSTRUMENT',
                 'judgerange': 'JUDGE',
                 'keypointrange': 'KEY_POINT',
                 'lawrange': 'LAW',
                 'moneyrange': 'MONEY',
                 'negativepointrange': 'NEGATIVE_POINT',
                 'numrange': 'NUMBER',
                 'orgrange': 'ORG',
                 'personrange': 'PERSON',
                 'phonerange': 'PHONE',
                 'positivepointrange': 'POSITIVE_POINT',
                 'provisionrange': 'PROVISION',
                 'quoterange': 'QUOTE',
                 'redactedrange': 'REDACTED',
                 'timerange': 'TIME',
                 'wordrange': 'WORD'
                 }

    checkDict = {'addresscheck': 'ADDRESS',
                 'barcheck': 'BARCODE',
                 'caselawcheck': 'CASELAW',
                 'citationcheck': 'CITATION',
                 'courtcheck': 'COURT',
                 'datecheck': 'DATE',
                 'emailcheck': 'EMAIL',
                 'instrumentcheck': 'INSTRUMENT',
                 'judgecheck': 'JUDGE',
                 'keypointcheck': 'KEY_POINT',
                 'lawcheck': 'LAW',
                 'moneycheck': 'MONEY',
                 'negativepointcheck': 'NEGATIVE_POINT',
                 'numcheck': 'NUMBER',
                 'orgcheck': 'ORG',
                 'personcheck': 'PERSON',
                 'phonecheck': 'PHONE',
                 'positivepointcheck': 'POSITIVE_POINT',
                 'provisioncheck': 'PROVISION',
                 'quotecheck': 'QUOTE',
                 'redactedcheck': 'REDACTED',
                 'timecheck': 'TIME',
                 'wordcheck': 'WORD'
                 }

    colorDict = {'addresscolor': ['ADDRESS', 'customaddresscolor'],
                 'barcolor': ['BARCODE', 'custombarcolor'],
                 'caselawcolor': ['CASELAW', 'customcaselawcolor'],
                 'citationcolor': ['CITATION', 'customcitationcolor'],
                 'courtcolor': ['COURT', 'customcourtcolor'],
                 'datecolor': ['DATE', 'customdatecolor'],
                 'emailcolor': ['EMAIL', 'customemailcolor'],
                 'instrumentcolor': ['INSTRUMENT', 'custominstrumentcolor'],
                 'judgecolor': ['JUDGE', 'customjudgecolor'],
                 'keypointcolor': ['KEY_POINT', 'customkeypointcolor'],
                 'lawcolor': ['LAW', 'customlawcolor'],
                 'moneycolor': ['MONEY', 'custommoneycolor'],
                 'negativepointcolor': ['NEGATIVE_POINT', 'customnegativepointcolor'],
                 'numcolor': ['NUMBER', 'customnumcolor'],
                 'orgcolor': ['ORG', 'customorgcolor'],
                 'personcolor': ['PERSON', 'custompersoncolor'],
                 'phonecolor': ['PHONE', 'customphonecolor'],
                 'positivepointcolor': ['POSITIVE_POINT', 'custompositivepointcolor'],
                 'provisioncolor': ['PROVISION', 'customprovisioncolor'],
                 'quotecolor': ['QUOTE', 'customquotecolor'],
                 'redactedcolor': ['REDACTED', 'customredactedcolor'],
                 'timecolor': ['TIME', 'customtimecolor'],
                 'wordcolor': ['WORD', 'customwordcolor']
                 }

    def checkboxcheck(value):
        if value == 'on':
            return 1
        else:
            return 0

    if request.method == 'POST':

        if 'view_settings' in request.POST:

            if 'user_settings' in request.FILES:
                usersettings = request.FILES['user_settings'].read()
                data = json.loads(usersettings)

            elif 'usersettings' in request.POST:
                with open((settings.MEDIA_DIR + request.POST.getlist('usersettings')[0]).strip()) as fp:
                    data = json.load(fp)

            else:
                return HttpResponse("Please select a setting to view")

            for key, value in data.items():
                if (value[0][0] < 1) and (value[0][0] < 1) and (value[0][0] < 1):
                    data[key][0] = (value[0][0] * 255,
                                    value[0][1] * 255,
                                    value[0][2] * 255)

            return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "viewsettings.html"),
                          {'settingsdata': data})

        #TODO: Remove this shit
        if 'usersettings' in request.POST:
            with open((settings.MEDIA_DIR + request.POST.getlist('usersettings')[0]).strip()) as fp:
                data = json.load(fp)
                gui.InvColorDictLabelstoColors = data
                request.session["usersettings"] = data
            return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "settingspage.html"),
                              {'user_settings': UserSettingsDocument.objects.filter(user=request.user)})

        # ######## USER WANTS TO SAVE SETTINGS
        if 'save_settings' in request.POST:

            for colorelement in colorDict.keys():
                if colorelement in request.POST.dict():
                    if request.POST.getlist(colorelement)[0] == "Custom":
                        hex = request.POST.getlist(colorDict[colorelement][1])[0]
                        gui.InvColorDictLabelstoColors[colorDict[colorelement][0]][0] = ImageColor.getcolor(hex, 'RGB')
                    else:
                        try:
                            gui.InvColorDictLabelstoColors[colorDict[colorelement][0]][0] = \
                                HiColors.ColorDict[(request.POST.getlist(colorelement)[0])][0]
                        except:
                            pass

            for highlightelement in highlightStyleDict.keys():
                if highlightelement in request.POST.dict():
                    try:
                        gui.InvColorDictLabelstoColors[highlightStyleDict[highlightelement]][1] = \
                            request.POST.getlist(highlightelement)[0]
                    except:
                        pass

            for rangeelement in rangeDict.keys():
                if rangeelement in request.POST.dict():
                    try:
                        gui.InvColorDictLabelstoColors[rangeDict[rangeelement]][2] = request.POST.getlist(rangeelement)[
                            0]
                    except:
                        pass

            for checkelement in checkDict.keys():
                if checkelement in request.POST.dict():
                    try:
                        gui.InvColorDictLabelstoColors[checkDict[checkelement]][3] = checkboxcheck(
                            request.POST.getlist(checkelement)[0])
                    except:
                        pass
                else:
                    try:
                        gui.InvColorDictLabelstoColors[checkDict[checkelement]][3] = 0
                    except:
                        pass

            # ## Getting key priority value
            prioritydict['ADDRESS'] = int(request.POST.getlist('addressprior')[0])
            prioritydict['BARCODE'] = int(request.POST.getlist('barprior')[0])
            prioritydict['CASELAW'] = int(request.POST.getlist('caselawprior')[0])
            # prioritydict['CITATION'] = int(request.POST.getlist('citationprior')[0])
            prioritydict['COURT'] = int(request.POST.getlist('courtprior')[0])
            prioritydict['DATE'] = int(request.POST.getlist('dateprior')[0])
            prioritydict['EMAIL'] = int(request.POST.getlist('emailprior')[0])
            # prioritydict['INSTRUMENT'] = int(request.POST.getlist('instrumentprior')[0])
            prioritydict['JUDGE'] = int(request.POST.getlist('judgeprior')[0])
            prioritydict['KEY_POINT'] = int(request.POST.getlist('keypointprior')[0])
            prioritydict['LAW'] = int(request.POST.getlist('lawprior')[0])
            prioritydict['MONEY'] = int(request.POST.getlist('moneyprior')[0])
            prioritydict['NEGATIVE_POINT'] = int(request.POST.getlist('negativepointprior')[0])
            prioritydict['NUMBER'] = int(request.POST.getlist('numberprior')[0])
            prioritydict['ORG'] = int(request.POST.getlist('orgprior')[0])
            prioritydict['PERSON'] = int(request.POST.getlist('personprior')[0])
            prioritydict['PHONE'] = int(request.POST.getlist('phoneprior')[0])
            prioritydict['POSITIVE_POINT'] = int(request.POST.getlist('positivepointprior')[0])
            # prioritydict['PROVISION'] = int(request.POST.getlist('provisionprior')[0])
            prioritydict['QUOTE'] = int(request.POST.getlist('quoteprior')[0])
            prioritydict['REDACTED'] = int(request.POST.getlist('redactedprior')[0])
            prioritydict['TIME'] = int(request.POST.getlist('timeprior')[0])
            prioritydict['WORD'] = int(request.POST.getlist('wordprior')[0])

            request.session['prioritydict'] = prioritydict  ## Storing the priority list into the request session


            save_as = request.POST.getlist('save_settings_as')[0]
            if len(save_as) < 1:
                save_as = 'default'
            path = settings.MEDIA_DIR + "/usersettings/" + str(request.user) + "_" + save_as + ".json"
            try:
                os.remove(path)
            except:
                pass

            with open(path, 'w') as fp:
                json.dump(gui.InvColorDictLabelstoColors, fp)

            response = HttpResponse(open(path, 'rb'), content_type='application/json')
            response['Content-Disposition'] = 'attachment; filename=' + save_as + '.json'
            form = usersettingsform()
            obj = form.save(commit=False)
            obj.user = request.user
            obj.setting_field = "/usersettings/" + str(request.user) + "_" + save_as + ".json"
            obj.save()

            return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "settingspage.html"),
                          {'user_settings': UserSettingsDocument.objects.filter(user=request.user)})

    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "settingspage.html"),
                  {'user_settings': UserSettingsDocument.objects.filter(user=request.user)})


def privacypolicy(request):
    return  render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "privacypolicy.html"))

def enduseragreement(request):
    return  render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "enduseragreement.html"))

def contactus(request):
    return  render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "contactus.html"))


def test(request):
    results = {}
    # result = go_to_sleep.delay(1)
    overlap = request.session['overlap']
    prioritydict = request.session['prioritydict']
    lst = ['/Users/rashbir/Desktop/Sidekicker/LaurenceWhite/Webapp/newwebapp/webapp/media/documents/Rules_re_Costs.pdf',
           '/Users/rashbir/Desktop/Sidekicker/LaurenceWhite/Webapp/newwebapp/webapp/media/documents/Spotwire_No_2_2004_FCA_571.pdf']
    a = gui.Highlight_Analyse.delay(lst, gui.InvColorDictLabelstoColors, False, False, False, False, False)
    # a = gui.test(lst, gui.InvColorDictLabelstoColors)
    # d = a.get()[0]

    if request.method == 'POST':
        a = request.POST.getlist('a')[0]
        d = AsyncResult(a).get()[0]
        return JsonResponse(d['ORG'])
    return render(request, os.path.join(TEMPLATE_DIR_PDFSCANNER, "test.html"),
                  context={'task_id': a.task_id,
                           'a': a})
