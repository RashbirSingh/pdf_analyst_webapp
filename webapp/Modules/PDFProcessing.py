import logging
from pdfrw import PdfReader
import io
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
import fitz
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger


class PDFProcessing:

    def __init__(self, path):
        logging.info("Called PDFProcessing")
        self.path = path
        # Will be added

    def get_info(self):
        """
        ----------
        Function
        ----------
        * Looks at meta data using PdfReader
        * Return information in meta data of pdf

        --------
        INPUT
        --------
        PDF Path

        -------
        RETURN
        -------
        Dictionary contain information as:
            Author': info.Info.Author,
            'Creator': info.Info.Creator,
           'Producer': info.Info.Producer,
           'Subject': info.Info.Subject,
           'Title': info.Info.Title,
           'NumberOfPages': number_of_pages

        """
        logging.info('Inside PDFProcessing get_info')
        PDFinfo = PdfReader(self.path)
        number_of_pages = len(PDFinfo.pages)
        information = {'Author': PDFinfo.Info.Author,
                       'Creator': PDFinfo.Info.Creator,
                       'Producer': PDFinfo.Info.Producer,
                       'Subject': PDFinfo.Info.Subject,
                       'Title': PDFinfo.Info.Title,
                       'NumberOfPages': number_of_pages
                       }
        logging.info('Exiting PDFProcessing get_info')
        return information

    def extract_text_from_pdf(self, way='pdfminer', outputType="text"):
        """
        ----------
        Function
        ----------
        * Opens PDF page by page and stores text inside text object
        --------
        INPUT
        --------
        pdf_path : path to the pdf
        way : library to extract text from pdf
        outputType : (default) Text
                     "text": (default) plain text with line breaks. No formatting, no text position details, no images.
                    "html": creates a full visual version of the page including any images.
                            This can be displayed with your internet browser.
                    "dict": same information level as HTML, but provided as a Python dictionary.
                            See TextPage.extractDICT() for details of its structure.
                    "rawdict": a super-set of TextPage.extractDICT().
                            It additionally provides character detail information like XML. See TextPage.extractRAWDICT() for details of its structure.
                    "xhtml": text information level as the TEXT version but includes images.
                            Can also be displayed by internet browsers.
                    "xml": contains no images, but full position and font information down to each single text character.
                            Use an XML module to interpret.


        -------
        RETURN
        -------
        if pdfminer:
            text : (str) raw text object
        if fitz
            textDataList : (list) List of text inside the pdf with size equals number of pages and index signifying page number.
        """

        """ Converting PDF to Text"""
        logging.info('Inside extract_text_from_pdf')
        if way == 'pdfminer':
            resource_manager = PDFResourceManager()
            fake_file_handle = io.StringIO()
            converter = TextConverter(resource_manager, fake_file_handle)
            page_interpreter = PDFPageInterpreter(resource_manager, converter)

            with open(self.path, 'rb') as fh:
                for page in PDFPage.get_pages(fh,
                                              caching=True,
                                              check_extractable=True):
                    page_interpreter.process_page(page)

                text = fake_file_handle.getvalue()

            converter.close()
            fake_file_handle.close()

            # if the text is non zero
            if text:
                logging.info('Successfully extracted text - Exiting extract_text_from_pdf')
                return text

        if way == 'fitz':
            Pdfdocument = fitz.open(self.path)
            textDataList = []
            for pagenumber in range(Pdfdocument.pageCount):
                page = Pdfdocument[pagenumber]
                textDataList.append(page.getText(outputType))
            return textDataList

        logging.info('PDF was not readable - Exiting extract_text_from_pdf')
