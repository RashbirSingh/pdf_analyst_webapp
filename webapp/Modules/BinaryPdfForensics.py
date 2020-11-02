import argparse
import csv
import hashlib
import os
import re
import shlex
import shutil
from datetime import datetime
from functools import partial
from shutil import copyfile
from subprocess import Popen, PIPE
import fitz
from PyPDF2 import PdfFileReader


class BinaryPdfForensics:

    def __init__(self, file_path):
        self.file_path = file_path

    def file_stats(self):
        """Calculate file statistics

        This method calculates the statistics which constitute
        a file's file system metadata, turning them into human
        readable strings ready for insertion into the metadata
        report.

        Args:
            file_path: The path (can be full or abbreviated) of
            the file to be tested.

        Returns:
            This method returns a list containing five string
            elements: the file's (1) absolute path, (2) human
            readable size, (3) access time, (4) modification
            time, and (5) change time.
        """
        stats = os.stat(self.file_path)
        file_abspath = os.path.abspath(self.file_path)
        # calculate file size in bytes
        byte_size = stats[6]
        if byte_size < 1000:
            human_size = str(byte_size) + ' bytes'
        # calculate file size in KBs
        elif byte_size < 1000000:
            human_size = '{0} KB ({1} bytes)'.format(
                str(byte_size / 1000.0),
                str(byte_size)
            )
        # calculate file size in MBs
        elif byte_size < 1000000000:
            human_size = '{0} MB ({1} bytes)'.format(
                str(byte_size / 1000000.0),
                str(byte_size)
            )
        # calculate file size in GBs
        elif byte_size < 1000000000000:
            human_size = '{0} GB ({1} bytes)'.format(
                str(byte_size / 1000000000.0),
                str(byte_size)
            )
        # calculate file size in TBs
        elif byte_size < 1000000000000000:
            human_size = '{0} TB ({1} bytes)'.format(
                str(byte_size / 1000000000000.0),
                str(byte_size)
            )
        # calculate file access time
        atime = datetime.utcfromtimestamp(int(stats[7])). \
            strftime('%Y-%m-%d %H:%M:%S')
        # calculate file modification time
        mtime = datetime.utcfromtimestamp(int(stats[8])). \
            strftime('%Y-%m-%d %H:%M:%S')
        # calculate file change time
        ctime = datetime.utcfromtimestamp(int(stats[9])). \
            strftime('%Y-%m-%d %H:%M:%S')
        file_sys_meta = [
            file_abspath,
            human_size,
            atime,
            mtime,
            ctime
        ]
        return file_sys_meta

    def file_hashes(self):
        """Calculates file hashes

        This method reads the input file as a binary stream,
        and then calculates hash digests of the file for each
        of the hashing algorithms supported by Python's
        hashlib.

        Args:
            file_path: The path (can be full or abbreviated) of
            the file to be tested.

        Returns:
            This method returns a list of string values for the
            digest of the file's hash for each hashing algorithm:
            (1) MD5, (2) SHA1, (3) SHA224, (4) SHA256, (5) SHA384,
            (6) SHA512.
        """
        with open(self.file_path, 'rb') as f:
            # initiate hashing algorithms
            d_md5 = hashlib.md5()
            d_sha1 = hashlib.sha1()
            d_sha224 = hashlib.sha224()
            d_sha256 = hashlib.sha256()
            d_sha384 = hashlib.sha384()
            d_sha512 = hashlib.sha512()
            # update hashing with partial byte stream buffer
            for buf in iter(partial(f.read, 128), b''):
                d_md5.update(buf)
                d_sha1.update(buf)
                d_sha224.update(buf)
                d_sha256.update(buf)
                d_sha384.update(buf)
                d_sha512.update(buf)
        # calculate digest for each hash
        md5_hash = d_md5.hexdigest()
        sha1_hash = d_sha1.hexdigest()
        sha224_hash = d_sha224.hexdigest()
        sha256_hash = d_sha256.hexdigest()
        sha384_hash = d_sha384.hexdigest()
        sha512_hash = d_sha512.hexdigest()
        # store hashes for return
        hash_list = [
            md5_hash,
            sha1_hash,
            sha224_hash,
            sha256_hash,
            sha384_hash,
            sha512_hash
        ]
        return hash_list

    def get_info_ref(self):
        """Tests if a PDF file contains an /Info reference

        This method reads the input file as a binary stream,
        and then performs a regex search, to determine if it
        contains a document information dictionary reference.
        This method only works with valid PDF files, and will
        produce unexpected errors if used on other file types.
        PDF file type should be verified first with the
        pdf_magic() method.

        Args:
            temp_path: The path (can be full or abbreviated) of
            the file to be tested.

        Returns:
            This method returns a tuple containing two values:
            (1) a Boolean value identifying whether or not the PDF
            contains the /Info reference, and (2) a list of any
            located binary document information dictionary string
            references. Example of the info value:

                /Info 2 0 R
        """
        with open(self.file_path, 'rb') as raw_file:
            read_file = raw_file.read()
            regex = b'[/]Info[\s0-9]*?R'
            pattern = re.compile(regex, re.DOTALL)
            info_ref = re.findall(pattern, read_file)
            info_ref = self.de_dupe_list(info_ref)
            if len(info_ref) == 0:
                info_ref_exists = False
            else:
                info_ref_exists = True
            return (info_ref_exists, info_ref)

    def get_info_obj(self):
        """Extracts /Info objects from PDF file

        This method reads the input file as a binary stream,
        and then calls the get_info_ref() function to get any
        /Info references in the file. Any located /Info refs
        are then used to locate any matching /Info objects
        in the file. Example of the /Info object:

            2 0 obj
            << ... >>
            endobj

        Args:
            temp_path: The path (can be full or abbreviated) of
            the file to be tested.

        Returns:
            This method returns a tuple containing the following
            elements: (1) a boolean value of whether or not an
            /Info object exists, and (2) a dictionary which maps
            the /Info references with their objects.
        """
        with open(self.file_path, 'rb') as raw_file:
            read_file = raw_file.read()
            info_ref_tuple = self.get_info_ref()
            info_obj_dict = {}
            for ref in info_ref_tuple[1]:
                info_ref = ref.decode()
                info_ref = info_ref.replace('/Info ', '') \
                    .replace(' R', '')
                info_ref = str.encode(info_ref)
                regex = b'[^0-9]' + info_ref + b'[ ]obj.*?endobj'
                pattern = re.compile(regex, re.DOTALL)
                info_obj = re.findall(pattern, read_file)
                info_obj = self.de_dupe_list(info_obj)
                if len(info_obj) > 0:
                    for obj in info_obj:
                        info_obj_dict[ref] = obj
            if len(info_obj_dict) == 0:
                info_obj_exists = False
            else:
                info_obj_exists = True
            return (info_obj_exists, info_obj_dict)

    def get_xmp_ref(self):
        """Tests if a PDF file contains a /Metadata reference

        This method reads the input file as a binary stream,
        and then performs a regex search, to determine if it
        contains an XMP metadata reference. This method only
        works with valid PDF files, and will produce unexpected
        errors if used on other file types. PDF file type should
        be verified first with the pdf_magic() method.

        Args:
            temp_path: The path (can be full or abbreviated) of
            the file to be tested.

        Returns:
            This method returns a tuple containing two values:
            (1) a Boolean value identifying whether or not the PDF
            contains the /Metadata reference, and (2) a list of
            any located binary XMP metadata string references.
            Example of the XMP metadata reference:

                /Metadata 3 0 R
        """
        with open(self.file_path, 'rb') as raw_file:
            read_file = raw_file.read()
            regex = b'[/]Metadata[\s0-9]*?R'
            pattern = re.compile(regex, re.DOTALL)
            xmp_ref = re.findall(pattern, read_file)
            xmp_ref = self.de_dupe_list(xmp_ref)
            if len(xmp_ref) == 0:
                xmp_ref_exists = False
            else:
                xmp_ref_exists = True
            return (xmp_ref_exists, xmp_ref)

    def de_dupe_list(self, list_var):
        """Removes duplicate elements from list

        This function reads the input list, and creates a new
        blank list. It iterates over each element of the input
        list and adds each unique elelemnt to the new list, with
        any duplicated elements being ignored.

        Args:
            list_var: Any list object.

        Returns:
            This function returns a list of unique elements from
            the input arg.
        """
        new_list = []
        for element in list_var:
            if element not in new_list:
                new_list.append(element)
        return new_list

    def get_xmp_obj(self):
        """Extracts /Metadata objects from PDF file

        This method reads the input file as a binary stream,
        and then calls the get_xmp_ref() function to get any
        /Info references in the file. Any located /Metadata refs
        are then used to locate any matching /Metadata objects
        in the file. Example of /Metadata object:

            3 0 obj
            <</Length 4718/Subtype/XML/Type/Metadata>>stream
            <?xpacket begin="ï»¿" id="W5M0MpCehiHzreSzNTczkc9d"?>
            <x:xmpmeta ...
            </x:xmpmeta>
            <?xpacket end="w"?>
            endstream
            endobj

        Args:
            temp_path: The path (can be full or abbreviated) of
            the file to be tested.

        Returns:
            This method returns a tuple containing the following
            elements: (1) a boolean value of whether or not an
            /Metadata object exists, and (2) a dictionary which
            maps the /Metadata references with their objects.
        """
        with open(self.file_path, 'rb') as raw_file:
            read_file = raw_file.read()
            xmp_ref_tuple = self.get_xmp_ref()
            xmp_obj_dict = {}
            for ref in xmp_ref_tuple[1]:
                xmp_ref = ref.decode()
                xmp_ref = xmp_ref.replace('/Metadata ', '') \
                    .replace(' R', '')
                xmp_ref = str.encode(xmp_ref)
                regex = b'[^0-9]' + xmp_ref + b'[ ]obj.*?endobj'
                pattern = re.compile(regex, re.DOTALL)
                xmp_obj = re.findall(pattern, read_file)
                xmp_obj = self.de_dupe_list(xmp_obj)
                if len(xmp_obj) > 0:
                    for obj in xmp_obj:
                        xmp_obj_dict[ref] = obj
            if len(xmp_obj_dict) == 0:
                xmp_obj_exists = False
            else:
                xmp_obj_exists = True
            return (xmp_obj_exists, xmp_obj_dict)

    def pdf_magic(self):
        """Tests if file contains PDF magic number

        This method reads the input file as a binary stream,
        and interrogates the first four bytes, to determine
        if they decode to the PDF magic number ('%PDF'). This
        is used to determine whether or not a file, regardless
        of its extension, is a PDF.

        Args:
            file_path: The path (can be full or abbreviated) of
            the file to be tested.

        Returns:
            This method returns a tuple containing two values:
            (1) a Boolean value identifying whether or not the
            input path is a PDF file, and (2) a string description
            of the magic assessment. For PDF files, this string
            value will contain the version of the PDF. Example of
            the magic number:

                %PDF-1.4

        Exceptions:
            UnicodeDecodeError: Bytes cannot be decoded.

            IsADirectoryError: Input is a directory.

            FileNotFoundError: Input path does not exist.
        """
        try:
            with open(self.file_path, 'rb') as raw_file:
                read_file = raw_file.read()
                magic_val = read_file[0:4].decode()
                pdf_version = read_file[1:8].decode()
                if magic_val == '%PDF':
                    return (True, pdf_version)
                else:
                    return (False, 'Non-PDF File')
        except UnicodeDecodeError:
            return (False, 'Non-PDF File')
        except IsADirectoryError:
            return (False, 'Directory')
        except FileNotFoundError:
            return (False, 'File Not Found')



    def get_image(self):
        if not os.path.exists('extractedImages'):
            os.makedirs('extractedImages')
        doc = fitz.open(self.file_path)
        for i in range(len(doc)):
            for img in doc.getPageImageList(i):
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n < 5:  # this is GRAY or RGB
                    pix.writePNG("extractedImages/p%s-%s.png" % (i, xref))
                else:  # CMYK: convert to RGB first
                    pix1 = fitz.Pixmap(fitz.csRGB, pix)
                    pix1.writePNG("extractedImages/p%s-%s.png" % (i, xref))
                    pix1 = None
                pix = None

    def pdftoimage(self):
        if not os.path.exists('pdftoimage'):
            os.makedirs('pdftoimage')
        pdf = PdfFileReader(open(self.file_path, 'rb'))
        noOfPages = pdf.getNumPages()

        doc = fitz.open(self.file_path)
        for pageNumber in range(noOfPages):
            page = doc.loadPage(pageNumber)  # number of page
            pix = page.getPixmap()
            output = 'pdftoimage/'+str(pageNumber+1)+'.png'
            pix.writePNG(output)




