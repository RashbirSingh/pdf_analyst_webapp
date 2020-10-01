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




