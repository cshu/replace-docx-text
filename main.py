#!/usr/bin/env python3

import argparse
import tempfile
import os
import pyperclip
import sys
import subprocess
import re
import shutil
import xml.sax.saxutils
from pathlib import Path
from os import listdir
from os.path import isfile
from datetime import datetime

if shutil.which('zip') is None:
 raise Exception('zip command not found')
if shutil.which('unzip') is None:
 raise Exception('unzip command not found')
parser = argparse.ArgumentParser()
parser.add_argument('docxfile')
parser.add_argument('--old', help='Old text', required=True)
parser.add_argument('--replace-whitespace-with-space', help='Replace whitespace with space', action='store_true')
parser.add_argument('--generate-pdf', help='Additional PDF generation', action='store_true')
parser.add_argument('--generate-pdf-with-pandoc', help='Additional PDF generation with pandoc', action='store_true')
nmsce: argparse.Namespace = parser.parse_args()
pattern: str = nmsce.old
replace_whitespace_with_space: bool = nmsce.replace_whitespace_with_space
generate_pdf: bool = nmsce.generate_pdf
generate_pdf_with_pandoc: bool = nmsce.generate_pdf_with_pandoc
doc_filename: str = nmsce.docxfile
#print(nmsce)
tmpdir: str = tempfile.gettempdir()
tmpdir = os.path.join(tmpdir, 'REPLACE_DOCX_TEXT_TMPDIR')
if os.path.lexists(tmpdir):
 shutil.rmtree(tmpdir)
os.makedirs(tmpdir)
tmpdoc = os.path.join(tmpdir, 'tmp.docx')
docxml = os.path.join(tmpdir, 'word/document.xml')
#if 2!=len(sys.argv):
# raise Exception('Must provide exactly 1 arg as DOCX filename')
#scr_filename: str = sys.argv[0]
#doc_filename: str = sys.argv[1]
if not doc_filename.lower().endswith('.docx'):
 raise Exception('DOCX filename not ending with .docx')
if not os.path.isfile(doc_filename):
 raise Exception('DOCX filename not an existing file')
#pattern: str | None = os.getenv('REPLACE_DOCX_TEXT_OLD')
#if pattern is None:
# raise Exception('REPLACE_DOCX_TEXT_OLD not found. Environment variable REPLACE_DOCX_TEXT_OLD is required.')
print('OLD:', pattern)
substitution: str = pyperclip.paste()
#if 'true' == os.getenv('REPLACE_DOCX_TEXT_REPLACE_WHITESPACE_WITH_SPACE'):
if replace_whitespace_with_space:
 substitution = re.sub('\\s+', ' ', substitution)
substitution = xml.sax.saxutils.escape(substitution)
print('NEW:', substitution)
shutil.copy2(doc_filename, tmpdoc)
subprocess.run(['unzip', 'tmp.docx'], cwd=tmpdir, check=True)
if not os.path.isfile(docxml):
 raise Exception('xml containing text is somehow missing')
docxmlpath: Path = Path(docxml)
docxmlstr = docxmlpath.read_text().replace(pattern, substitution)
docxmlpath.write_text(docxmlstr)
subprocess.run(['zip', '-f', 'tmp.docx'], cwd=tmpdir, check=True)
newdocx = doc_filename[:-5]+'_new.docx'
print('New DOCX:', newdocx)
shutil.copy2(tmpdoc, newdocx)
shutil.rmtree(tmpdir)
#if 'true' == os.getenv('REPLACE_DOCX_TEXT_CONVERT_TO_PDF'):
if generate_pdf:
 if generate_pdf_with_pandoc and shutil.which('pandoc') and shutil.which('pdflatex'):
  subprocess.run(['pandoc', '-o', newdocx[:-5]+'.pdf', '-f', 'docx', newdocx], check=True)
 elif shutil.which('libreoffice') is None:
  print('libreoffice not found. PDF not generated.')
 else:
  subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', newdocx], check=True)
print('DONE')
