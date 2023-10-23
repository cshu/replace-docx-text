#!/usr/bin/env python3

import json
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
#parser.add_argument('--old', help='Old text', required=True)
parser.add_argument('--old', help='Old text to be replaced with clipboard text')
parser.add_argument('--rep', help='Replacement JSON file mapping to file paths')
parser.add_argument('--rep-json', help='Replacement JSON file')
parser.add_argument('--replace-whitespace-with-space', help='Replace whitespace with space', action='store_true')
parser.add_argument('--generate-pdf', help='Additional PDF generation', action='store_true')
parser.add_argument('--generate-pdf-with-pandoc', help='Additional PDF generation with pandoc', action='store_true')
parser.add_argument('--get-pdf-num-of-pages', help='Print PDF number of pages', action='store_true')
parser.add_argument('--keep-temp-files', help='End the program without deleting temp files', action='store_true')
nmsce: argparse.Namespace = parser.parse_args()
rep: str = nmsce.rep
rep_json: str = nmsce.rep_json
pattern: str = nmsce.old
replace_whitespace_with_space: bool = nmsce.replace_whitespace_with_space
generate_pdf: bool = nmsce.generate_pdf
generate_pdf_with_pandoc: bool = nmsce.generate_pdf_with_pandoc
get_pdf_num_of_pages: bool = nmsce.get_pdf_num_of_pages
keep_temp_files: bool = nmsce.keep_temp_files
doc_filename: str = nmsce.docxfile
#print(nmsce)
if rep is None and rep_json is None and pattern is None:
 raise Exception('At least one option among --rep, --old, and --rep-json need to be specified')
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
shutil.copy2(doc_filename, tmpdoc)
subprocess.run(['unzip', 'tmp.docx'], cwd=tmpdir, check=True)
if not os.path.isfile(docxml):
 raise Exception('xml containing text is somehow missing')
docxmlpath: Path = Path(docxml)
docxmlstr = docxmlpath.read_text()


if pattern:
 print('OLD:', pattern)
 substitution: str = pyperclip.paste()
 #if 'true' == os.getenv('REPLACE_DOCX_TEXT_REPLACE_WHITESPACE_WITH_SPACE'):
 if replace_whitespace_with_space:
  substitution = re.sub('\\s+', ' ', substitution)
 substitution = xml.sax.saxutils.escape(substitution)
 print('NEW:', substitution)
 docxmlstr = docxmlstr.replace(pattern, substitution)
if rep_json:
 rep_lst = json.loads(Path(rep_json).read_text())
 for reobj in rep_lst:
  if 'old' in reobj and 'new' in reobj:
   print('OLD:', reobj['old'])
   print('NEW:', reobj['new'])
   if 'count' in reobj:
    docxmlstr = docxmlstr.replace(reobj['old'], reobj['new'], reobj['count'])
   else:
    docxmlstr = docxmlstr.replace(reobj['old'], reobj['new'])
if rep:
 rep_lst = json.loads(Path(rep).read_text())
 for reobj in rep_lst:
  if 'old' in reobj and 'new' in reobj:
   print('OLD:', reobj['old'])
   print('NEW:', reobj['new'])
   nstr: str = Path(reobj['new']).read_text()
   if 'count' in reobj:
    docxmlstr = docxmlstr.replace(reobj['old'], nstr, reobj['count'])
   else:
    docxmlstr = docxmlstr.replace(reobj['old'], nstr)

docxmlpath.write_text(docxmlstr)
subprocess.run(['zip', '-f', 'tmp.docx'], cwd=tmpdir, check=True)
newdocx = doc_filename[:-5]+'_new.docx'
print('New DOCX:', newdocx)
shutil.copy2(tmpdoc, newdocx)
if not keep_temp_files:
 shutil.rmtree(tmpdir)
#if 'true' == os.getenv('REPLACE_DOCX_TEXT_CONVERT_TO_PDF'):
if generate_pdf:
 if generate_pdf_with_pandoc and shutil.which('pandoc') and shutil.which('pdflatex'):
  subprocess.run(['pandoc', '-o', newdocx[:-5]+'.pdf', '-f', 'docx', newdocx], check=True)
 elif shutil.which('libreoffice') is None:
  print('libreoffice not found. PDF not generated.')
 else:
  subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', newdocx], check=True)
  dst: str = shutil.move(newdocx[:-5]+'.pdf', doc_filename[:-5]+'.pdf')
  if get_pdf_num_of_pages and shutil.which('pdfinfo'):
   subprocess.run('pdfinfo '+dst+' | grep -- ^Pages', shell=True, check=True)#fixme if dst contains special character this will fail
#print('DONE')
