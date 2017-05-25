#!/usr/bin/env python

"""
modif_docx.py: Short demonstration how to use python-docx (http://python-docx.readthedocs.io)
to modify a document using bookmarks.

Dependencies:
 -  Python 3.5 (should work with other versions)
 -  lxml >= 2.3.2
 -  python-docx >= 0.8.6
"""


from docx import Document
from docx.oxml.shared import qn
from docx.oxml.text.run import CT_R
__author__ = "Pascal Renauld"
__copyright__ = "Copyright 2017"
__license__ = "WTFPL"
__version__ = "0.1.1.0.1.1.1.1.0.1.1.0.1.1.1.0.0.1.1.0.0.1.0.1"

""" Return the paragraph associated to a bookmark
For details, see: https://stackoverflow.com/questions/24965042/python-docx-insertion-point """
def get_bookmark_parent(doc, bookmark_name):
    doc_element = doc.part.element
    bookmarks_list = doc_element.findall('.//' + qn('w:bookmarkStart'))
    for bookmark in bookmarks_list:
        if bookmark.get(qn('w:name')) == bookmark_name:
            return bookmark.getparent()
    return None
""" Replace the text of the first run found in a paragraph """
def replace_run_text(paragraph, text):
    for child in paragraph:
        if isinstance(child, CT_R):
            child.text = text
            return

document = Document("ransomware_report modif.docx")

par_name = get_bookmark_parent(document, "name_001")
replace_run_text(par_name, "Potato")
par_ext = get_bookmark_parent(document, "extensions_001")
replace_run_text(par_ext, ".spud")
par_note = get_bookmark_parent(document, "note_001")
replace_run_text(par_note, "GiveMeYourStarch.txt")
par_algo = get_bookmark_parent(document, "algo_001")
replace_run_text(par_algo, "Mash")
par_comment = get_bookmark_parent(document, "comment_001")
replace_run_text(par_comment, "Party Kartoffeln")

document.save("ransomware_report modif_v2.0.docx")
