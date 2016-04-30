=============
ooxmlreplacer
=============
ooxmlreplacer is a small package that allows you to search and replace in Office OpenXML files.


Supported file formats
======================
* docx, docm
* pptx, pptm


How to use ooxmlreplacer
========================
::

    from ooxmlreplacer import docxreplacer, pptxreplacer
    docxreplacer.replace(infile='test.docx', outfile='test.docx', find_what='foo', replace_with='bar', match_case=False)
    pptxreplacer.replace(infile='test.pptx', outfile='test.pptx', find_what='foo', replace_with='bar', match_case=False)

