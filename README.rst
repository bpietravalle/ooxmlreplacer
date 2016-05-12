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


Using ooxmlreplacer in your script
----------------------------------
::

    from ooxmlreplacer import docxreplacer, pptxreplacer
    docxreplacer.replace(infile='test.docx', outfile='test.docx', find_what='foo', replace_with='bar', match_case=False)
    pptxreplacer.replace(infile='test.pptx', outfile='test.pptx', find_what='foo', replace_with='bar', match_case=False)


Running ooxmlreplacer from commandline
--------------------------------------
::

    python docxreplacer.py "infile.docx" "outfile.docx" -f "foo" -r "bar"
    python pptxreplacer.py "infile.pptx" "outfile.pptx" -f "foo" -r "bar"

