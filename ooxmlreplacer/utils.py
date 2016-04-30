#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os.path
import uuid
import zipfile

from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{%s}' % W_NS

A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
A = '{%s}' % A_NS

S_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
S = '{%s}' % S_NS

XML_NS = 'http://www.w3.org/XML/1998/namespace'
XML = '{%s}' % XML_NS

TMX14_NS = 'http://www.lisa.org/tmx14'
TMX14 = '{%s}' % TMX14_NS

CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
CT = '{%s}' % CT_NS

REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
REL = '{%s}' % REL_NS

R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R = '{%s}' % R_NS

XLF_NS = 'urn:oasis:names:tc:xliff:document:1.2'
XLF = '{%s}' % XLF_NS

SDL_NS = 'http://sdl.com/FileTypes/SdlXliff/1.0'
SDL = '{%s}' % SDL_NS

MEM_NS = 'http://www.memsource.com/mxlf/2.0'
MEM = '{%s}' % MEM_NS


def extract_parts(filename, include_relationship_parts=False):
    """
    Extract parts from the file specified.
    :param filename: file from which to extract parts
    :type filename: str | unicode
    :param include_relationship_parts: True to include relationship parts
    :type include_relationship_parts: bool
    :return: dict of part names and their contents extracted
    :rtype: list[dict[str, str]]
    """
    filename = to_unicode(filename)
    parts = []
    try:
        with zipfile.ZipFile(filename) as zf:
            root = etree.fromstring(zf.read('[Content_Types].xml'))
            for name in [elem.get('PartName')[1:] for elem in root.iter(CT + 'Override')]:
                part = {'name': name, 'content': zf.read(name)}
                parts.append(part)
            if include_relationship_parts:
                for item in zf.infolist():
                    if item.filename.endswith('.rels'):
                        parts.append({'name': item.filename, 'content': zf.read(item.filename)})

    except zipfile.BadZipfile as e:
        print(e)

    return parts


def save_parts(parts, filename, new_filename):
    """
    Save parts in a new file.
    Any parts in the old file (filename) that do not exist in the specified parts (parts) are also added
    into the new file (new_filename)
    :param parts: parts to be saved in new_filename
    :type parts: list[dict[str, str]]
    :param filename: file to be merged into new_filename
    :type filename: str | unicode
    :param new_filename: file to which the parts will be saved
    :type new_filename: str | unicode
    :return:
    """
    filename = to_unicode(filename)
    new_filename = to_unicode(new_filename)
    file_dir = os.path.split(filename)[0]
    temp_file = os.path.join(file_dir, str(uuid.uuid4()))
    try:
        with zipfile.ZipFile(filename, 'r') as old_file:
            with zipfile.ZipFile(temp_file, 'w', zipfile.ZIP_DEFLATED) as new_file:
                for item in old_file.infolist():
                    if len([pt for pt in parts if pt['name'] == item.filename]) == 0:
                        data = old_file.read(item.filename)
                        new_file.writestr(item, data)

                for part in parts:
                    new_file.writestr(part['name'], part['content'])
    except zipfile.BadZipfile as e:
        print(e)
    if os.path.exists(new_filename):
        os.remove(new_filename)
    os.rename(temp_file, new_filename)


def to_unicode(obj, encoding='utf-8'):
    u"""
    Convert a basestring object to a unicode object.
    If the object is already unicode, return the object as-is.
    :type obj: basestring
    :param obj: object to convert
    :type encoding: str | unicode
    :param encoding: encoding used for unicode conversion
    :rtype obj: unicode | unknown
    :return obj: unicoded object
    """
    if isinstance(obj, basestring):
        if not isinstance(obj, unicode):
            obj = unicode(obj, encoding)
    return obj
