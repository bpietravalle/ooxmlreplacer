#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import print_function
import argparse
import os.path
from copy import deepcopy

from lxml import etree

from utils import XML, A, to_unicode, extract_parts, save_parts


def __contains(find_what, text, match_case):
    if not match_case:
        find_what = find_what.upper()
        text = text.upper()
    return find_what in text


def __is_match(find_what, text, match_case):
    if not match_case:
        find_what = find_what.upper()
        text = text.upper()
    return find_what == text


def __get_plain_text(p):
    text = ''
    for r in p.iterchildren(A + 'r'):
        for t in r.iterchildren(A + 't'):
            text += t.text or ''
    return text


def __split_into_runs(p):
    for r in p.iterchildren(A + 'r'):
        r_props = None
        for r_child in r:
            if r_child.tag == A + 'rPr':
                r_props = r_child
            elif r_child.tag == A + 't' and r_child.text is not None:
                for char in r_child.text:
                    new_run = etree.Element(A + 'r')
                    if r_props is not None:
                        new_run.append(deepcopy(r_props))
                    t = etree.SubElement(new_run, A + 't')
                    t.text = char
                    if char == ' ':
                        t.set(XML + 'space', 'preserve')
                    r.addprevious(new_run)
            else:
                new_run = etree.Element(A + 'r')
                if r_props is not None:
                    new_run.append(deepcopy(r_props))
                new_run.append(deepcopy(r_child))
                r.addprevious(new_run)
        p.remove(r)


def __get_first_child(parent, tag):
    for p in parent.iterchildren(tag):
        return p
    return None


def __replace_runs(p, find_what, replace_with, match_case):
    while True:
        cont = False
        runs = [r for r in p if r.tag == A + 'r']
        for i in range(len(runs) - len(find_what) + 1):
            match = True
            for idx, char in enumerate(find_what):
                t = __get_first_child(runs[i + idx], A + 't')
                if t is None or t.text is None:
                    match = False
                    break
                if __is_match(char, t.text, match_case):
                    continue
                match = False
                break
            if match:
                if len(replace_with) > 0:
                    run_props = __get_first_child(runs[i], A + 'rPr')
                    new_run = etree.Element(A + 'r')
                    if run_props is not None:
                        new_run.append(deepcopy(run_props))
                    t = etree.SubElement(new_run, A + 't')
                    t.text = replace_with
                    if replace_with[0] == ' ' or replace_with[-1] == ' ':
                        t.set(XML + 'space', 'preserve')
                    runs[i].addnext(new_run)
                for idx in range(len(find_what)):
                    p.remove(runs[i + idx])
                cont = True
                break
        if not cont:
            break


def __merge_runs(p):
    while True:
        cont = False
        for run in p.iterchildren(A + 'r'):
            last_run = run.getprevious()
            if last_run is None or last_run.tag != A + 'r':
                continue
            run_props = __get_first_child(run, A + 'rPr')
            last_run_props = __get_first_child(last_run, A + 'rPr')
            if (run_props is None and last_run_props is not None) or (run_props is not None and last_run_props is None):
                continue
            if (run_props is None and last_run_props is None) or (
                        etree.tostring(run_props, encoding='utf-8', with_tail=False) == etree.tostring(last_run_props,
                                                                                                       encoding='utf-8',
                                                                                                       with_tail=False)):
                last_wt = __get_first_child(last_run, A + 't')
                wt = __get_first_child(run, A + 't')
                if last_wt is not None and wt is not None:
                    last_wt.text += wt.text or ''
                    if len(last_wt.text) > 0 and (last_wt.text[0] == ' ' or last_wt.text[-1] == ' '):
                        last_wt.set(XML + 'space', 'preserve')
                    run.tag = 'TO_BE_REMOVED'
                    cont = True
        etree.strip_elements(p, 'TO_BE_REMOVED')
        if not cont:
            break


def __replace_paragraph(paragraph, find_what, replace_with, match_case):
    text = __get_plain_text(paragraph)

    if __contains(find_what, text, match_case):
        print('OLD: ' + text)
        __split_into_runs(paragraph)
        __replace_runs(paragraph, find_what, replace_with, match_case)
        __merge_runs(paragraph)
        print('NEW: ' + __get_plain_text(paragraph), end='\n\n')


def __replace_part(part, find_what, replace_with, match_case):
    for p in part.iterdescendants(A + 'p'):
        __replace_paragraph(p, find_what, replace_with, match_case)

    return etree.tostring(part, encoding='utf-8', xml_declaration=True, standalone=True)


def replace(infile, outfile, find_what, replace_with, match_case=False):
    u"""Replace find_what with replace_with in pptx or pptm.
    :param infile: file in which replacement will be performed
    :type infile: str | unicode
    :param outfile: file in which the new content will be saved
    :type outfile: str | unicode
    :param find_what: text to search for
    :type find_what: str | unicode
    :param replace_with: text to replace find_what with
    :type replace_with: str | unicode
    :param match_case: True to make search case-sensitive
    :type match_case: bool
    """
    if not os.path.isfile(infile):
        raise ValueError('infile not found.')
    if not outfile:
        raise ValueError('outfile must be specified.')
    if not find_what:
        raise ValueError('find_what must be specified')

    if replace_with is None:
        replace_with = u''

    infile = to_unicode(infile)
    outfile = to_unicode(outfile)
    find_what = to_unicode(find_what)
    replace_with = to_unicode(replace_with)

    parts = extract_parts(infile)
    for part in parts:
        part['content'] = __replace_part(etree.fromstring(part['content']), find_what, replace_with, match_case)

    save_parts(parts, infile, outfile)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Find and Replace script for MS PowerPoint documents (pptx/pptm)')
    parser.add_argument('infile', metavar='INFILE', action='store', help='Specify the target file path.')
    parser.add_argument('outfile', metavar='OUTFILE', action='store', help='Specify the resultant file path')
    parser.add_argument('-f', '--findwhat', dest='find_what', required=True, action='store',
                        help='Specify characters to search for.')
    parser.add_argument('-r', '--replacewith', dest='replace_with', action='store',
                        help='Specify characters to replace with. ' +
                             'If nothing is specified, characters matched will be deleted from the resultant file.')
    parser.add_argument('-m', '--matchcase', dest='match_case', action='store_true', help='Case-sensitive search')
    args = parser.parse_args()
    replace(args.infile, args.outfile, args.find_what, args.replace_with, args.match_case)

