import argparse
import sys
import zipfile
import xml.dom.minidom as dom
import odpmkd


def process_odp(in_fname, out_fname, is_remove_notes=False, is_remove_hidden=False):
    with zipfile.ZipFile(in_fname) as input_odp, \
         zipfile.ZipFile(out_fname, 'w', zipfile.ZIP_DEFLATED) as output_odp:
        for i in input_odp.infolist():
            if i.filename == 'content.xml':
                infile = input_odp.open(i)
                doc = dom.parseString(infile.read())
                changed = False
                if is_remove_notes:
                    remove_notes(doc)
                    changed = True
                if is_remove_hidden:
                    remove_hidden_slides(doc)
                    changed = True
                if changed:
                    output_odp.writestr(i.filename, doc.toxml())
                else:
                    output_odp.writestr(i.filename, infile.read())
            else:
                infile = input_odp.open(i)
                output_odp.writestr(i.filename, infile.read())


def remove_notes(doc):
    notes = doc.getElementsByTagName('presentation:notes')
    for note in notes:
        note.parentNode.removeChild(note)


def remove_hidden_slides(doc):
    hiden_page_styles = odpmkd.get_hidden_page_styles(doc)
    pages = doc.getElementsByTagName('draw:page')
    for page in pages:
        if page.attributes['draw:style-name'].value in hiden_page_styles:
            page.parentNode.removeChild(page)


def main(argv=None):
    if argv is None:
        argv = sys.argv

    argument_parser = argparse.ArgumentParser(description='OpenDocument Presentation Tools')
    argument_parser.add_argument('-i', '--input', required=True, help='ODP file to process')
    argument_parser.add_argument('-o', '--output', required=True, help='Path to the processed ODP file')
    argument_parser.add_argument('--remove_notes', required=False, help='Remove the notes from the presentation',
                                 action='store_true')
    argument_parser.add_argument('--remove_hidden_pages', required=False,
                                 help='Remove the hidden pages from the presentation', action='store_true')
    args = argument_parser.parse_args()

    if 'input' in args:
        process_odp(args.input, args.output, args.remove_notes, args.remove_hidden_pages)
    else:
        argument_parser.print_help()
        return

