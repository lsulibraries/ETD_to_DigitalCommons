#! /usr/bin/env python3

import os
import csv
from collections import namedtuple
from datetime import datetime

import openpyxl


# these items' files were chpt-by-chpt pdfs, joined into one pdf post upload.
# we must be careful to ingest to Digital Commons the joined files instead of the split files.

joined_postupload = ('etd-06182004-122626', 'etd-09012004-114224', 'etd-0327102-091522', 'etd-0707103-142120',
                     'etd-0710102-054039', 'etd-0409103-184148', 'etd-04152004-142117', 'etd-0830102-145811',
                     'etd-0903103-141852', )


def parse_workbook(workbook_name):
    sourcepath = 'data/databasetables'
    filename = 'prod_etd_{}_database.xlsx'.format(workbook_name)
    fullpath = os.path.join(sourcepath, filename)
    return openpyxl.load_workbook(fullpath)


def list_all_sheets(workbook):
    # non-essential sanity check function
    sheets = [sheet for sheet in available_wb.get_sheet_names()]
    print(sheets)

def show_all_info(urn):
    # non-essential script that prints all the info related to a urn.
    print("Main example:", main_sheet[urn], "\n")
    print("Filenames example:", filenames_sheet[urn], "\n")
    print("Keywords example:", keywords_sheet[urn], "\n")
    print("Advisors example:", advisors_sheet[urn], "\n")
    print("Catalog example:", catalog_sheet[urn], "\n")


def parse_main_sheet():
    """ returns a dictionary in form of:
    {urn: NamedTuple
     urn: NamedTuple
    }
    NamedTuple is expected to have attributes: (urn first_name middle_name last_name suffix author_email 
                                                publish_email degree department dtype title abstract availability
                                                availability_description copyright_statement ddate sdate adate
                                                cdate rdate pid url notice notice_response timestamp
                                                survey_completed)
                                            or: (urn first_name middle_name last_name suffix author_email
                                                publish_email degree department dtype title abstract availability
                                                availability_description copyright_statement ddate sdate adate
                                                cdate rdate pid url notices timestamp)
    )
    """
    main_dict = dict()
    for wb in (available_wb, submitted_wb, withheld_wb):
        current_sheet = wb.get_sheet_by_name('etd_main table')
        for num, row in enumerate(current_sheet.iter_rows()):
            if num == 0:
                headers = (i.value for i in row)
                MainSheet = namedtuple('MainSheet', headers)
                continue
            values = (i.value for i in row)
            item = MainSheet(*values)
            main_dict[item.urn] = item
    return main_dict


def parse_filename_sheet():
    """returns a dictionary in form of:
    {urn: {filename: (path, size, available, description, page_count, timestamp),
           filename: (path, size, available, description, page_count, timestamp),}
     urn: {filename: (path, size, available, description, page_count, timestamp),}
    """
    filenames_sheet = dict()
    for wb in (available_wb, submitted_wb, withheld_wb):
        current_sheet = wb.get_sheet_by_name('filename_by_urn table')
        for num, row in enumerate(current_sheet.iter_rows()):
            if num == 0:
                continue
            [urn, filename, path, size, available, description, page_count, timestamp] = [i.value for i in row]
            file_dict = {filename: (path, size, available, description, page_count, timestamp)}
            if urn not in filenames_sheet:
                filenames_sheet[urn] = file_dict
            else:
                row_timestamp = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
                if filename in filenames_sheet[urn] and filenames_sheet[urn][filename][3] and filenames_sheet[urn][filename][3] != "NULL":
                    last_timestamp = datetime.strptime(filenames_sheet[urn][filename][3], "%Y-%m-%d %H:%M:%S")
                    if row_timestamp > last_timestamp:
                        filenames_sheet[urn][filename] = (path, size, available, description, page_count, timestamp)
                else:
                    filenames_sheet[urn][filename] = (path, size, available, description, page_count, timestamp)
    return filenames_sheet


def parse_keyword_sheet():
    """ returns a dictionary in form of:
    {urn: [(keyword, timestamp),
           (keyword, timestamp),
           ]}
    """
    keywords_sheet = dict()
    for wb in (available_wb, submitted_wb, withheld_wb):
        current_sheet = wb.get_sheet_by_name('keyword_by_urn table')
        for num, row in enumerate(current_sheet.iter_rows()):
            if num == 0:
                continue
            [keyword, urn, timestamp] = [i.value for i in row]
            if urn not in keywords_sheet:
                keywords_sheet[urn] = [(keyword, timestamp), ]
            else:
                keywords_sheet[urn].append((keyword, timestamp))
    return keywords_sheet


def parse_advisors_sheet():
    """ returns a dictionary in form of:
        {urn: {advisor: (advisor_title, advisor_email, approval, timestamp),
               advisor: (advisor_title, advisor_email, approval, timestamp),}
         urn: {advisor: (advisor_title, advisor_email, approval, timestamp),}
         }
    """
    advisors_sheet = dict()
    for wb in (available_wb, submitted_wb, withheld_wb):
        current_sheet = wb.get_sheet_by_name('advisor_by_urn table')
        for num, row in enumerate(current_sheet.iter_rows()):
            if num == 0:
                continue
            [urn, advisor_name, advisor_title, advisor_email, approval, timestamp] = [i.value for i in row]
            advisor_dict = {advisor_name: (advisor_title, advisor_email, approval, timestamp)}
            if urn not in advisors_sheet:
                advisors_sheet[urn] = advisor_dict
            else:
                row_timestamp = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
                if advisor_name in advisors_sheet[urn]:
                    last_timestamp = datetime.strptime(advisors_sheet[urn][advisor_name][3], "%Y-%m-%d %H:%M:%S")
                    if row_timestamp > last_timestamp:
                        advisors_sheet[urn][advisor_name] = (advisor_title, advisor_email, approval, timestamp)
                else:
                    advisors_sheet[urn][advisor_name] = (advisor_title, advisor_email, approval, timestamp)
    return advisors_sheet


def parse_catalog_sheet():
    catalog_sheet = dict()
    sourcepath = 'data/Catalogtables'
    sourcefile = 'CatalogETDSelectMetadata.csv'
    with open(os.path.join(sourcepath, sourcefile)) as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',')
        count = 0
        for num, row in enumerate(csvreader):
            if num == 0:
                continue
            # row = [i.strip('/').strip(']').strip('[').replace('])', '').replace('\n', '') for i in row]
            Title, Subtitle, AuthorFromTitleField, Author, SeriesDate, PubDate, URL = row
            urn = [i for i in URL.split('/') if 'etd-' in i]
            urn = urn[0]
            if not urn:
                print('No urn for URL:', URL)
            if urn not in catalog_sheet:
                catalog_sheet[urn] = [(Title, Subtitle, AuthorFromTitleField, Author, SeriesDate, PubDate, URL), ]
            else:
                catalog_sheet[urn].append((Title, Subtitle, AuthorFromTitleField, Author, SeriesDate, PubDate, URL))
                count += 1
    output_srt = ''
    for k, item in catalog_sheet.items():
        if len(item) > 1:
            output_srt += '\n{}\n'.format(k)
            for thing in item:
                output_srt += '\n'.join(thing)
                output_srt += '\n'
            output_srt += '\n\n'
    return catalog_sheet


available_wb = parse_workbook('available')
submitted_wb = parse_workbook('submitted')
withheld_wb = parse_workbook('withheld')

# merges the matching sheets from all 3 workbooks into one datastructure per sheet-type.
main_sheet = parse_main_sheet()
catalog_sheet = parse_catalog_sheet()
filenames_sheet = parse_filename_sheet()
keywords_sheet = parse_keyword_sheet()
advisors_sheet = parse_advisors_sheet()


