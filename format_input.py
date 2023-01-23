#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
import time
import argparse
import tempfile
import traceback
import xlsxwriter
import numpy as np
import pandas as pd
from colorama import init
init()

def menu(args):
    parser = argparse.ArgumentParser(description = "This script reads the exported (.csv|.txt) files from Scopus, Web of Science, PubMed, PubMed Central or Dimensions databases and turns each of them into a new file with an unique format. This script will ignore duplicated records.", epilog = "Thank you!")
    parser.add_argument("-t", "--type_file", choices = ofi.ARRAY_TYPE, required = True, type = str.lower, help = ofi.mode_information(ofi.ARRAY_TYPE, ofi.ARRAY_DESCRIPTION))
    parser.add_argument("-i", "--input_file", required = True, help = "Input file .csv or .txt")
    parser.add_argument("-o", "--output", help = "Output folder")
    parser.add_argument("--version", action = "version", version = "%s %s" % ('%(prog)s', ofi.VERSION))
    args = parser.parse_args()

    ofi.TYPE_FILE = args.type_file
    file_name = os.path.basename(args.input_file)
    file_path = os.path.dirname(args.input_file)
    if file_path is None or file_path == "":
        file_path = os.getcwd().strip()

    ofi.INPUT_FILE = os.path.join(file_path, file_name)
    if not ofi.check_path(ofi.INPUT_FILE):
        ofi.show_print("%s: error: the file '%s' doesn't exist" % (os.path.basename(__file__), ofi.INPUT_FILE), showdate = False, font = oscihub.YELLOW)
        ofi.show_print("%s: error: the following arguments are required: -i/--input_file" % os.path.basename(__file__), showdate = False, font = oscihub.YELLOW)
        exit()

    if args.output:
        output_name = os.path.basename(args.output)
        output_path = os.path.dirname(args.output)
        if output_path is None or output_path == "":
            output_path = os.getcwd().strip()

        ofi.OUTPUT_PATH = os.path.join(output_path, output_name)
        created = ofi.create_directory(ofi.OUTPUT_PATH)
        if not created:
            ofi.show_print("%s: error: Couldn't create folder '%s'" % (os.path.basename(__file__), ofi.OUTPUT_PATH), showdate = False, font = oscihub.YELLOW)
            exit()
    else:
        ofi.OUTPUT_PATH = os.getcwd().strip()
        ofi.OUTPUT_PATH = os.path.join(ofi.OUTPUT_PATH, 'output_format')
        ofi.create_directory(ofi.OUTPUT_PATH)

class FormatInput:

    def __init__(self):
        self.VERSION = 1.0

        self.INPUT_FILE = None
        self.TYPE_FILE = None
        self.OUTPUT_PATH = None

        self.ROOT_DIR = os.path.dirname(os.path.realpath(__file__))
        self.LOG_NAME = "run_%s_%s.log" % (os.path.splitext(os.path.basename(__file__))[0], time.strftime('%Y%m%d'))
        self.LOG_FILE = None

        # Menu
        self.TYPE_SCOPUS = "scopus"
        self.TYPE_WOS = "wos"
        self.TYPE_PUBMED = "pubmed"
        self.TYPE_PUBMED_CENTRAL = "pmc"
        self.TYPE_DIMENSIONS = "dimensions"
        self.TYPE_TXT = "txt"
        self.DESCRIPTION_SCOPUS = "Indicates that the file (.csv) was exported from Scopus"
        self.DESCRIPTION_WOS = "Indicates that the file (.csv) was exported from Web of Science"
        self.DESCRIPTION_PUBMED = "Indicates that the file (.csv) was exported from PubMed"
        self.DESCRIPTION_PUBMED_CENTRAL = "Indicates that the file (.txt) was exported from PubMed Central, necessarily in MEDLINE format"
        self.DESCRIPTION_DIMENSIONS = "Indicates that the file (.csv) was exported from Dimensions"
        self.DESCRIPTION_TXT = "Indicates that it is a text file (.txt)"
        self.ARRAY_TYPE = [self.TYPE_SCOPUS, self.TYPE_WOS, self.TYPE_PUBMED, self.TYPE_PUBMED_CENTRAL, self.TYPE_DIMENSIONS, self.TYPE_TXT]
        self.ARRAY_DESCRIPTION = [self.DESCRIPTION_SCOPUS, self.DESCRIPTION_WOS, self.DESCRIPTION_PUBMED, self.DESCRIPTION_PUBMED_CENTRAL, self.DESCRIPTION_DIMENSIONS, self.DESCRIPTION_TXT]

        # Scopus
        self.scopus_col_authors = 'Authors'
        self.scopus_col_title = 'Title'
        self.scopus_col_year = 'Year'
        self.scopus_col_doi = 'DOI'
        self.scopus_col_document_type = 'Document Type'
        self.scopus_col_language = 'Language of Original Document'
        self.scopus_col_cited_by = 'Cited by'
        # self.scopus_col_access_type = 'Access Type'
        # self.scopus_col_source = 'Source'

        # Web of Science (WoS)
        self.wos_col_authors = 'AU'
        self.wos_col_title = 'TI'
        self.wos_col_year = 'PY'
        self.wos_col_doi = 'DI'
        self.wos_col_document_type = 'DT'
        self.wos_col_language = 'LA'
        self.wos_col_cited_by = 'TC'

        # PubMed
        self.pubmed_col_authors = 'Authors'
        self.pubmed_col_title = 'Title'
        self.pubmed_col_year = 'Publication Year'
        self.pubmed_col_doi = 'DOI'
        self.pubmed_col_document_type = '' # Doesn't exist
        self.pubmed_col_language = '' # Doesn't exist
        self.pubmed_col_cited_by = '' # Doesn't exist

        # PubMed Central
        self.pmc_col_authors = 'Authors'
        self.pmc_col_title = 'Title'
        self.pmc_col_year = 'Publication Year'
        self.pmc_col_doi = 'DOI'
        self.pmc_col_document_type = 'Document Type'
        self.pmc_col_language = 'Language'
        self.pmc_col_cited_by = '' # Doesn't exist

        # Dimensions
        self.dimensions_col_authors = 'Authors'
        self.dimensions_col_title = 'Title'
        self.dimensions_col_year = 'PubYear'
        self.dimensions_col_doi = 'DOI'
        self.dimensions_col_document_type = 'Publication Type'
        self.dimensions_col_language = '' # Doesn't exist
        self.dimensions_col_cited_by = 'Times cited'

        # Xls Summary
        self.XLS_FILE = 'input_<type>.xlsx'
        self.XLS_SHEET_DETAIL = 'Detail'
        self.XLS_SHEET_WITHOUT_DOI = 'Without DOI'
        self.XLS_SHEET_DUPLICATES = 'Duplicates'

        # Xls Columns
        self.xls_col_item = 'Item'
        self.xls_col_title = 'Title'
        self.xls_col_year = 'Year'
        self.xls_col_doi = 'DOI'
        self.xls_col_document_type = 'Document Type'
        self.xls_col_language = 'Language'
        self.xls_col_cited_by = 'Cited By'
        self.xls_col_authors = 'Author(s)'

        self.xls_col_duplicate_type = 'Duplicate Type'
        self.xls_val_by_doi = 'By DOI'
        self.xls_val_by_title = 'By Title'

        self.xls_columns_csv = [self.xls_col_item,
                                self.xls_col_title,
                                self.xls_col_year,
                                self.xls_col_doi,
                                self.xls_col_document_type,
                                self.xls_col_language,
                                self.xls_col_cited_by,
                                self.xls_col_authors]

        self.xls_columns_txt = [self.xls_col_item,
                                self.xls_col_doi]

        # PubMed Central | MEDLINE
        self.MEDLINE_START = ['AB  -',
                              'AD  -',
                              'AID -',
                              'AU  -',
                              'AUID-',
                              'CN  -',
                              'DEP -',
                              'DP  -',
                              'FAU -',
                              'FIR -',
                              'GR  -',
                              'IP  -',
                              'IR  -',
                              'IS  -',
                              'JT  -',
                              'LA  -',
                              'LID -',
                              'MID -',
                              'OAB -',
                              'OABL-',
                              'PG  -',
                              'PHST-',
                              'PMC -',
                              'PMID-',
                              'PT  -',
                              'SO  -',
                              'TA  -',
                              'TI  -',
                              'VI  -']

        self.START_PMC = 'PMC -'
        self.START_PMID = 'PMID-'
        self.START_DATE = 'DEP -'
        self.START_TITLE = 'TI  -'
        self.START_ABSTRACT = 'AB  -'
        self.START_LANGUAGE = 'LA  -'
        self.START_PUBLICATION_TYPE = 'PT  -'
        self.START_JOURNAL_TYPE = 'JT  -'
        self.START_DOI = 'SO  -'
        self.START_AUTHOR = 'FAU -'

        self.param_pmc = 'pmc'
        self.param_pmid = 'pmid'
        self.param_date = 'data'
        self.param_title = 'title'
        self.param_language = 'language'
        self.param_abstract = 'abstract'
        self.param_publication_type = 'publication-type'
        self.param_journal_type = 'journal-type'
        self.param_doi = 'doi'
        self.param_author = 'author'

        # Fonts
        self.RED = '\033[31m'
        self.GREEN = '\033[32m'
        self.YELLOW = '\033[33m'
        self.BIRED = '\033[1;91m'
        self.BIGREEN = '\033[1;92m'
        self.END = '\033[0m'

    def show_print(self, message, logs = None, showdate = True, font = None):
        msg_print = message
        msg_write = message

        if font:
            msg_print = "%s%s%s" % (font, msg_print, self.END)

        if showdate is True:
            _time = time.strftime('%Y-%m-%d %H:%M:%S')
            msg_print = "%s %s" % (_time, msg_print)
            msg_write = "%s %s" % (_time, message)

        print(msg_print)
        if logs:
            for log in logs:
                if log:
                    with open(log, 'a', encoding = 'utf-8') as f:
                        f.write("%s\n" % msg_write)
                        f.close()

    def start_time(self):
        return time.time()

    def finish_time(self, start, message = None):
        finish = time.time()
        runtime = time.strftime("%H:%M:%S", time.gmtime(finish - start))
        if message is None:
            return runtime
        else:
            return "%s: %s" % (message, runtime)

    def create_directory(self, path):
        output = True
        try:
            if len(path) > 0 and not os.path.exists(path):
                os.makedirs(path)
        except Exception as e:
            output = False
        return output

    def check_path(self, path):
        _check = False
        if path:
            if len(path) > 0 and os.path.exists(path):
                _check = True
        return _check

    def mode_information(self, array1, array2):
        _information = ["%s: %s" % (i, j) for i, j in zip(array1, array2)]
        return " | ".join(_information)

    def read_txt_file(self):
        content = open(self.INPUT_FILE, 'r').readlines()

        collect_unique = {}
        collect_duplicate_doi = {}
        nr_doi = []
        index = 1
        for idx, line in enumerate(content, start = 1):
            line = line.strip()
            if line != '':
                flag_unique = False
                doi = line.lower()
                if doi not in nr_doi:
                    nr_doi.append(doi)
                    flag_unique = True

                collect = {}
                collect[self.xls_col_doi] = doi

                if flag_unique:
                    collect_unique.update({index: collect})
                    index += 1
                else:
                    collect[self.xls_col_duplicate_type] = self.xls_val_by_doi
                    collect_duplicate_doi.update({idx: collect})

        collect_papers = {self.XLS_SHEET_DETAIL: collect_unique,
                          self.XLS_SHEET_DUPLICATES: collect_duplicate_doi}

        return collect_papers

    def read_csv_file(self):
        _input_file = self.INPUT_FILE
        if self.TYPE_FILE == self.TYPE_SCOPUS:
            separator = ','
            _col_doi = self.scopus_col_doi
        elif self.TYPE_FILE == self.TYPE_WOS:
            separator = '\t'
            _col_doi = self.wos_col_doi
        elif self.TYPE_FILE == self.TYPE_PUBMED:
            separator = ','
            _col_doi = self.pubmed_col_doi
        elif self.TYPE_FILE == self.TYPE_PUBMED_CENTRAL:
            _input_file = self.read_medline_file(self.INPUT_FILE)
            separator = ','
            _col_doi = self.pmc_col_doi
        elif self.TYPE_FILE == self.TYPE_DIMENSIONS:
            separator = ','
            _col_doi = self.dimensions_col_doi

        df = pd.read_csv(filepath_or_buffer = _input_file, sep = separator, header = 0, index_col = False) # low_memory = False
        # df = df.where(pd.notnull(df), None)
        df = df.replace({np.nan: None})

        # Get DOIs
        collect_unique_doi = {}
        collect_duplicate_doi = {}
        collect_without_doi = {}
        nr_doi = []
        for idx, row in df.iterrows():
            flag_unique = False
            flag_duplicate_doi = False
            flag_without_doi = False

            doi = row[_col_doi]
            if doi:
                doi = doi.strip()
                doi = doi.lower()
                doi = doi[:-1] if doi.endswith('.') else doi
                if doi not in nr_doi:
                    nr_doi.append(doi)
                    flag_unique = True
                else:
                    flag_duplicate_doi = True
            else:
                flag_without_doi = True

            collect = {}
            if self.TYPE_FILE == self.TYPE_SCOPUS:
                collect[self.xls_col_authors] = row[self.scopus_col_authors].strip() if row[self.scopus_col_authors] else row[self.scopus_col_authors]
                collect[self.xls_col_title] = row[self.scopus_col_title].strip() if row[self.scopus_col_title] else row[self.scopus_col_title]
                collect[self.xls_col_year] = row[self.scopus_col_year]
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.scopus_col_document_type].strip() if row[self.scopus_col_document_type] else row[self.scopus_col_document_type]
                collect[self.xls_col_language] = row[self.scopus_col_language].strip() if row[self.scopus_col_language] else row[self.scopus_col_language]
                collect[self.xls_col_cited_by] = row[self.scopus_col_cited_by] if row[self.scopus_col_cited_by] else 0
            elif self.TYPE_FILE == self.TYPE_WOS:
                collect[self.xls_col_authors] = row[self.wos_col_authors].strip() if row[self.wos_col_authors] else row[self.wos_col_authors]
                collect[self.xls_col_title] = row[self.wos_col_title].strip() if row[self.wos_col_title] else row[self.wos_col_title]
                collect[self.xls_col_year] = row[self.wos_col_year]
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.wos_col_document_type].strip() if row[self.wos_col_document_type] else row[self.wos_col_document_type]
                collect[self.xls_col_language] = row[self.wos_col_language].strip() if row[self.wos_col_language] else row[self.wos_col_language]
                collect[self.xls_col_cited_by] = row[self.wos_col_cited_by] if row[self.wos_col_cited_by] else row[self.wos_col_cited_by]
            elif self.TYPE_FILE == self.TYPE_PUBMED:
                collect[self.xls_col_authors] = row[self.pubmed_col_authors].strip() if row[self.pubmed_col_authors] else row[self.pubmed_col_authors]
                collect[self.xls_col_title] = row[self.pubmed_col_title].strip() if row[self.pubmed_col_title] else row[self.pubmed_col_title]
                collect[self.xls_col_year] = row[self.pubmed_col_year]
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_language] = None
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_PUBMED_CENTRAL:
                collect[self.xls_col_authors] = row[self.pmc_col_authors].strip() if row[self.pmc_col_authors] else row[self.pmc_col_authors]
                collect[self.xls_col_title] = row[self.pmc_col_title].strip() if row[self.pmc_col_title] else row[self.pmc_col_title]
                collect[self.xls_col_year] = row[self.pmc_col_year]
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.pmc_col_document_type].strip() if row[self.pmc_col_document_type] else row[self.pmc_col_document_type]
                collect[self.xls_col_language] = row[self.pmc_col_language].strip() if row[self.pmc_col_language] else row[self.pmc_col_language]
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_DIMENSIONS:
                collect[self.xls_col_authors] = row[self.dimensions_col_authors].strip() if row[self.dimensions_col_authors] else row[self.dimensions_col_authors]
                collect[self.xls_col_title] = row[self.dimensions_col_title].strip() if row[self.dimensions_col_title] else row[self.dimensions_col_title]
                collect[self.xls_col_year] = row[self.dimensions_col_year]
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.dimensions_col_document_type].strip() if row[self.dimensions_col_document_type] else row[self.dimensions_col_document_type]
                collect[self.xls_col_language] = None
                collect[self.xls_col_cited_by] = row[self.dimensions_col_cited_by] if row[self.dimensions_col_cited_by] else row[self.dimensions_col_cited_by]

            if flag_unique:
                collect_unique_doi.update({idx + 1: collect})
            if flag_duplicate_doi:
                collect[self.xls_col_duplicate_type] = self.xls_val_by_doi
                collect_duplicate_doi.update({idx + 1: collect})
            if flag_without_doi:
                collect_without_doi.update({idx + 1: collect})

        # Get titles
        collect_unique = {}
        collect_duplicate_title = {}
        nr_title = []
        index = 1
        for idx, row in collect_unique_doi.items():
            flag_unique = False

            title = row[self.xls_col_title]
            if title:
                title = title.strip()
                title = title.lower()
                title = title[:-1] if title.endswith('.') else title
                if title not in nr_title:
                    nr_title.append(title)
                    flag_unique = True
            else:
                flag_unique = True

            if flag_unique:
                collect_unique.update({index: row})
                index += 1
            else:
                row[self.xls_col_duplicate_type] = self.xls_val_by_title
                collect_duplicate_title.update({idx: row})

        collect_duplicate = {}
        collect_duplicate = collect_duplicate_doi.copy()
        collect_duplicate.update(collect_duplicate_title)
        collect_duplicate = {item[0]: item[1] for item in sorted(collect_duplicate.items())}

        collect_papers = {self.XLS_SHEET_DETAIL: collect_unique,
                          self.XLS_SHEET_WITHOUT_DOI: collect_without_doi,
                          self.XLS_SHEET_DUPLICATES: collect_duplicate}

        return collect_papers

    def save_summary_xls(self, data_paper):

        def create_sheet(oworkbook, sheet_type, dictionary, styles_title, styles_rows):
            if self.TYPE_FILE == self.TYPE_TXT:
                _xls_columns = self.xls_columns_txt.copy()
            else:
                _xls_columns = self.xls_columns_csv.copy()

            if sheet_type == self.XLS_SHEET_DUPLICATES:
                _xls_columns.append(self.xls_col_duplicate_type)

            _last_col = len(_xls_columns) - 1

            worksheet = oworkbook.add_worksheet(sheet_type)
            worksheet.freeze_panes(row = 1, col = 0) # Freeze the first row.
            worksheet.autofilter(first_row = 0, first_col = 0, last_row = 0, last_col = _last_col) # 'A1:H1'
            worksheet.set_default_row(height = 14.5)

            # Add columns
            for icol, column in enumerate(_xls_columns):
                worksheet.write(0, icol, column, styles_title)

            # Add rows
            if self.TYPE_FILE == self.TYPE_TXT:
                worksheet.set_column(first_col = 0, last_col = 0, width = 7)  # Column A:A
                worksheet.set_column(first_col = 1, last_col = 1, width = 33) # Column B:B
                if sheet_type == self.XLS_SHEET_DUPLICATES:
                    worksheet.set_column(first_col = 2, last_col = 2, width = 19) # Column C:C
            else:
                worksheet.set_column(first_col = 0, last_col = 0, width = 7)  # Column A:A
                worksheet.set_column(first_col = 1, last_col = 1, width = 40) # Column B:B
                worksheet.set_column(first_col = 2, last_col = 2, width = 8)  # Column C:C
                worksheet.set_column(first_col = 3, last_col = 3, width = 33) # Column D:D
                worksheet.set_column(first_col = 4, last_col = 4, width = 18) # Column E:E
                worksheet.set_column(first_col = 5, last_col = 5, width = 12) # Column F:F
                worksheet.set_column(first_col = 6, last_col = 6, width = 11) # Column G:G
                worksheet.set_column(first_col = 7, last_col = 7, width = 36) # Column H:H
                if sheet_type == self.XLS_SHEET_DUPLICATES:
                    worksheet.set_column(first_col = 8, last_col = 8, width = 19) # Column I:I

            icol = 0
            for irow, (index, item) in enumerate(dictionary.items(), start = 1):
                col_doi = item[self.xls_col_doi]
                if sheet_type == self.XLS_SHEET_DUPLICATES:
                    duplicate_type = item[self.xls_col_duplicate_type]

                if self.TYPE_FILE == self.TYPE_TXT:
                    worksheet.write(irow, icol + 0, index, styles_rows)
                    worksheet.write(irow, icol + 1, col_doi, styles_rows)
                    if sheet_type == self.XLS_SHEET_DUPLICATES:
                        worksheet.write(irow, icol + 2, duplicate_type, styles_rows)
                else:
                    worksheet.write(irow, icol + 0, index, styles_rows)
                    worksheet.write(irow, icol + 1, item[self.xls_col_title], styles_rows)
                    worksheet.write(irow, icol + 2, item[self.xls_col_year], styles_rows)
                    worksheet.write(irow, icol + 3, col_doi, styles_rows)
                    worksheet.write(irow, icol + 4, item[self.xls_col_document_type], styles_rows)
                    worksheet.write(irow, icol + 5, item[self.xls_col_language], styles_rows)
                    worksheet.write(irow, icol + 6, item[self.xls_col_cited_by], styles_rows)
                    worksheet.write(irow, icol + 7, item[self.xls_col_authors], styles_rows)
                    if sheet_type == self.XLS_SHEET_DUPLICATES:
                        worksheet.write(irow, icol + 8, duplicate_type, styles_rows)

        workbook = xlsxwriter.Workbook(self.XLS_FILE)

        # Styles
        cell_format_title = workbook.add_format({'bold': True,
                                                 'font_color': 'white',
                                                 'bg_color': 'black',
                                                 'align': 'center',
                                                 'valign': 'vcenter'})
        cell_format_row = workbook.add_format({'text_wrap': True, 'valign': 'top'})

        create_sheet(workbook, self.XLS_SHEET_DETAIL, data_paper[self.XLS_SHEET_DETAIL], cell_format_title, cell_format_row)
        if self.TYPE_FILE != self.TYPE_TXT:
            create_sheet(workbook, self.XLS_SHEET_WITHOUT_DOI, data_paper[self.XLS_SHEET_WITHOUT_DOI], cell_format_title, cell_format_row)
        create_sheet(workbook, self.XLS_SHEET_DUPLICATES, data_paper[self.XLS_SHEET_DUPLICATES], cell_format_title, cell_format_row)

        workbook.close()

    def get_language(self, code):
        # https://www.nlm.nih.gov/bsd/language_table.html
        hash_data = {
            'afr': 'Afrikaans',
            'alb': 'Albanian',
            'amh': 'Amharic',
            'ara': 'Arabic',
            'arm': 'Armenian',
            'aze': 'Azerbaijani',
            'ben': 'Bengali',
            'bos': 'Bosnian',
            'bul': 'Bulgarian',
            'cat': 'Catalan',
            'chi': 'Chinese',
            'cze': 'Czech',
            'dan': 'Danish',
            'dut': 'Dutch',
            'eng': 'English',
            'epo': 'Esperanto',
            'est': 'Estonian',
            'fin': 'Finnish',
            'fre': 'French',
            'geo': 'Georgian',
            'ger': 'German',
            'gla': 'Scottish Gaelic',
            'gre': 'Greek, Modern',
            'heb': 'Hebrew',
            'hin': 'Hindi',
            'hrv': 'Croatian',
            'hun': 'Hungarian',
            'ice': 'Icelandic',
            'ind': 'Indonesian',
            'ita': 'Italian',
            'jpn': 'Japanese',
            'kin': 'Kinyarwanda',
            'kor': 'Korean',
            'lat': 'Latin',
            'lav': 'Latvian',
            'lit': 'Lithuanian',
            'mac': 'Macedonian',
            'mal': 'Malayalam',
            'mao': 'Maori',
            'may': 'Malay',
            'mul': 'Multiple languages',
            'nor': 'Norwegian',
            'per': 'Persian, Iranian',
            'pol': 'Polish',
            'por': 'Portuguese',
            'pus': 'Pushto',
            'rum': 'Romanian, Rumanian, Moldovan',
            'rus': 'Russian',
            'san': 'Sanskrit',
            'slo': 'Slovak',
            'slv': 'Slovenian',
            'spa': 'Spanish',
            'srp': 'Serbian',
            'swe': 'Swedish',
            'tha': 'Thai',
            'tur': 'Turkish',
            'ukr': 'Ukrainian',
            'und': 'Undetermined',
            'urd': 'Urdu',
            'vie': 'Vietnamese',
            'wel': 'Welsh'
        }

        r = 'Unknown'
        if code in hash_data:
            r = hash_data[code]

        return r

    def remove_endpoint(self, text):
        _text = text.strip()

        while(_text[-1] == '.'):
            _text = _text[0:len(_text) - 1]
            _text = _text.strip()

        return _text

    def block_continue(self, text):
        _continue = True
        for _start in self.MEDLINE_START:
            if text.startswith(_start):
                _continue = False
                break
        return _continue

    def get_data(self, text, array, start_param):
        if text.startswith(start_param):
            _line = text.replace(start_param, '').strip()
            array.append(_line)
            # continue

    def read_medline_file(self, file):

        def rename_publication_type(text):
            doc_type = None
            if text == 'Journal Article':
                doc_type = 'Article'
            elif text == 'Journal Article Case Report':
                doc_type = 'Case Report'
            elif text == 'Journal Article Editorial':
                doc_type = 'Editorial'
            elif text == 'Journal Article Letter':
                doc_type = 'Letter'
            elif text == 'Journal Article News':
                doc_type = 'News'
            elif text == 'Journal Article Review':
                doc_type = 'Review'
            else:
                doc_type = text
            return doc_type

        medline_data = {}
        with open(file, 'r', encoding = 'utf8') as fr:
            item_dict = {self.param_pmc: None,
                         self.param_pmid: None,
                         self.param_date: None,
                         self.param_title: None,
                         self.param_language: None,
                         self.param_abstract: None,
                         self.param_publication_type: None,
                         self.param_journal_type: None,
                         self.param_doi: None,
                         self.param_author: None}
            index = 1

            flag_start = False
            flag_title = False
            flag_abstract = False
            flag_doi = False

            arr_pmc = []
            arr_pmid = []
            arr_language = []
            arr_journal_type = []
            arr_publication_type = []
            arr_date = []
            arr_title = []
            arr_abstract = []
            arr_doi = []
            arr_author = []

            for line in fr:
                line = line.strip()
                if line:
                    # PMC
                    if line.startswith(self.START_PMC):
                        # Check
                        if arr_pmc:
                            _item_dict = item_dict.copy()
                            _item_dict.update({self.param_pmc: arr_pmc})
                            _item_dict.update({self.param_pmid: arr_pmid})
                            _item_dict.update({self.param_language: arr_language})
                            _item_dict.update({self.param_journal_type: arr_journal_type})
                            _item_dict.update({self.param_publication_type: arr_publication_type})
                            _item_dict.update({self.param_date: arr_date})
                            _item_dict.update({self.param_title: arr_title})
                            _item_dict.update({self.param_abstract: arr_abstract})
                            _item_dict.update({self.param_doi: arr_doi})
                            _item_dict.update({self.param_author: arr_author})
                            medline_data.update({index: _item_dict})
                            index += 1

                            flag_start = False
                            flag_title = False
                            flag_abstract = False
                            flag_doi = False

                            arr_pmc = []
                            arr_pmid = []
                            arr_language = []
                            arr_journal_type = []
                            arr_publication_type = []
                            arr_date = []
                            arr_title = []
                            arr_abstract = []
                            arr_doi = []
                            arr_author = []

                        flag_start = True
                        _line = line.replace(self.START_PMC, '').strip()
                        arr_pmc.append(_line)
                        continue

                    if flag_start:
                        self.get_data(line, arr_pmid, self.START_PMID)
                        self.get_data(line, arr_language, self.START_LANGUAGE)
                        self.get_data(line, arr_journal_type, self.START_JOURNAL_TYPE)
                        self.get_data(line, arr_publication_type, self.START_PUBLICATION_TYPE)
                        self.get_data(line, arr_date, self.START_DATE)
                        self.get_data(line, arr_author, self.START_AUTHOR)

                        # Title
                        if line.startswith(self.START_TITLE):
                            flag_title = True
                            _line = line.replace(self.START_TITLE, '').strip()
                            arr_title.append(_line)
                            continue
                        if flag_title:
                            if self.block_continue(line):
                                arr_title.append(line)
                                continue
                            else:
                                flag_title = False

                        # Abstract
                        if line.startswith(self.START_ABSTRACT):
                            flag_abstract = True
                            _line = line.replace(self.START_ABSTRACT, '').strip()
                            arr_abstract.append(_line)
                            continue
                        if flag_abstract:
                            if self.block_continue(line):
                                arr_abstract.append(line)
                                continue
                            else:
                                flag_abstract = False

                        # DOI
                        if line.startswith(self.START_DOI):
                            flag_doi = True
                            _line = line.replace(self.START_DOI, '').strip()
                            arr_doi.append(_line)
                            continue
                        if flag_doi:
                            if self.block_continue(line):
                                arr_doi.append(line)
                                continue
                            else:
                                flag_doi = False

            if arr_pmc:
                _item_dict = item_dict.copy()
                _item_dict.update({self.param_pmc: arr_pmc})
                _item_dict.update({self.param_pmid: arr_pmid})
                _item_dict.update({self.param_language: arr_language})
                _item_dict.update({self.param_journal_type: arr_journal_type})
                _item_dict.update({self.param_publication_type: arr_publication_type})
                _item_dict.update({self.param_date: arr_date})
                _item_dict.update({self.param_title: arr_title})
                _item_dict.update({self.param_abstract: arr_abstract})
                _item_dict.update({self.param_doi: arr_doi})
                _item_dict.update({self.param_author: arr_author})
                medline_data.update({index: _item_dict})
        fr.close()

        for index, item in medline_data.items():
            _publication_type = rename_publication_type(' '.join(item[self.param_publication_type]))
            item.update({self.param_pmc: ' '.join(item[self.param_pmc])})
            item.update({self.param_pmid: ' '.join(item[self.param_pmid])})
            item.update({self.param_journal_type: ' '.join(item[self.param_journal_type])})
            item.update({self.param_publication_type: _publication_type})
            item.update({self.param_title: ' '.join(item[self.param_title])})
            item.update({self.param_abstract: ' '.join(item[self.param_abstract])})
            item.update({self.param_author: '; '.join(item[self.param_author])})

            _language_raw = item[self.param_language]
            _language = []
            for code in _language_raw:
                _language.append(self.get_language(code))
            item.update({self.param_language: ' '.join(_language)})

            _date = ' '.join(item[self.param_date])
            if _date:
                _date = _date[0:4]
            item.update({self.param_date: _date})

            _doi_raw = ' '.join(item[self.param_doi])
            _doi_raw = _doi_raw.split('doi:')
            _doi = ''
            if len(_doi_raw) > 1:
                _doi = self.remove_endpoint(_doi_raw[1])
            item.update({self.param_doi: _doi})

        # Temporary file .csv
        fw_tmp = tempfile.NamedTemporaryFile(mode = 'w+t',
                                             encoding = 'utf-8',
                                             prefix = 'medline_output_',
                                             suffix = '.csv')

        fw_tmp.write('"%s","%s","%s","%s","%s","%s","%s","%s","%s"\n' % ('PMID',
                                                                         self.pmc_col_title,
                                                                         self.pmc_col_authors,
                                                                         self.pmc_col_year,
                                                                         'PMCID',
                                                                         self.pmc_col_doi,
                                                                         self.pmc_col_language,
                                                                         self.pmc_col_document_type,
                                                                         'Journal Type'))
        for _, detail in medline_data.items():
            fw_tmp.write('"%s","%s","%s","%s","%s","%s","%s","%s","%s"\n' % (detail[self.param_pmid],
                                                                             detail[self.param_title],
                                                                             detail[self.param_author],
                                                                             detail[self.param_date],
                                                                             detail[self.param_pmc],
                                                                             detail[self.param_doi],
                                                                             detail[self.param_language],
                                                                             detail[self.param_publication_type],
                                                                             detail[self.param_journal_type],
                                                                             # detail[self.param_abstract]
                                                                             ))
        fw_tmp.seek(0)
        # fw_tmp.close()

        return fw_tmp

def main(args):
    try:
        start = ofi.start_time()
        menu(args)

        ofi.LOG_FILE = os.path.join(ofi.OUTPUT_PATH, ofi.LOG_NAME)
        ofi.XLS_FILE = os.path.join(ofi.OUTPUT_PATH, ofi.XLS_FILE.replace('<type>', ofi.TYPE_FILE))
        ofi.show_print("#############################################################################", [ofi.LOG_FILE], font = ofi.BIGREEN)
        ofi.show_print("############################### Format Input ################################", [ofi.LOG_FILE], font = ofi.BIGREEN)
        ofi.show_print("#############################################################################", [ofi.LOG_FILE], font = ofi.BIGREEN)

        # Read input file
        input_information = {}
        if ofi.TYPE_FILE == ofi.TYPE_TXT:
            ofi.show_print("Reading the .txt file", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_txt_file()
        elif ofi.TYPE_FILE == ofi.TYPE_SCOPUS:
            ofi.show_print("Reading the .csv file from Scopus", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_WOS:
            ofi.show_print("Reading the .csv file from Web of Science", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_PUBMED:
            ofi.show_print("Reading the .csv file from PubMed", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_PUBMED_CENTRAL:
            ofi.show_print("Reading the .txt file from PubMed Central", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_DIMENSIONS:
            ofi.show_print("Reading the .csv file from Dimensions", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        ofi.show_print("Input file: %s" % ofi.INPUT_FILE, [ofi.LOG_FILE])
        ofi.show_print("", [ofi.LOG_FILE])

        ofi.save_summary_xls(input_information)
        ofi.show_print("Output file: %s" % ofi.XLS_FILE, [ofi.LOG_FILE], font = ofi.GREEN)

        ofi.show_print("", [ofi.LOG_FILE])
        ofi.show_print(ofi.finish_time(start, "Elapsed time"), [ofi.LOG_FILE])
        ofi.show_print("Done!", [ofi.LOG_FILE])
    except Exception as e:
        ofi.show_print("\n%s" % traceback.format_exc(), [ofi.LOG_FILE], font = ofi.RED)
        ofi.show_print(ofi.finish_time(start, "Elapsed time"), [ofi.LOG_FILE])
        ofi.show_print("Done!", [ofi.LOG_FILE])

if __name__ == '__main__':
    ofi = FormatInput()
    main(sys.argv)
