import argparse
import csv

import sys
import re

# pywin32
import win32com.client
from win32com.client import Dispatch, constants

from docx_parser_converter import ConversionConfig, docx_to_html


############ counter labels
EMAILS_SENT = 'emails_sent'


SUBJECT = 'subject'

SPAN_TAG_REGEX = re.compile(r'(<span[^<>]+>)([^<>]*)</span>')
MAIN_TAG_START_REGEX = re.compile(r'(<main( [^<>]*|)>)(<)')
MAIN_TAG_END_REGEX = re.compile(r'>(</main>)')


def create_email_to_outlook(to_email_address: str, subject: str, body_html: str, attachments: list = []):
    # https://itsec.media/post/python-send-outlook-email/

#    print(f'to_email_address="{to_email_address}"')

    const=win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
#    newMail.Subject = "I AM SUBJECT!!"
    newMail.Subject = subject
    # newMail.Body = "I AM\nTHE BODY MESSAGE!"
    newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
#    newMail.HTMLBody = "<HTML><BODY>Enter the <span style='color:red'>message</span> text here.</BODY></HTML>"
    newMail.HTMLBody = body_html
#    newMail.To = "email@demo.com"
    newMail.To = to_email_address
    #attachment1 = r"C:\Temp\example.pdf"
#    newMail.Attachments.Add(Source=attachment1)
    for attachment in attachments:
        newMail.Attachments.Add(Source=attachment)

#    newMail.display()
    #newMail.Send()
    #newMail.Save()

    return newMail


def simplify_html_styles(my_html: str) -> str:
    result_list = []
    simplified_span_text = ''
    simplified_span_tag = None
    pos = 0
    length = len(my_html)
    while pos < length and (re_result:=SPAN_TAG_REGEX.search(my_html, pos)) is not None:
        span_tag = re_result.group(1)
        span_text = re_result.group(2)
        if pos == re_result.start() and simplified_span_tag == span_tag:
            simplified_span_text += span_text
        else:
            if simplified_span_tag is not None:
                result_list.append(f'{simplified_span_tag}{simplified_span_text}</span>')
            
            result_list.append(my_html[pos: re_result.start()])
            simplified_span_tag = span_tag
            simplified_span_text = span_text

        pos = re_result.end()

    if simplified_span_tag is not None:
        result_list.append(f'{simplified_span_tag}{simplified_span_text}</span>')
    result_list.append(my_html[pos:])

    return ''.join(result_list)


def create_conversion_config():
    return ConversionConfig(paragraph_separator="\n\n")
        

# https://stackoverflow.com/questions/66604439/python-convert-word-to-html
def docx_to_html_str(docx_filename: str, config:ConversionConfig = create_conversion_config()) -> str:
#    docx_filename = "path_to_your_docx_file.docx"
    html_output = docx_to_html(docx_filename, config=config)

    html_output = simplify_html_styles(html_output)

    return html_output


def read_csv(csv_filename: str, delimiter: str = ';') -> list:
    result = []

    with open(csv_filename, 'r') as my_file:
        reader = csv.DictReader(my_file, delimiter=delimiter)
        for row in reader:
            result.append(row)

    return result


def get_counter_value(counters: dict, label: str) -> int:
    return counters.get(label, 0)


def increment_counter(counters: dict, label: str, increment: int = 1) -> int:
    result = get_counter_value(counters, label)

    result += increment
    counters[label] = result

    return result


def get_subject(same_person_subject: dict) -> str:
    return same_person_subject.get(SUBJECT, '')


def update_same_person_subject(same_person_subject: dict, new_subject: str):
    result = get_subject(same_person_subject)

    if len(result) == 0:
        same_person_subject[SUBJECT] = new_subject
    else:
        same_person_subject[SUBJECT] = f'{result} / {new_subject}'


def is_empty(elem) -> bool:
    return elem is None or len(elem) == 0


def send_or_draft_email(to_email_address: str, subject: str, body_html: str, send: bool, counters: dict):
    my_email = create_email_to_outlook(to_email_address, subject, body_html)

    my_email.display()
    if send:
        print(f'Sending email for: {to_email_address}')
        my_email.Send()
    else:
        print(f'e-mail to draft for: {to_email_address}')
        my_email.Save()

    increment_counter(counters, EMAILS_SENT)


def split_html(my_html: str) -> tuple:
    start_part = ''
    main_part = my_html
    end_part = ''

    re_result = MAIN_TAG_START_REGEX.search(my_html)
    if re_result is not None:
        assert (re_result2:= MAIN_TAG_END_REGEX.search(my_html, re_result.end())) is not None, \
               f'</main> not found in html:\n{my_html}'
        main_tag_start_end = re_result.end(1)
        main_tag_end_start = re_result2.start(1)
        start_part = my_html[:main_tag_start_end]
        main_part = my_html[main_tag_start_end:main_tag_end_start]
        end_part = my_html[main_tag_end_start:]

    return start_part, main_part, end_part


def join_htmls(result_html: str, html_to_add: str) -> str:
    start_part1, main_part1, end_part1 = split_html(result_html)
    start_part2, main_part2, end_part2 = split_html(html_to_add)

    if is_empty(start_part1):
        start_part1 = start_part2

    if is_empty(end_part1):
        end_part1 = end_part2

    result_html = f'{start_part1}{main_part1}{main_part2}{end_part1}'

    return result_html


def join_body_htmls(body_html_list: str, separator_html: str) -> str:
    my_body_html = ""
    if len(body_html_list) > 0:

        for body_html in body_html_list[:-1]:
            my_body_html = join_htmls(my_body_html, body_html)
            my_body_html = join_htmls(my_body_html, separator_html)

        body_html = body_html_list[-1]
        my_body_html = join_htmls(my_body_html, body_html)

    return my_body_html


def send_or_draft_email_list(to_email_address: str, subject: str, \
                             body_html_list: str, separator_html: str, \
                             send: bool, counters: dict):
    body_html = join_body_htmls(body_html_list, separator_html)

    send_or_draft_email(to_email_address, subject, body_html, send, counters)


def create_argsparser(program_name: str) -> argparse.ArgumentParser:
    result: ArgumentParser = argparse.ArgumentParser(prog=program_name)

    result.add_argument('-subject_col_name', nargs='?', required=True, type=str, help='Column name for the subject in csv')
    result.add_argument('-to_email_address_col_name', nargs='?', required=True, type=str, help='Column name for the to e-mail address in csv')
    result.add_argument('-email_body_template_docx', nargs='?', required=True, type=str, help="Email template in word's .docx format")
    result.add_argument('-body_separator_docx', nargs='?', required=True, type=str, help="Body in word's .docx format")
    result.add_argument('-data_csv', nargs='?', required=True, type=str, help="Input csv file name")
    result.add_argument('-send', action=argparse.BooleanOptionalAction, help="If you wanted to send the emails, instead of creating drafts, then use this flag")
    result.add_argument('-join_same_person_emails', action=argparse.BooleanOptionalAction, help="If you wantd to join same person emails, then use this flag")

    return result


def instance_template(template_str: str, labels: dict) -> str:
    result = template_str
    for label in labels:
        key = f'%%{label}%%'
        value = labels[label]

        result = result.replace(key, value)

    return result


def main():
    program_name = sys.argv[0]
    my_args = create_argsparser(program_name).parse_known_args(sys.argv)[0]

    subject_col_name: str = my_args.subject_col_name
    to_email_address_col_name: str = my_args.to_email_address_col_name
    email_body_template_docx: str = my_args.email_body_template_docx
    body_separator_docx: str = my_args.body_separator_docx
    data_csv: str = my_args.data_csv

    send: bool = my_args.send
    join_same_person_emails: bool = my_args.join_same_person_emails

    my_csv_rows = read_csv(data_csv)

    email_body_template_html = docx_to_html_str(email_body_template_docx)
    separator_html = docx_to_html_str(body_separator_docx)

    counters = {EMAILS_SENT: 0}
    csv_line = 2
    same_person_htmls = []
    same_person_subject = {}
    prev_to_email_address = None
    for csv_row in my_csv_rows:
        to_email_address = csv_row.get(to_email_address_col_name, None)
        subject = csv_row.get(subject_col_name, None)

        if is_empty(to_email_address) and is_empty(subject):
            if len(same_person_htmls) > 0:
                send_or_draft_email_list(prev_to_email_address, get_subject(same_person_subject), \
                                         same_person_htmls, separator_html, send, counters)
                same_person_subject = {}
                same_person_htmls = []
            prev_to_email_address = None
            continue
  
        assert not is_empty(to_email_address), f'{to_email_address_col_name} empty at line {csv_line} in "{data_csv}"'
        assert not is_empty(subject), f'{subject_col_name} empty at line {csv_line} in "{data_csv}"'

        to_email_address = to_email_address.strip()
        subject = subject.strip()

        single_email_html = instance_template(email_body_template_html, csv_row)

        if join_same_person_emails:
            if prev_to_email_address is not None:
                if prev_to_email_address != to_email_address:
                    if len(same_person_htmls) > 0:
                        send_or_draft_email_list(prev_to_email_address, get_subject(same_person_subject), \
                                                 same_person_htmls, separator_html, send, counters)
                        same_person_subject = {}
                        same_person_htmls = []
 
            update_same_person_subject(same_person_subject, subject)
            same_person_htmls.append(single_email_html)

        else:
            send_or_draft_email(to_email_address, subject, single_email_html, send, counters)

        prev_to_email_address = to_email_address
        csv_line += 1


    if len(same_person_htmls) > 0:
        send_or_draft_email_list(prev_to_email_address, get_subject(same_person_subject), \
                                 same_person_htmls, separator_html, send, counters)
        same_person_subject = {}
        same_person_htmls = []

    print()
    num_sent_emails = get_counter_value(counters, EMAILS_SENT)
    if send:
        print(f'{num_sent_emails} e-mails sent')
    else:
        print(f'{num_sent_emails} e-mails to draft')

if __name__ == "__main__":
    main()

