#  Working envrionment: Python 2.7
import re
from openpyxl import load_workbook
from openpyxl.styles import colors,PatternFill
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
from pdfminer.psparser import PSLiteral, literal_name
style = PatternFill(fill_type='solid',fgColor=colors.RED)



logic_texts = ['and', 'or']
verdict_texts = ['supported', 'not supported', 'present', 'not present or not active', 'disabled or not supported']


# multiple replace
def multi_replace(str, arg1, arg2, *args):
    str = str.replace(arg1, arg2)
    for arg in args:
        str = str.replace(arg, arg2)
    return str

#  extract data from pdf-acroform fields
def load_fields_from_pdf(field, T=''):
    #  Recursively load form fields
    form = field.get('Kids', None)
    t = field.get('T')
    if t is None:
        t = T
    else:
        #  Add its father name
        t = T + '.' + t if T != '' else t
    """ Following is to repeat fields that have "Kids", now is commented because 
    1. There could be multiple fileds who shared the same field name.
    2. For buttons, the parents has "V" value already, don't need to dig in Kids.
    """
    # if form and t:
    #     return [load_fields_from_pdf(resolve1(f), t) for f in form]
    # else:

    # Some field types, like signatures, need extra resolving
    value = resolve1(field.get('AS')) if resolve1(field.get('AS')) is not None else resolve1(field.get('V'))
    #  if output is PSLiteral type, transfer it into str type through "literal_name" function
    if isinstance(value, PSLiteral):
        return (t, literal_name(value))
    else:
        return (t, resolve1(value))

#  split data into dictionary
def split_data(field, d={}):
    flag = True if len(field) != 2 else False
    if flag:
        for f in field:
            split_data(f, d)
    elif isinstance(field[0], (tuple, list)):
        for f in field:
            split_data(f, d)
    else:
        key = field[0] if field[0] is not None else field[0]
        d[key] = field[1]
    return d

#  load ICS data from decrypted pdf docutment
def load_data_from_pdf(pdf):
    with open(pdf, 'rb') as file:
        parser = PDFParser(file)
        doc = PDFDocument(parser)
        parser.set_document(doc)
        outcome = [load_fields_from_pdf(resolve1(f)) for f in resolve1(doc.catalog['AcroForm'])['Fields']]
        # format the outcome of data extract from ics pdf
        outcome = split_data(outcome)
        if outcome['greater than 4 dynamic reader limit drl sets']:
            outcome['greater than 4 dynamic reader limit drl sets'] = True if int(outcome['greater than 4 dynamic reader limit drl sets']) > 4 else False
        if outcome['greater than 4 dynamic reader limit drl sets']:
            outcome['greater than 4 dynamic reader limit drl sets'] = True if int(outcome['greater than 4 dynamic reader limit drl sets']) > 4 else False
        for key in outcome:
            outcome[key].replace('Off', False).replace('No', False).replace('Yes', True)
        return split_data(outcome)

def readdata(ics_result):
    try:
        checklist = load_workbook('template_RGpath.xlsx')
    except IOError as e:
        raw_input(e)
        exit()
    temp_sheet = checklist['Sheet1'] # list all possible configurtion
    for path_sheet in [checklist['MSD Path(Online Only)'], checklist['MSD Path(Online Capable)'], checklist['qVSDC Path']]:
        for condition_cell in path_sheet['B']:
            condition = condition_cell.value  # get condition description from template
            outcome = ''  # contain logic method and verdict calculation the final result
            if condition not in [None, 'CONDITIONS ']:
                # format conditions
                condition = condition.lower()
                condition = multi_replace(condition, '\n', ' ', '   ', '  ')    # remove line break and make sure there's only one blank between words
                condition = multi_replace(condition, '[', '', ']', '(', ')')    # remove all parentheses
                condition = condition.replace('disbaled', 'disabled')         # fix TYPOs
                # translate conditions
                for item in temp_sheet['A']:
                    if item.value and re.search(item.value, condition):
                        if temp_sheet['B' + str(item.row)].value is None:
                            verdict = True
                        elif re.search(r'.* and .*', temp_sheet['B' + str(item.row)].value):
                            verdict = ''
                            pieces = temp_sheet['B' + str(item.row)].value.split(' and ')
                            for p in pieces:
                                verdict += ics_result[p]
                                verdict += ' and '
                            verdict = eval(verdict[:-5])
                        elif re.search(r'.* or .*', temp_sheet['B' + str(item.row)].value):
                            verdict = ''
                            pieces = temp_sheet['B' + str(item.row)].value.split(' or ')
                            for p in pieces:
                                verdict += ics_result[p]
                                verdict += ' or '
                            verdict = eval(verdict[:-4])
                        else:
                            verdict = ics_result[temp_sheet['B' + str(item.row)].value]



                        # and
                        #     piece = temp_sheet['B' + str(item.row)].value.split(' and ')





    #                     if re.search(item.value + r' (supported|not supported|present|not present or not active|disabled or not supported)', condition):
    #                         modifier = re.search(item.value + r' (supported|not supported|present|not present or not active|disabled or not supported)', condition).group(0)
    #                         modifier = modifier.replace(item.value, '').strip()
    #                         pass
    #                     else:
    #                         condition = condition.replace(item.value, item.value + ' supported')
    #                         modifier = 'supported'
    #                     if re.search(item.value + r' (supported|not supported|present|not present or not active|disabled or not supported) (and|or)', condition):
    #                         pass
    #                     else:
    #                         condition = re.sub(item.value + r' (supported|not supported|present|not present or not active|disabled or not supported)', r'\g<0> and', condition)
    #                     if modifier == 'supported' or modifier == 'present':
    #                         verdict = str(verdict) #
    #                     elif modifier == 'not supported':
    #                         verdict = str(not verdict) #
    #                     elif modifier == 'not present or not active' or modifier == 'disabled or not supported':
    #                         condition = re.sub(item.value + r' (supported|not supported|present|not present or not active|disabled or not supported) (and|or)', '', condition)
    #                         continue
    #                     condition = re.sub(item.value + r' (supported|not supported|present|not present or not active|disabled or not supported)', verdict, condition)
    #
    #             condition = condition.strip()
    #             condition = condition[:-4]
    #             condition = 'True' if not condition else eval(condition)
    #             subcondition_area = [sub.column for sub in path_sheet[condition_cell.row] if path_sheet[str(sub.column) + '2'].value == 'Y or N/A']
    #             result_area = [res.column for res in path_sheet[condition_cell.row] if path_sheet[str(res.column) + '2'].value == 'RESULT']
    #             for index, subcondition in enumerate(subcondition_area):
    #                 subcondition = True if path_sheet[subcondition + str(condition_cell.row)].value == 'Y' else False
    #                 path_sheet[result_area[index] + str(condition_cell.row)].value = 'PASS' if subcondition and condition else 'N/A'
    # checklist.save('checklist_output.xlsx')
    # checklist.close()

                        # condition = condition.replace(item.value.strip(), 'True')  # pending to add real result
                        #
                        # for verdict in verdict_texts:
                        #     print re.match(verdict, condition.strip())
                        #     print condition
                        #     raw_input()

                # if condition.strip():
                #     print condition.strip(), cell

                # divide = condition.replace('AND', ',').replace('OR', ',').replace('(', '').replace(')', '').lower()
                # divide = divide.replace(' not supported', 'n ').replace(' supported', 'y ').split(',')
                # divide = [item.strip() for item in divide]
                # for piece in divide:
                #     if piece[-1] == ']':
                #         divide[divide.index(piece)] += 'y'
                #         piece += 'y'
                #     if piece[piece.find(']') + 1] not in ['y', 'n']:
                #         addin = [p[-1] for p in divide[divide.index(piece):] if p[-2] in ['y', 'n']]
                #         if cell.row == 1058:
                #             print addin, piece
                #         divide[divide.index(piece)] += addin
                #         piece += addin
                    # #  check if all end with y or n
                    # if piece[-1] not in ['y', 'n']:
                    #     raw_input(piece)
    #                 for key in ics.keys():
    #                     if piece[:(len(piece) -1)] == key:
    #                         divide[divide.index(piece)] = ics[key] if piece[-1] == 'y' else not ics[key]
    #             condition = condition.lower().replace(' not supported', '').replace(' supported', '')
    #             for i in range(len(divide)):
    #                 if condition.find('[') != -1:
    #                     condition = condition.replace(condition[condition.find('['): condition.find(']') + 1], str(divide[i])) + '\n'
    #             checklist_sheet['C' + str(cell.row)] = 'PASS' if eval(condition) == True else 'NA'
    #         elif condition is not None and condition.strip() == 'NA':
    #             checklist_sheet['C' + str(cell.row)] = 'PASS'
    # checklist.save('checklist_output.xlsx')
    # checklist.close()



#define_ics_configuration()
#ics = get_ics_value()
#readdata()
ics = load_data_from_pdf('out1.pdf')
readdata(ics)
