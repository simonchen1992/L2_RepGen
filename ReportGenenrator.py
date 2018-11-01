#  Working envrionment: Python 2.7
import re
from os import listdir,getcwd
from openpyxl import load_workbook
from openpyxl.styles import colors,PatternFill
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
from pdfminer.psparser import PSLiteral, literal_name

STYLE_RED = PatternFill(fill_type='solid',fgColor=colors.RED)
LOGIC_TEXT = r' (and|or)'
VERDICT_TEXT = r' (supported|not supported|present|not present or not active|disabled or not supported)'

# connect to template excel
def template_open():
    global expectResult, realResult
    try:
        expectResult = load_workbook('template_RGpath.xlsx')
        realResult = load_workbook('template_realResult.xlsx')
    except IOError as e:
        raw_input(e)
        exit()

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
    if outcome['Max Dynamic Reader Limit sets supported']:
        outcome['Max Dynamic Reader Limit sets supported'] = True if int(outcome['Max Dynamic Reader Limit sets supported']) > 4 else False
    if outcome['Product Configuration']:
        outcome['Product Configuration'] = True if outcome['Product Configuration'] == '(A) PCDA (IRWIN Reader) / S-ICR' else False
    for key in outcome:
        if outcome[key] == 'Yes':
            outcome[key] = True
        elif outcome[key] in ['Off', 'No']:
            outcome[key] = False
    return outcome

# Generate results for sub-condition
def ics_static_save(raw):
    temp_sheet = expectResult['ICS_Config_Static'] # list all possible configurtion
    for key in temp_sheet['B']:
        if key.value is None:
            verdict = 'True'
        elif re.search(r'.* and .*', key.value):
            verdict = ''
            pieces = key.value.split(' and ')
            for p in pieces:
                verdict += str(raw[p])
                verdict += ' and '
            verdict = eval(verdict[:-5])
        elif re.search(r'.* or .*', key.value):
            verdict = ''
            pieces = key.value.split(' or ')
            for p in pieces:
                verdict += str(raw[p])
                verdict += ' or '
            verdict = eval(verdict[:-4])
        else:
            verdict = raw[key.value] if key.value != "MSD Support" else (not raw[key.value])
        if verdict is None:
            raise Exception('There is blank to be filled at row %s. Please retry after fill it.'%key.row)
        temp_sheet['C' + str(key.row)] = str(verdict)

def gen_expect_result():
    temp_sheet = expectResult['ICS_Config_Static'] # list all possible configurtion
    for path_sheet in [expectResult['MSD Path(Online Only)'], expectResult['MSD Path(Online Capable)'], expectResult['qVSDC Path']]:
        for condition_cell in path_sheet['B']:
            condition = condition_cell.value  # get condition description from template
            #outcome = ''  # contain logic method and verdict calculation the final result
            if condition not in [None, 'CONDITIONS ']:
                if (temp_sheet['C2'].value == 'True' and path_sheet in [expectResult['MSD Path(Online Only)'], expectResult['MSD Path(Online Capable)']])\
                        or (temp_sheet['C3'].value == 'True' and path_sheet == expectResult['MSD Path(Online Only)'])\
                        or (temp_sheet['C4'].value == 'True' and path_sheet == expectResult['MSD Path(Online Capable)']):
                    condition = False
                else:
                    # format conditions
                    condition = condition.lower()
                    condition = multi_replace(condition, '\n', ' ', '   ', '  ')    # remove line break and make sure there's only one blank between words
                    condition = multi_replace(condition, '[', '', ']', '(', ')')    # remove all parentheses
                    condition = condition.replace('disbaled', 'disabled')         # fix TYPOs
                    # translate conditions
                    for item in temp_sheet['A']:
                        if item.value and re.search(item.value, condition):
                            if re.search(item.value + VERDICT_TEXT, condition):
                                modifier = re.search(item.value + VERDICT_TEXT, condition).group(0)
                                modifier = modifier.replace(item.value, '').strip()
                                pass
                            else:
                                condition = condition.replace(item.value, item.value + ' supported')
                                modifier = 'supported'
                            if re.search(item.value + VERDICT_TEXT + LOGIC_TEXT, condition):
                                pass
                            else:
                                condition = re.sub(item.value + VERDICT_TEXT, r'\g<0> and', condition)
                            if modifier == 'supported' or modifier == 'present':
                                verdict = temp_sheet['C' + str(item.row)].value
                            elif modifier == 'not supported':
                                verdict = True if temp_sheet['C' + str(item.row)].value == 'True' else False
                                verdict = str(not verdict)
                            elif modifier in ['not present or not active', 'disabled or not supported']:
                                condition = re.sub(item.value + VERDICT_TEXT + LOGIC_TEXT, '', condition)
                                continue
                            condition = re.sub(item.value + VERDICT_TEXT, verdict, condition)
                    condition = condition.strip()
                    condition = condition[:-4]
                    condition = True if not condition else eval(condition)
                subcondition_area = [sub.column for sub in path_sheet[condition_cell.row] if path_sheet[str(sub.column) + '2'].value == 'Y or N/A']
                result_area = [res.column for res in path_sheet[condition_cell.row] if path_sheet[str(res.column) + '2'].value == 'RESULT']
                for index, subcondition in enumerate(subcondition_area):
                    subcondition = True if path_sheet[subcondition + str(condition_cell.row)].value == 'Y' else False
                    if path_sheet[subcondition_area[index] + '1'].value == 'qVSDC/MSD active mode' and temp_sheet['C31'].value == 'False':
                        subcondition = False
                    path_sheet[result_area[index] + str(condition_cell.row)].value = 'PASS' if subcondition and condition else 'N/A'
    expectResult.save('expectResult.xlsx')
    expectResult.close()

def gen_real_result():
    titles = [t.value for t in realResult['Titles']['A']]  # list all possible titles
    f = [f.value for f in realResult['Titles']['B']]
    for path_sheet in [realResult['MSD Path(Online Only)'], realResult['MSD Path(Online Capable)'], realResult['qVSDC Path']]:
        for title in path_sheet['1']:
            if title.value in titles:
                f_1 = [f[i] for i,v in enumerate(titles)]
                raw_input(f_1)


    # Filepath = getcwd()
    # for Filename in listdir(Filepath):
    #     Filename = re.sub(r'\.\w+', '', Filename)
    #     print Filename

template_open()
# ics = load_data_from_pdf('out1.pdf')
# ics_static_save(ics)
# gen_expect_result()
gen_real_result()
