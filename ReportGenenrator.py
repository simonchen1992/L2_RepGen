from openpyxl import load_workbook
from openpyxl.styles import colors,PatternFill

#  format the checklist: all logic signal inside [ICS Configuration] shall be lower case, otherwise shall be upper case
def format_and_or(string):
    output = string.lower()
    str_or = output.split(' or ')
    for i in str_or:
        if str_or[-1] != i:
            if i.rfind(']') > i.rfind('['):
                str_or[str_or.index(i)] += ' OR '
            else:
                str_or[str_or.index(i)] += ' or '
    output = ''.join(str_or)
    str_and = output.split(' and ')
    for i in str_and:
        if str_and[-1] != i:
            if i.rfind(']') > i.rfind('['):
                str_and[str_and.index(i)] += ' AND '
            else:
                str_and[str_and.index(i)] += ' and '
    output = ''.join(str_and)
    return output

#  format the checklist: parenthesis
def format_parenthesis(string):
    return False if string.count('[') != string.count(']') or string.count('(') != string.count(')') else True

#  get all non-repetitive [ICS Configuration] from checklist
def define_ics_configuration():
    list = []
    try:
        checklist = load_workbook('qVSDC_MSD_Reader_Checklist_v2.1.3c_Mar.xlsx')
    except IOError as e:
        raw_input(e)
        exit()
    for checklist_sheet in [checklist['MSD Path(Online Only)'], checklist['MSD Path(Online Capable)'], checklist['qVSDC Path']]:
        for cell in checklist_sheet['B']:
            string = cell.value
            if string not in ['NA', 'Test deleted', None]:
                if format_parenthesis(string) == False:
                    raw_input('Please modify the parenthesis in this row correctly:' + str(cell.row))
                string = string.replace('\n', ' ')
                string = format_and_or(string)
                # catch all configuration items
                while string.find('[') != -1:
                    key = string[string.find('['):(string.find(']') + 1)]
                    string = string[(string.find(']') + 1):]
                    if key.lower() not in list:
                        list.append(key.lower())
                        print key
    checklist.close()

#  fetch all [ICS Configuration] value from customer ics pdf document
def get_ics_value():
    dic = {}
    try:
        ics = load_workbook('checklist.xlsx')
    except IOError as e:
        raw_input(e)
        exit()
    ics_sheet = ics['ICS_Configuration']
    for cell in ics_sheet['A']:
        if cell.value is None:
            break
        dic[cell.value.lower()] = True
    return dic


def readdata(ics):
    try:
        checklist = load_workbook('checklist.xlsx')
    except IOError as e:
        raw_input(e)
        exit()
    checklist_sheet = checklist['checklist']
    for cell in checklist_sheet['B']:
        string = cell.value
        if string is not None and string.strip() not in ['NA', 'Test deleted']:
            string = string.replace('\n', ' ')
            string = format_and_or(string)

            divide = string.replace('AND', ',').replace('OR', ',').replace('(', '').replace(')', '').lower()
            divide = divide.replace(' not supported', 'n ').replace(' supported', 'y ').split(',')
            divide = [item.strip() for item in divide]
            for piece in divide:
                if piece[-1] == ']':
                    divide[divide.index(piece)] += 'y'
                    piece += 'y'
                if piece[piece.find(']') + 1] not in ['y', 'n']:
                    addin = [p[-1] for p in divide[divide.index(piece):] if p[-2] in ['y', 'n']]
                    if cell.row == 1058:
                        print addin, piece
                    divide[divide.index(piece)] += addin
                    piece += addin
                # #  check if all end with y or n
                # if piece[-1] not in ['y', 'n']:
                #     raw_input(piece)
                for key in ics.keys():
                    if piece[:(len(piece) -1)] == key:
                        divide[divide.index(piece)] = ics[key] if piece[-1] == 'y' else not ics[key]
            string = string.lower().replace(' not supported', '').replace(' supported', '')
            for i in range(len(divide)):
                if string.find('[') != -1:
                    string = string.replace(string[string.find('['): string.find(']') + 1], str(divide[i])) + '\n'
            checklist_sheet['C' + str(cell.row)] = 'PASS' if eval(string) == True else 'NA'
        elif string is not None and string.strip() == 'NA':
            checklist_sheet['C' + str(cell.row)] = 'PASS'
    checklist.save('checklist_output.xlsx')
    checklist.close()



define_ics_configuration()
#ics = get_ics_value()
#config = readdata(ics)