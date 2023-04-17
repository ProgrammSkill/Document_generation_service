import os
from win32com.shell import shell, shellcon
from datetime import datetime
from openpyxl import *

path_file = shell.SHGetKnownFolderPath(shellcon.FOLDERID_Downloads)

def Date_conversion(date, type=None):
    arr_date = date.split('-')
    day = arr_date[0]
    arr_month = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
    month = arr_date[1]
    month = int(month)
    month_conversion = arr_month[month-1]
    year = arr_date[2]

    if type == 'quotes':
        if day[0] == '0':
            day = day[1]
        date = '\"' + day + '\" ' + month_conversion + ' ' + year + ' г.'
    elif type == 'word_month':
        if day[0] == '0':
            day = day[1]
        date = day + ' ' + month_conversion + ' ' + year + ' г.'
    else:
        date = day + '.' + arr_date[1] + '.' + year

    return date

def generation_about_arrival(request):
    document_type = request.POST.get('document_type')
    base = request.POST.get('base')
    obj_end_date = datetime.strptime(request.POST.get('end_date'), '%Y-%m-%d')
    end_date = obj_end_date.strftime("%d-%m-%Y")
    initiator = request.POST.get('initiator')
    person = request.POST.get('person')

    path_file_doc = 'document_generation_app/document_templates/about_arrival.xlsx'
    doc = load_workbook(path_file_doc)
    sheet = doc.active

    if document_type == 'visa':
        sheet["H36"] = 'X'
    elif document_type == 'resident_card':
        sheet["AJ36"] = 'X'
    elif document_type == 'temporary_residence_permit':
        sheet["BT36"] = 'X'
    elif document_type == 'temporary_residence_permit_for_the_purpose_of_education':
        sheet["DD36"] = 'X'

    if document_type != 'absent':
        series = request.POST.get('series')
        number = request.POST.get('number')
        if len(series) == 4:
            sheet["J40"] = series[0]
            sheet["N40"] = series[1]
            sheet["R40"] = series[2]
            sheet["V40"] = series[3]
        else:
            sheet["J40"] = series[0]
            sheet["N40"] = series[1]

        if len(number) == 6:
            sheet["AD40"] = number[0]
            sheet["AH40"] = number[1]
            sheet["AL40"] = number[2]
            sheet["AP40"] = number[3]
            sheet["AT40"] = number[4]
            sheet["AX40"] = number[5]


    global path_file
    path = path_file
    if os.path.exists(path + '/' + 'about_arrival.xlsx') == False:
        doc.save(path + '/' + 'about_arrival.xlsx')
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'about_arrival{i}.xlsx') == False:
                path = path_file + '/' + f'about_arrival{i}.xlsx'
                doc.save(path)
                break
            i += 1