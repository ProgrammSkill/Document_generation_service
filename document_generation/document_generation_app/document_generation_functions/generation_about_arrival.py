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
        elif len(number) == 7:
            sheet["AD40"] = number[0]
            sheet["AH40"] = number[1]
            sheet["AL40"] = number[2]
            sheet["AP40"] = number[3]
            sheet["AT40"] = number[4]
            sheet["AX40"] = number[5]
            sheet["BB40"] = number[6]

        obj_date_issue = datetime.strptime(request.POST.get('date_issue'), '%Y-%m-%d')
        date_issue = obj_date_issue.strftime("%d-%m-%Y")
        date_issue = Date_conversion(date_issue)
        arr_date = date_issue.split('.')
        # day
        sheet['I31'], sheet['M31'] = arr_date[0][0], arr_date[0][1]
        # month
        sheet['Z31'], sheet['AD31'] = arr_date[1][0], arr_date[1][1]
        # year
        sheet['AL31'] = arr_date[2][0]
        sheet['AP31'] = arr_date[2][1]
        sheet['AT31'] = arr_date[2][2]
        sheet['AX31'] = arr_date[2][3]

        obj_date_sell_by = datetime.strptime(request.POST.get('sell_by'), '%Y-%m-%d')
        date_sell_by = obj_date_sell_by.strftime("%d-%m-%Y")
        date_sell_by = Date_conversion(date_sell_by)
        arr_date = date_sell_by.split('.')
        # day
        sheet['BN31'], sheet['BR31'] = arr_date[0][0], arr_date[0][1]
        # month
        sheet['CD31'], sheet['CH31'] = arr_date[1][0], arr_date[1][1]
        # year
        sheet['CP31'] = arr_date[2][0]
        sheet['CT31'] = arr_date[2][1]
        sheet['CX31'] = arr_date[2][2]
        sheet['DB31'] = arr_date[2][3]

    purpose_departure = request.POST.get('purpose_departure')
    if purpose_departure == 'official':
        sheet["AD44"] = 'X'
    elif purpose_departure == 'tourism':
        sheet["AQ44"] = 'X'
    elif purpose_departure == 'business':
        sheet["BD44"] = 'X'
    elif purpose_departure == 'studies':
        sheet["BO44"] = 'X'
    elif purpose_departure == 'job':
        sheet["CA44"] = 'X'
    elif purpose_departure == 'private':
        sheet["CN44"] = 'X'
    elif purpose_departure == 'transit':
        sheet["DB44"] = 'X'
    elif purpose_departure == 'humanitarian':
        sheet["AD46"] = 'X'
    else:
        sheet["AP46"] = 'X'

    phone = request.POST.get('phone')
    if phone != '':
        pass


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