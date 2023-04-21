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

# def Filling_cells(path_file_doc, name_sheet, value, row, list_columns):
#     doc = load_workbook(f'{path_file_doc}')
#     sheet = doc[f"{name_sheet}"]
#     list_symbols = list(value)
#     last_col = list_columns[len(list_columns) - 1]
#     row = row
#     index = 0
#     stop = False
#     for symbol in list_symbols:
#         for col in range(index, len(list_columns)):
#             cell = list_columns[index] + str(row)
#             print('test')
#             sheet[f'{cell}'] = symbol
#             index += 1
#             if col == f'{last_col}':
#                 stop = True
#             break
#         if stop == True:
#             break


def Generation_about_arrival(request):
    document_type = request.POST.get('document_type')
    purpose_departure = request.POST.get('purpose_departure')
    phone = request.POST.get('phone')
    job_title = request.POST.get('job_title').upper()
    list_job = list(job_title)
    receiving_side = request.POST.get('receiving_side')

    path_file_doc = 'document_generation_app/document_templates/about_arrival.xlsx'
    doc = load_workbook(path_file_doc)
    sheet = doc["sheet 1"]

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

    if phone != '':
        trans_table = {ord('('): None, ord(')'): None, ord(' '): None, ord('-'): None}
        phone = phone.translate(trans_table)
        sheet["CD46"] = phone[0]
        sheet["CH46"] = phone[1]
        sheet["CL46"] = phone[2]
        sheet["CP46"] = phone[3]
        sheet["CT46"] = phone[4]
        sheet["CX46"] = phone[5]
        sheet["DB46"] = phone[6]
        sheet["DF46"] = phone[7]
        sheet["DJ46"] = phone[8]
        sheet["DN46"] = phone[9]

    list_columns = ['R', 'V', 'Z', 'AD', 'AH', 'AL', 'AP', 'AT', 'AX', 'BB', 'BF', 'BJ', 'BN', 'BR', 'BV', 'BZ', 'CD',
                    'CH', 'CL', 'CP', 'CT', 'CX', 'DB', 'DF', 'DJ', 'DN']
    row = 48
    index = 0
    stop = False
    for symbol in list_job:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'DN':
                stop = True
            break
        if  stop == True:
             break


    sheet = doc["sheet 3"]
    if receiving_side == 'legal_entity':
        sheet["CP3"] = 'X'
    else:
        sheet["DN3"] = 'X'
        surname_receiving_side = request.POST.get('surname_receiving_side').upper()
        list_surname_receiving_side = list(surname_receiving_side)

        list_columns = ['N', 'R', 'V', 'Z', 'AD', 'AH', 'AL', 'AP', 'AT', 'AX', 'BB', 'BF', 'BJ', 'BN', 'BR', 'BV', 'BZ',
                        'CD', 'CH', 'CL', 'CP', 'CT', 'CX', 'DB', 'DF', 'DJ', 'DN']
        row = 5
        index = 0
        stop = False
        for symbol in list_surname_receiving_side:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'DN':
                    stop = True
                break
            if stop == True:
                break

        name_receiving_side = request.POST.get('name_receiving_side').upper()
        list_name_receiving_side = list(name_receiving_side)
        row = 7
        index = 0
        stop = False
        for symbol in list_name_receiving_side:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'DN':
                    stop = True
                break
            if stop == True:
                break


        patronymic_receiving_side = request.POST.get('patronymic_receiving_side').upper()
        list_patronymic_receiving_side = list(patronymic_receiving_side)
        list_columns = ['AH', 'AL', 'AP', 'AT', 'AX', 'BB', 'BF', 'BJ', 'BN', 'BR', 'BV', 'BZ','CD','CH', 'CL',
                        'CP', 'CT', 'CX', 'DB', 'DF', 'DJ', 'DN']
        row = 9
        index = 0
        stop = False
        for symbol in list_patronymic_receiving_side:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'DN':
                    stop = True
                break
            if stop == True:
                break

        type_of_identity_document = request.POST.get('type_of_identity_document').upper()
        list_type_of_identity_document = list(type_of_identity_document)
        list_columns = ['F', 'J', 'N', 'R', 'V', 'Z', 'AD', 'AH', 'AL', 'AP', 'AT']

        row = 11
        index = 0
        stop = False
        for symbol in list_type_of_identity_document:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'AT':
                    stop = True
                break
            if stop == True:
                break

        series_receiving_side = request.POST.get('series_receiving_side').upper()
        list_series_receiving_side = list(series_receiving_side)
        list_columns = ['BF', 'BJ', 'BN', 'BR']

        row = 11
        index = 0
        stop = False
        for symbol in list_series_receiving_side:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'BR':
                    stop = True
                break
            if stop == True:
                break

        number_receiving_side = request.POST.get('number_receiving_side').upper()
        list_number_receiving_side = list(number_receiving_side)
        list_columns = ['BZ', 'CD', 'CH', 'CL', 'CP', 'CT', 'CX', 'DB', 'DF', 'DJ', 'DN']

        row = 11
        index = 0
        stop = False
        for symbol in list_number_receiving_side:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'DN':
                    stop = True
                break
            if stop == True:
                break

        obj_date_issue_receiving_side = datetime.strptime(request.POST.get('date_issue_receiving_side'), '%Y-%m-%d')
        arr_date_issue_receiving_side = Date_conversion(obj_date_issue_receiving_side.strftime("%d-%m-%Y")).split('.')
        # day
        sheet["I13"], sheet["M13"] = arr_date_issue_receiving_side[0][0], arr_date_issue_receiving_side[0][1]
        # month
        sheet["Z13"], sheet["AD13"] = arr_date_issue_receiving_side[1][0], arr_date_issue_receiving_side[1][1]
        # year
        sheet["AL13"] = arr_date_issue_receiving_side[2][0]
        sheet["AP13"] = arr_date_issue_receiving_side[2][1]
        sheet["AT13"] = arr_date_issue_receiving_side[2][2]
        sheet["AX13"] = arr_date_issue_receiving_side[2][3]


        obj_sell_by_receiving_side = datetime.strptime(request.POST.get('sell_by_receiving_side'), '%Y-%m-%d')
        arr_sell_by_receiving_side = Date_conversion(obj_sell_by_receiving_side.strftime("%d-%m-%Y")).split('.')
        # day
        sheet["BN13"], sheet["BR13"] = arr_sell_by_receiving_side[0][0], arr_sell_by_receiving_side[0][1]
        # month
        sheet["CD13"], sheet["CH13"] = arr_sell_by_receiving_side[1][0], arr_sell_by_receiving_side[1][1]
        # year
        sheet["CP13"] = arr_sell_by_receiving_side[2][0]
        sheet["CT13"] = arr_sell_by_receiving_side[2][1]
        sheet["CX13"] = arr_sell_by_receiving_side[2][2]
        sheet["DB13"] = arr_sell_by_receiving_side[2][3]


        # Place of esidence
        region = request.POST.get('region').upper()
        list_region = list(region)
        list_columns = ['Z', 'AD', 'AH', 'AL', 'AP', 'AT', 'AX', 'BB', 'BF', 'BJ', 'BN', 'BR', 'BV', 'BZ',
                        'CD', 'CH', 'CL', 'CP', 'CT', 'CX', 'DB', 'DF', 'DJ', 'DN']
        row = 17
        index = 0
        for symbol in list_region:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                if list_columns[index] == 'BQ':
                    sheet[f'{cell}'] = symbol
                    row += 2
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break

        area = request.POST.get('area').upper()
        list_area = list(area)
        list_columns = ['V', 'Z', 'AD', 'AH', 'AL', 'AP', 'AT', 'AX', 'BB', 'BF', 'BJ', 'BN', 'BR', 'BV', 'BZ',
                        'CD', 'CH', 'CL', 'CP', 'CT', 'CX', 'DB', 'DF', 'DJ', 'DN']
        row = 21
        index = 0
        stop = False
        for symbol in list_area:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'DN':
                    stop = True
                break
            if stop == True:
                break

        city = request.POST.get('city').upper()
        list_city = list(city)
        list_columns = ['AD', 'AH', 'AL', 'AP', 'AT', 'AX', 'BB', 'BF', 'BJ', 'BN', 'BR', 'BV', 'BZ',
                        'CD', 'CH', 'CL', 'CP', 'CT', 'CX', 'DB', 'DF', 'DJ', 'DN']
        row = 23
        index = 0
        stop = False
        for symbol in list_city:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'DN':
                    stop = True
                break
            if stop == True:
                break

        street = request.POST.get('street').upper()
        list_street = list(street)
        list_columns = ['V', 'Z', 'AD', 'AH', 'AL', 'AP', 'AT', 'AX', 'BB', 'BF', 'BJ', 'BN', 'BR', 'BV', 'BZ',
                        'CD', 'CH', 'CL', 'CP', 'CT', 'CX', 'DB', 'DF', 'DJ', 'DN']
        row = 25
        index = 0
        stop = False
        for symbol in list_street:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'DN':
                    stop = True
                break
            if stop == True:
                break

        house = request.POST.get('house').upper()
        list_house = list(house)
        list_columns = ['J', 'N', 'R', 'V']
        row = 27
        index = 0
        stop = False
        for symbol in list_house:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'V':
                    stop = True
                break
            if stop == True:
                break

        frame = request.POST.get('frame').upper()
        list_frame = list(frame)
        list_columns = ['AH', 'AL', 'AP', 'AT', 'AX']
        row = 27
        index = 0
        stop = False
        for symbol in list_frame:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'AX':
                    stop = True
                break
            if stop == True:
                break

        structure = request.POST.get('structure').upper()
        list_structure = list(structure)
        list_columns = ['BN', 'BR', 'BW', 'BZ']
        row = 27
        index = 0
        stop = False
        for symbol in list_structure:
            for col in range(index, len(list_structure)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'BZ':
                    stop = True
                break
            if stop == True:
                break

        apartment = request.POST.get('apartment').upper()
        list_apartment = list(apartment)
        list_columns = ['CP', 'CT', 'CX', 'DB']
        row = 27
        index = 0
        stop = False
        for symbol in list_apartment:
            for col in range(index, len(list_structure)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'DB':
                    stop = True
                break
            if stop == True:
                break


    # for sheet in doc:
    #     print(sheet.title)

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