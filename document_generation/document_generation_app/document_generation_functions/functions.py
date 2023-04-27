from win32com.shell import shell, shellcon

def Get_path_file():
    return shell.SHGetKnownFolderPath(shellcon.FOLDERID_Downloads)

def Date_conversion_from_obj_date(validated_date):
    if validated_date.day < 9:
        day = '0' + str(validated_date.day)
    else:
        day = str(validated_date.day)

    if validated_date.day < 9:
        month = '0' + str(validated_date.month)
    else:
        month = str(validated_date.month)

    year = str(validated_date.year)

    return day + '-' + month + '-' + year

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
