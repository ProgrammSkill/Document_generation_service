from win32com.shell import shell, shellcon

def Get_path_file():
    return shell.SHGetKnownFolderPath(shellcon.FOLDERID_Downloads)

def Date_conversion_from_obj_date(validated_date):
    if validated_date.day < 9:
        day = '0' + str(validated_date.day)
    else:
        day = str(validated_date.day)

    if validated_date.month < 9:
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

def FirstNameDeclension(first_name):
    ending = first_name[len(first_name) - 1]

    # FEMALE
    if ending == 'а':
        ending = 'ы'
    elif ending == 'я' or first_name[len(first_name) - 2] + ending == 'ль':
        ending = 'и'
    elif ending == 'т':
        ending = 'ты'
    # MALE
    elif ending == 'н' or ending == 'к' or ending == 'т' or ending == 'д' or ending == 'с' or ending == 'г' or \
            ending == 'м' or ending == 'р' or ending == 'в' or ending == 'б' or ending == 'л':
        ending = 'а'
        return first_name + ending
    elif ending == 'й':
        ending = 'я'
    elif ending == 'ь':
        ending = 'я'
    else:
        return first_name

    first_name = first_name[:len(first_name) - 1] + ending
    return first_name


def SurnameDeclension(surname):
    ending = surname[len(surname) - 1]

    # FEMALE
    if ending == 'а' or ending == 'я':
        surname = surname[:len(surname) - 1] + 'ой'
        return surname
    elif ending == 'ь':
        surname[len(surname) - 1] = 'й'
        return surname
    # MALE
    elif ending == 'в':
        surname = surname + 'а'
        return surname
    elif ending == 'й':
        surname = surname[:len(surname) - 2] + 'ского'
        return surname
    else:
        return surname

def LastNameDeclension(last_name):
    ending = last_name[len(last_name) - 1]

    if ending == 'а':
        return last_name[:len(last_name) - 1] + 'ы'
    elif ending == 'ч':
        return last_name + 'а'
    else:
        return last_name

def CountryDeclination(country):
    ending = country[len(country) - 1]

    if ending == 'я':
        country = country[:len(country) - 1] + 'и'
    elif ending == 'ы':
        country = country[:len(country) - 1]
    elif (country[len(country) - 2] == 'н' or country[len(country) - 2] == 'м') and ending == 'а':
        country = country[:len(country) - 1] + 'ы'
    elif country[len(country) - 2] == 'к' and ending == 'а':
        country = country[0:len(country)-1] + 'и'
    elif country[len(country) - 2] == 'ш' and ending == 'а':
        country = country[:len(country) - 1] + 'и'
    elif country[len(country) - 2] == 'д' and ending == 'а':
        country = country[:len(country) - 1] + 'ы'
    elif ending == 'й' or ending == 'ь':
        country = country[:len(country) - 1] + 'я'
    elif ending == 'н' or ending == 'к' or ending == 'т' or ending == 'а' or ending == 'л' or ending == 'с' or \
            ending == 'д' or ending == 'р' or ending == 'м' or ending == 'з':
        country = country[:len(country)] + 'а'

    return country

print(CountryDeclination('Казахстан'))