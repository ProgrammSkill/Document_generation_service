import os
import pymssql as pymssql
import requests
from django.core.management.commands import shell
from django.http import HttpResponse
from django.shortcuts import render, redirect
from .forms import *
from docx import Document
from docxtpl import DocxTemplate
from docx.shared import *
from win32com.shell import shell, shellcon
from datetime import datetime
from openpyxl import *
import pytz

# import function of document generations
from document_generation_app.document_generation_functions.generation_about_arrival import generation_about_arrival

path_file = shell.SHGetKnownFolderPath(shellcon.FOLDERID_Downloads)



def CompanyAPI():
    # id_company = "b5deb54b-01bd-4bc3-87d0-21359b046e2a"
    # response = requests.get("http://secretochka.ru:48910/company-service/api/v1/manager/companyes/" + id_company).json()
    # actual_address = requests.get("http://secretochka.ru:48910/company-service/api/v1/companyes/" + id_company + "/info/actualaddresses").json()
    # response["ActualAddreses"] = actual_address
    response = {
            "id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
            "inn": "7804525530",
            "kpp": "780401001",
            "name": "Капитал Кадры",
            "organizationalForm": "ООО",
            "ogrn": "1117746358608",
            "okved": "123123123",
            "legalAddress": {
                "name": "195299, Г.Санкт-Петербург, пр-кт Гражданский, д. 119 ЛИТЕР А, офис 8"
            },
            "contactInfo": {
                "id": 0,
                "name": "Олег",
                "phone": "+78989321223",
                "email": "user@example.com"
            },
            "bank": {
                "id": 0,
                "bankId": "12341234",
                "correspondentAccount": "12345123451234512345",
                "paymentAccount": "12345123451234512345"
            },
            "ActualAddreses": [{"name": "195299, Г.Санкт-Петербург, пр-кт Гражданский, д. 119 ЛИТЕР А, офис 8"}]
    }

    return response

def IndividualAPI():
    pass

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


# Create your views here.
def index(request):
    return render(request,'document_generation_app/index.html')

def Employment_contract_Document(request):
    return render(request,'document_generation_app/employment_contract.html')


def Generate_Employment_contract_Document(request):
    if request.method =='POST':
        company = CompanyAPI()
        organization = company["organizationalForm"] + ' "' + company["name"] + '"'
        phone = company["contactInfo"]["phone"]
        inn = company['inn']
        kpp = company['kpp']
        ogrn = company['ogrn']
        paymentAccount = company['bank']['paymentAccount']
        correspondentAccount = company['bank']['correspondentAccount']
        legalAddress = company['legalAddress']["name"]
        ActualAddreses = company['ActualAddreses'][0]["name"]
        CEO = 'Сталюковой Екатерина Александровны'

        individual = 'Гайназаров Кайратбек'
        passport_series = 'AC'
        passport_number = '4348554'

        number = request.POST.get('number')
        job_title = request.POST.get('job_title')
        salary = request.POST.get('salary')
        obj_start_date = datetime.strptime(request.POST.get('start_date'), '%Y-%m-%d')
        start_date = obj_start_date.strftime("%d-%m-%Y")
        obj_end_date = datetime.strptime(request.POST.get('end_date'), '%Y-%m-%d')
        end_date = obj_end_date.strftime("%d-%m-%Y")
        cause = request.POST.get('cause')
        start_time = request.POST.get('start_time')
        if start_time[0] == "0":
            start_time = start_time[1:]
        end_time = request.POST.get('end_time')
        if end_time[0] == "0":
            end_time = end_time[1:]

        path_file_doc = 'document_generation_app/document_templates/employment_contract.docx'
        doc = DocxTemplate(path_file_doc)


        context = {
            'organization': organization,
            'number': number,
            'job_title': job_title,
            'salary': salary,
            'startDateQuotes': Date_conversion(start_date, 'quotes'),
            'startDateWordMonth': Date_conversion(start_date, 'word_month'),
            'startDateStandart': Date_conversion(start_date),
            'endDateStandart': Date_conversion(end_date),
            'startTime': start_time,
            'endTime': end_time,
            'cause': cause,
            'inn': inn,
            'kpp': kpp,
            'phone': phone,
            'legalAddress': legalAddress,
            'ActualAddreses': ActualAddreses,
            'paymentAccount': paymentAccount,
            'correspondentAccount': correspondentAccount,
            'CEO': CEO,
            'individual': individual
        }

        doc.render(context)

        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'employment_contract.docx') == False:
            doc.save(path + '/' + 'employment_contract.docx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'employment_contract{i}.docx') == False:
                    path = path_file + '/' + f'employment_contract{i}.docx'
                    doc.save(path)
                    break
                i += 1

    return redirect('employment_contract')



def GPC_Agreement(request):
    return render(request,'document_generation_app/gpc_contract.html')


def Generate_GPC_Document(request):
    if request.method =='POST':
        company = CompanyAPI()
        inn = company['inn']
        kpp = company['kpp']
        ogrn = company['ogrn']
        legalAddress = company['legalAddress']["name"]
        organization = company["organizationalForm"] + " " + company["name"]
        CEO = 'Сталюковой Екатерина Александровны'
        individual = 'Гайназаров Кайратбек'
        passport_series = 'AC'
        passport_number = '4348554'

        arr_value_table = []


        number = request.POST.get('number')
        obj_start_date = datetime.strptime(request.POST.get('start_date'), '%Y-%m-%d')
        start_date = Date_conversion(obj_start_date.strftime("%d-%m-%Y"), 'word_month')
        obj_end_date = datetime.strptime(request.POST.get('end_date'), '%Y-%m-%d')
        end_date = Date_conversion(obj_end_date.strftime("%d-%m-%Y"), 'word_month')
        address = request.POST.get('address')
        table_doc = request.POST.get('table_doc')

        # Getting input from a page
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(11)

        doc.add_paragraph("").add_run("ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ   № " + number).bold = True
        doc.add_paragraph(organization + ", именуемое в дальнейшем «Заказчик», в лице  Генерального директора " + CEO + ", действующей на основании Устава, с одной стороны, и гражданин Киргизии Гайназаров Кайратбек,именуемый в дальнейшем «Исполнитель», с другой стороны, вместе именуемые «Стороны», заключили настоящий договор сроком до " + end_date + ", в дальнейшем именуемый «Договор», о нижеследующем:")

        doc.add_paragraph("").add_run("1.1 ПРЕДМЕТ ДОГОВОРА").bold = True
        doc.add_paragraph("1.1 Исполнитель обязуется по заданию Заказчика лично оказать услуги указанные в Приложении №1 к данному договору по адресу: Ленинградская область город Каменногорск ул Заозерная дом 1 ,  а Заказчик обязуется эти услуги оплатить.\n" +
        "1.2  Услуги считаются оказанными после подписания акта приема-сдачи Услуг Заказчиком или его уполномоченным представителем")

        doc.add_paragraph("").add_run("2.  ПРАВА И ОБЯЗАННОСТИ СТОРОН").bold = True

        doc.add_paragraph("").add_run("2.1 ЗАКАЗЧИК ОБЯЗАН:").bold = True
        doc.add_paragraph("2.1.1. Предоставить точную информацию Исполнителю о времени и месте оказания Услуг. В случае изменения времени и места оказания Услуг, Заказчик обязан сообщить об этом Исполнителю не менее чем за два часа. Если Заказчик не сообщил об отмене заказа или сообщил менее чем за два часа, Исполнитель имеет право взыскать неустойку (штраф) в размере стоимости проезда.\n" +
        "2.1.2. Обеспечить присутствие своего представителя на объекте оказания услуг в случае, если иное не оговорено с Исполнителем.\n" +
        "2.1.3.  По окончанию оказания услуг подписать Акт приема-сдачи услуг.\n" +
        "2.1.4.  Оплатить услуги Исполнителя на основании подписанного акта приема-передачи в соответствии с положениями п.3.2 настоящего Договора.\n" +
        "2.1.5. Вместе с Исполнителем составить Перечень имущества, перемещаемого в процессе оказания услуг, если такое действие потребовано Исполнителем.")

        doc.add_paragraph("").add_run("2.2. ИСПОЛНИТЕЛЬ ОБЯЗАН:").bold = True
        doc.add_paragraph("2.2.1.Оказать услуги надлежащего качества с соблюдением требований техники безопасности.\n" +
        "2.2.2.  Оказать услуги в полном объеме и в срок, согласованный с Заказчиком.\n" +
	    "2.2.3.  Без опозданий приступить к оказанию услуг.\n" +
	    "2.2.4.  Не допускать простой транспортных средств, в случае предоставления их Заказчиком.\n" +
	    "2.2.5.  Составить вместе с Заказчиком Перечень имущества, перемещаемого в процессе оказания услуг, если такое действие потребовано Заказчиком.\n" +
	    "2.2.6.  Вернуть материальные ценности, полученные от Заказчика для выполнения услуг.\n" +
	    "2.2.7.  Соблюдать требования по охране труда, пожарной безопасности, производственной санитарии, иные локальные нормативные акты, в том числе приказы (распоряжения), инструкции и т. д., действующие на территории выполнения им работ, оказания услуг.\n" +
	    "2.2.8.  При возникновении ситуации, представляющей угрозу жизни и здоровью людей, сохранности имущества, незамедлительно сообщать о случившемся Заказчику или непосредственному руководителю. Происшедшие несчастные случаи с Исполнителем или его представителями расследуются и учитываются Исполнителем, с привлечением в состав комиссии представителя Заказчика.\n" +
	    "2.2.9.  В случае отсутствия угрозы для жизни и здоровья Исполнителя незамедлительно принимать меры по устранению причин и условий, препятствующих нормальному выполнению работы.\n" +
	    "2.3.    ЗАКАЗЧИК ВПРАВЕ: в любой момент наблюдать, находиться на объекте оказания услуг и контролировать ход оказания услуг\n" +
	    "2.4.    ИСПОЛНИТЕЛЬ ВПРАВЕ: приостановить оказание услуг, в случае если представляется невозможным исполнить их без нарушения требований техники безопасности, при этом обязан незамедлительно оповестить об таких обстоятельствах Заказчика.")

        doc.add_paragraph("").add_run("3.ПОРЯДОК РАСЧЕТОВ ПО ДОГОВОРУ").bold = True
        # doc.add_page_break()
        doc.add_paragraph("3.1. Стоимость оказанных услуг рассчитывается на основании согласованных между сторонами тарифов (Приложение № 1, которое является неотъемлемой частью настоящего Договора).\n" +
	    "3.2. Расчеты между Сторонами производятся в рублях Российской Федерации любым способом, не запрещенным законодательством РФ.\n" +
	    "3.2.1.Оплата производиться путем перечисления на карту в течение 15 календарных дней с момента подписания Акта оказанных услуг/выполненных работ.")

        doc.add_paragraph("").add_run("4.СРОК ДЕЙСТВИЯ ДОГОВОРА").bold = True
        doc.add_paragraph("4.1.    Срок выполнения услуг (срок действия настоящего Договора) устанавливается.\n" +
        "Начало работ: " + start_date + "\n" +
        "Окончание работ: " + end_date + "\n" +
        "4.2. Окончание срока действия Договора не освобождает стороны от исполнения обязательств, возникающих в период действия Договора.\n" +
        "4.3. Перенос сроков начала и окончания Работ, дополнительно согласовывается Сторонами.\n" +
        "4.4. В случае исполнения обязательств обеими Сторонами настоящего договора возможно расторжение в одностороннем порядке со стороны Заказчика. Датой расторжения признается дата вручения/направления почтовым уведомлением Уведомления о прекращении договора.")

        doc.add_paragraph("").add_run("5. ОТВЕТСТВЕННОСТЬ СТОРОН").bold = True
        doc.add_paragraph("5.1. В случае, если Заказчик задерживает выплату в установленный срок, Исполнитель вправе потребовать дополнительно перечисления пени в размере 0,01 % в день от общей суммы задолженности Заказчика.\n" +
        "5.2. Исполнитель не несет ответственности за повреждение или гибель имущества Заказчика, если, несмотря на предупреждение Заказчика о невозможности обеспечить безопасное перемещение любого предмета, Заказчик, его сотрудник или лицо, находящееся по адресу отправления или назначения, настаивают на принятии Исполнителем данного предмета к обработке.\n" +
        "5.3. В случае нарушения Заказчиком условий Договора, которые дают Исполнителю право отказаться от исполнения заявки либо фактически приводят к невозможности исполнения, обязательства Исполнителя по выполнению заявки считаются прекращенными. При этом Заказчик возмещает Исполнителю причиненные в результате этого убытки.\n" +
        "5.4. Исполнитель несет ответственность за наличие предметов, описанных в Перечне перемещаемого имущества, подписанного представителем Исполнителя.\n" +
        "5.5. Исполнителя, незамедлительно составляется акт о причинении материального ущерба в присутствии представителей Заказчика и Исполнителя.\n" +
        "5.6. За утерянные, по словам Заказчика, предметы, не описанные в Перечне перемещаемого имущества, Исполнитель ответственности не несет.\n" +
        "5.7. В случае перехода в другую компанию (заключения договора) в рамках объекта на котором оказывается услуга, Исполнитель выплачивает компенсацию Заказчику в размере 20000 (Двадцати тысяч) рублей.")

        doc.add_paragraph("").add_run("6.РАЗРЕШЕНИЕ СПОРОВ").bold = True
        doc.add_paragraph("6.1. Договором установлен обязательный досудебный порядок разрешения споров между Сторонами по Договору:  до обращения в суд с иском Сторона, имеющая какое-либо требование к другой Стороне по Договору, обязана направить ей письменную претензию с обоснованием своего требования, к которой должны быть приложены оригиналы либо удостоверенные направившей претензию  Стороной копии документов, подтверждающих это требование. Срок ответа на претензию устанавливается 10 рабочих дней.\n" +
        "6.2. Все споры и разногласия, которые могут возникнуть в рамках Договора или в связи с ним, по возможности, разрешаются путем переговоров между Сторонами, при не достижении согласия в суде.")

        doc.add_paragraph("").add_run("7.ФОРС-МАЖОРНЫЕ ОБСТОЯТЕЛЬСТВА").bold = True
        doc.add_paragraph("7.1. Любая сторона, не выполнившая обязательства по Договору вследствие обстоятельств непреодолимой силы, обязана уведомить вторую сторону не позднее чем через 3(три) рабочих дня после того, как ей стало известно о наличии данных обстоятельств.\n" +
        "7.2. В случае, если форс-мажорные обстоятельства повлияли на выполнение условий по Договору, стороны в течение 10 (десяти) рабочих дней решают о возможности продолжения отношений по Договору.\n" +
        "7.3. К обстоятельствам непреодолимой силы следует относить: военные действия, чрезвычайное положение в регионе, стихийные бедствия, гражданские беспорядки.")

        doc.add_paragraph("").add_run("8. ПРОЧИЕ УСЛОВИЯ").bold = True
        doc.add_paragraph("8.1. Договор составлен и подписан в двух экземплярах, имеющих одинаковую юридическую силу. Для каждой стороны по одному экземпляру.\n" +
	    "8.2. Все изменения и дополнения к Договору действительны, если они оформлены в письменном виде, согласованы и подписаны обеими Сторонами.\n" +
	    "8.3. Исполнитель дает свое согласие на получение от Заказчика смс-сообщений на номер  рекламного и иного характера.")

        doc.add_paragraph("").add_run("9. ЮРИДИЧЕСКИЕ АДРЕСА, РЕКВИЗИТЫ И ПОДПИСИ СТОРОН").bold = True
        table = doc.add_table(rows=2, cols=2)
        table.style = 'Table Grid'
        col = table.rows[0].cells
        col[0].text = 'ПОДРЯДЧИК'
        col[1].text = 'ЗАКАЗЧИК'
        col = table.rows[1].cells
        col[0].text = '"____" _________________ 20___ г.'
        col[1].text = organization + '\nЮридический адрес: ' + legalAddress + '\n' +\
        'ИНН: ' + inn + '\n' +\
        'КПП: ' + kpp + '\n' +\
        'ОРГН: ' + ogrn + '\n\n' +\
        '_____________ / ' + CEO + '/\n\n' +\
        'М.П.\n\n"____" _________________ 20___ г.'

        doc.add_page_break()
        doc.add_paragraph("").add_run("Приложение №1").bold = True
        doc.add_paragraph("К Договору возмездного оказания услуг № " + number + "\n" +
        "От " + start_date)

        doc.add_paragraph("").add_run("Размер оплаты по настоящему Договору").bold = True
        table = doc.add_table(rows=2, cols=3)
        table.style = 'Table Grid'
        col = table.rows[0].cells
        col[0].text = '№'
        col[1].text = 'Тип услуги'
        col[2].text = 'Цена за услугу'
        col = table.rows[1].cells
        col[0].text = '1'
        col[1].text = ' '
        col[2].text = '500 руб'

        doc.add_paragraph("")

        doc.add_paragraph("").add_run("Расчеты производятся в рублях.").bold = True
        doc.add_paragraph("").add_run("Подписи сторон:").bold = True
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        col = table.rows[0].cells
        col[0].text = 'Исполнитель\n\n_____________ / ' + individual + '/\nМ.П\n'
        col[1].text = 'Заказчик\n\n_____________ / ' + CEO + '/\nМ.П\n'
        # col = table.rows[1].cells
        # col[0].text = 'М.П'
        # col[1].text = 'М.П'

        doc.add_page_break()
        doc.add_paragraph(individual +
        "\nПаспорт иностранного\n" +
        "гражданина серия: " + passport_series + " номер:\n" +
        passport_number)

        doc.add_paragraph("Исх № ____ от ________")

        doc.add_paragraph("Уведомление\nО прекращении договора возмездного оказания услуг.")

        doc.add_paragraph("В соответствии с пунктом 4.4 Договора возмездного оказания услуг № 48 /06-22 от 7 июня 2022 г. (далее Договор), заключенного между " + organization + " и гражданином Киргизии " + individual + ", настоящим уведомляем Вас о том, что " + organization + " в одностороннем порядке расторгает вышеуказанный Договор.\n" +
        organization + " подтверждает исполнение обязательств по оплате услуг в полном объеме. Услуги Исполнителем оказаны лично в объеме по условиям вышеуказанного  Договора.\n" +
        organization + " просит считать Договор возмездного оказания услуг №  " + number + " от " + start_date + "\nрасторгнутым с даты направления  настоящего уведомления.")
        doc.add_paragraph("")
        doc.add_paragraph("")
        doc.add_paragraph("Генерального директора __________________  /" + CEO + "/")


        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'GPC agreement.docx') == False:
            doc.save(path + '/' + 'GPC agreement.docx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'GPC agreement{i}.docx') == False:
                    path = path_file + '/' + f'GPC agreement{i}.docx'
                    doc.save(path)
                    break
                i += 1
        # print(path)

        # host = "83.220.171.210"
        # port = "1433"
        # user = "sa"
        # password = "Secretochka2442"
        #
        # # try:
        # connection = pymssql.connect(
        #     host=host,
        #     port=port,
        #     user=user,
        #     password=password,
        # )
        # #     print("Успех")
        # # except Exception as ex:
        # #     print("Ошибка")
        # dbCursor = connection.cursor()
        # print(dbCursor)

        # obj = Company()
        # print(obj.response.type)
    return redirect('gpc_contract')



def Removal_order(request):
    return render(request, 'document_generation_app/removal_order.html')


def Generate_removal_older(request):
    if request.method =='POST':
        company = CompanyAPI()
        organization = company["organizationalForm"] + ' "' + company["name"] + '"'
        phone = company["contactInfo"]["phone"]
        inn = company['inn']
        kpp = company['kpp']
        ogrn = company['ogrn']
        paymentAccount = company['bank']['paymentAccount']
        correspondentAccount = company['bank']['correspondentAccount']
        legalAddress = company['legalAddress']["name"]
        ActualAddreses = company['ActualAddreses'][0]["name"]
        CEO = 'Сталюковой Екатерина Александровны'

        individual = 'Гайназаров Кайратбек'
        passport_series = 'AC'
        passport_number = '4348554'

        number = request.POST.get('number')
        obj_start_date = datetime.strptime(request.POST.get('start_date'), '%Y-%m-%d')
        start_date = obj_start_date.strftime("%d-%m-%Y")


        path_file_doc = 'document_generation_app/document_templates/removal_order.docx'
        doc = DocxTemplate(path_file_doc)


        context = {
            'organization': organization,
            'number': number,
            'startDateQuotes': Date_conversion(start_date, 'quotes'),
            'startDateWordMonth': Date_conversion(start_date, 'word_month'),
            'startDateStandart': Date_conversion(start_date),
            'inn': inn,
            'kpp': kpp,
            'phone': phone,
            'legalAddress': legalAddress,
            'ActualAddreses': ActualAddreses,
            'paymentAccount': paymentAccount,
            'correspondentAccount': correspondentAccount,
            'CEO': CEO,
            'individual': individual
        }

        doc.render(context)

        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'removal_order.docx') == False:
            doc.save(path + '/' + 'removal_order.docx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'removal_order{i}.docx') == False:
                    path = path_file + '/' + f'removal_order{i}.docx'
                    doc.save(path)
                    break
                i += 1

    return redirect('removal_order')


def Notice_conclusion(request):
    return render(request, 'document_generation_app/notice_conclusion.html')


def Generate_Notice_conclusion(request):
    if request.method == 'POST':
        company = CompanyAPI()
        organization = company["organizationalForm"] + ' "' + company["name"] + '"'
        list_organization = list(organization.upper())
        phone = company["contactInfo"]["phone"].upper()
        list_phone = list(phone)
        inn = company['inn']
        kpp = company['kpp']
        list_inn_kpp = list(f'{inn}' + '/' + f'{kpp}')
        ogrn = company['ogrn']
        list_ogrn = list('ОГРН ' + ogrn)
        paymentAccount = company['bank']['paymentAccount'].upper()
        correspondentAccount = company['bank']['correspondentAccount'].upper()
        legalAddress = company['legalAddress']["name"].upper()
        list_legalAddress = list(legalAddress)
        ActualAddreses = company['ActualAddreses'][0]["name"].upper()
        CEO = 'Сталюкова Екатерина Александровна'

        individual = 'Гайназаров Кайратбек'.upper()
        # passport_series = 'AC'.upper()
        # passport_number = '4348554'.upper()

        name = request.POST.get('name').upper()
        list_name = list(name)
        job_title = request.POST.get('job_title').upper()
        list_job = list(job_title)
        base = request.POST.get('base')
        obj_start_date = datetime.strptime(request.POST.get('start_date'), '%Y-%m-%d')
        start_date = obj_start_date.strftime("%d-%m-%Y")
        address = request.POST.get('address').upper()
        list_address = list(address)
        person = request.POST.get('person')

        path_file_doc = 'document_generation_app/document_templates/notice_conclusion.xlsx'
        doc = load_workbook(path_file_doc)
        sheet = doc.active

        list_columns = ['A', 'C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']
        
        row = 11
        index = 0
        for symbol in list_name:
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


        row = 41
        index = 0
        for symbol in list_organization:
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


        row = 58
        index = 0
        for symbol in list_ogrn:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        row = 73
        index = 0
        for symbol in list_inn_kpp:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        row = 76
        index = 0
        for symbol in list_legalAddress:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                if list_columns[index] == 'BQ':
                    sheet[f'{cell}'] = symbol
                    row += 3
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break


        row = 85
        index = 0
        list_columns_phone = ['U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']
        for symbol in list_phone:
            for col in range(index, len(list_columns_phone)):
                cell = list_columns_phone[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        row = 157
        index = 0
        for symbol in list_job:
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


        if base == 'Трудовой договор':
            sheet['A167'] = 'X'
        elif base == 'Гражданско-правовой договор на выполнение работ (оказание услуг)':
            sheet['W167'] = 'X'


        start_date = Date_conversion(start_date)
        arr_date = start_date.split('.')
        # day
        sheet['AX172'] = arr_date[0][0]
        sheet['AZ172'] = arr_date[0][1]
        # month
        sheet['BC172'] = arr_date[1][0]
        sheet['BE172'] = arr_date[1][1]
        # year
        sheet['BH172'] = arr_date[2][0]
        sheet['BJ172'] = arr_date[2][1]
        sheet['BL172'] = arr_date[2][2]
        sheet['BN172'] = arr_date[2][3]


        row = 178
        index = 0
        for symbol in list_address:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                if list_columns[index] == 'BQ':
                    sheet[f'{cell}'] = symbol
                    if row == 180:
                        row += 3
                    else:
                        row += 2
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break

        sheet['AK191'] = CEO

        if person == 'person_proxy':
            try:
                full_name = request.POST.get('full_name')
                passportSeries = request.POST.get('series')
                passportNumber = request.POST.get('number')
                obj_date_issue = datetime.strptime(request.POST.get('date_issue'), '%Y-%m-%d')
                date_issue = obj_date_issue.strftime("%d-%m-%Y")
                issued_by = request.POST.get('issued_by')

                sheet["AE201"] = full_name
                passportSeries = passportSeries[0] + passportSeries[1] + ' ' + passportSeries[2] + passportSeries[3]
                sheet["G203"] = passportSeries
                sheet["X203"] = passportNumber
                sheet["AR203"] = Date_conversion(date_issue)
                sheet["J205"] = issued_by
            except:
                pass

        else:
            sheet["AE201"] = CEO


        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'notice_conclusion.xlsx') == False:
            doc.save(path + '/' + 'notice_conclusion.xlsx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'notice_conclusiont{i}.xlsx') == False:
                    path = path_file + '/' + f'notice_conclusiont{i}.xlsx'
                    doc.save(path)
                    break
                i += 1

    return redirect('notice_conclusion')


def Termination_noticen(request):
    return render(request, 'document_generation_app/termination_notice.html')


def Generate_Termination_notice(request):
    if request.method == 'POST':
        company = CompanyAPI()
        organization = company["organizationalForm"] + ' "' + company["name"] + '"'
        list_organization = list(organization.upper())
        phone = company["contactInfo"]["phone"].upper()
        list_phone = list(phone)
        inn = company['inn']
        kpp = company['kpp']
        list_inn_kpp = list(f'{inn}' + '/' + f'{kpp}')
        ogrn = company['ogrn']
        list_ogrn = list('ОГРН ' + ogrn)
        paymentAccount = company['bank']['paymentAccount'].upper()
        correspondentAccount = company['bank']['correspondentAccount'].upper()
        legalAddress = company['legalAddress']["name"].upper()
        list_legalAddress = list(legalAddress)
        ActualAddreses = company['ActualAddreses'][0]["name"].upper()
        CEO = 'Сталюкова Екатерина Александровна'

        individual = 'Гайназаров Кайратбек'.upper()
        # passport_series = 'AC'.upper()
        # passport_number = '4348554'.upper()

        name = request.POST.get('name').upper()
        list_name = list(name)
        job_title = request.POST.get('job_title').upper()
        list_job = list(job_title)
        base = request.POST.get('base')
        obj_end_date = datetime.strptime(request.POST.get('end_date'), '%Y-%m-%d')
        end_date = obj_end_date.strftime("%d-%m-%Y")
        initiator = request.POST.get('initiator')
        person = request.POST.get('person')

        path_file_doc = 'document_generation_app/document_templates/termination_notice.xlsx'
        doc = load_workbook(path_file_doc)
        sheet = doc.active

        list_columns = ['A', 'C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
                        'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
                        'BQ']

        row = 11
        index = 0
        for symbol in list_name:
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

        row = 41
        index = 0
        for symbol in list_organization:
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

        row = 58
        index = 0
        for symbol in list_ogrn:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break

        row = 72
        index = 0
        for symbol in list_inn_kpp:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break

        row = 75
        index = 0
        for symbol in list_legalAddress:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                if list_columns[index] == 'BQ':
                    sheet[f'{cell}'] = symbol
                    row += 3
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break

        row = 84
        index = 0
        list_columns_phone = ['U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW',
                              'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']

        for symbol in list_phone:
            for col in range(index, len(list_columns_phone)):
                cell = list_columns_phone[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        row = 151
        index = 0
        for symbol in list_job:
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

        if base == 'Трудовой договор':
            sheet['A161'] = 'X'
        elif base == 'Гражданско-правовой договор на выполнение работ (оказание услуг)':
            sheet['W161'] = 'X'

        start_date = Date_conversion(end_date)
        arr_date = start_date.split('.')
        # day
        sheet['AX167'] = arr_date[0][0]
        sheet['AZ167'] = arr_date[0][1]
        # month
        sheet['BC167'] = arr_date[1][0]
        sheet['BE167'] = arr_date[1][1]
        # year
        sheet['BH167'] = arr_date[2][0]
        sheet['BJ167'] = arr_date[2][1]
        sheet['BL167'] = arr_date[2][2]
        sheet['BN167'] = arr_date[2][3]


        if initiator == 'yes':
            sheet['A174'] = 'X'
        else:
            sheet['W174'] = 'X'


        sheet['AK183'] = CEO

        if person == 'person_proxy':
            try:
                full_name = request.POST.get('full_name')
                passportSeries = request.POST.get('series')
                passportNumber = request.POST.get('number')
                obj_date_issue = datetime.strptime(request.POST.get('date_issue'), '%Y-%m-%d')
                date_issue = obj_date_issue.strftime("%d-%m-%Y")
                issued_by = request.POST.get('issued_by')

                sheet["AE193"] = full_name
                passportSeries = passportSeries[0] + passportSeries[1] + ' ' + passportSeries[2] + passportSeries[3]
                sheet["G195"] = passportSeries
                sheet["X195"] = passportNumber
                sheet["AR195"] = Date_conversion(date_issue)
                sheet["J197"] = issued_by
            except:
                pass

        else:
            sheet["AE193"] = CEO

        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'termination_notice.xlsx') == False:
            doc.save(path + '/' + 'termination_notice.xlsx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'termination_notice{i}.xlsx') == False:
                    path = path_file + '/' + f'termination_notice{i}.xlsx'
                    doc.save(path)
                    break
                i += 1

    return redirect('termination_notice')


def Right_not_to_withhold_pit(request):
    return render(request, 'document_generation_app/right_not_to_withhold_pit.html')


def Generate_Right_not_to_withhold_pit(request):
    if request.method == 'POST':
        company = CompanyAPI()
        organization = company["organizationalForm"] + ' "' + company["name"] + '"'
        list_organization = list(organization.upper())
        phone = company["contactInfo"]["phone"].upper()
        list_phone = list(phone)
        inn = company['inn']
        list_inn = list(inn)
        kpp = company['kpp']
        ogrn = company['ogrn']
        list_ogrn = list(ogrn)
        paymentAccount = company['bank']['paymentAccount'].upper()
        correspondentAccount = company['bank']['correspondentAccount'].upper()
        legalAddress = company['legalAddress']["name"].upper()
        list_legalAddress = list(legalAddress)
        ActualAddreses = company['ActualAddreses'][0]["name"].upper()
        CEO = 'Сталюкова Екатерина Александровна'

        individual = 'Гайназаров Кайратбек'.upper()
        # passport_series = 'AC'.upper()
        # passport_number = '4348554'.upper()

        code = request.POST.get('code')
        list_code = list(code)


        path_file_doc = 'document_generation_app/document_templates/right_not_to_withhold_pit.xlsx'
        doc = load_workbook(path_file_doc)
        sheet = doc.active

        list_columns = ['X', 'AA', 'AD', 'AG', 'AJ',
                        'AM', 'AP', 'AS', 'AV', 'AY', 'BB', 'BH']

        row = 2
        index = 0
        for symbol in list_inn:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        list_columns = ['X', 'AA', 'AD', 'AG', 'AJ',
                        'AM', 'AP', 'AS', 'AV']
        row = 4
        index = 0
        for symbol in list_ogrn:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        # Input code
        sheet["CF8"] = list_code[0]
        sheet["CJ8"] = list_code[1]
        sheet["CO8"] = list_code[2]
        sheet["CT8"] = list_code[3]


        list_columns = ['B', 'C', 'E', 'F', 'G', 'I', 'L', 'N', 'P', 'S', 'U', 'W', 'Z', 'AC', 'AF', 'AI','AL',
                        'AO', 'AR', 'AU', 'AX', 'AY', 'BA', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BE', 'BJ', 'BK', 'BL',
                        'BM', 'BO', 'BP', 'BR', 'BS', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CE']
        row = 11
        index = 0
        for symbol in list_organization:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                if list_columns[index] == 'CE':
                    sheet[f'{cell}'] = symbol
                    row += 2
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break


        now_date = datetime.now(pytz.timezone('UTC'))
        now_date = Date_conversion(now_date.strftime('%d-%m-%Y'))
        list_now_date = list(now_date)

        list_columns = ['U', 'W', 'Z', 'AC', 'AF',
                        'AI', 'AL', 'AO', 'AR', 'AU']
        row = 46
        index = 0
        for symbol in list_now_date:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'right_not_to_withhold_pit.xlsx') == False:
            doc.save(path + '/' + 'right_not_to_withhold_pit.xlsx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx') == False:
                    path = path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx'
                    doc.save(path)
                    break
                i += 1

    return redirect('right_not_to_withhold_pit')


def About_arrival(request):
    return render(request, 'document_generation_app/about_arrival.html')


def Generate_About_arrival(request):
    if request.method == 'POST':
        generation_about_arrival(request)

    return redirect('about_arrival')


# {
#     "id": "1fa85f64-1727-4862-b3fc-2c963f66afa4",
#     "inn": "7804525530",
#     "kpp": "780401001",
#     "name": "Капитал Кадры",
#     "organizationalForm": "ООО",
#     "ogrn": "1117746358608",
#     "okved": "123456",
#     "legalAddress": {
#         "name": "195299, Г.Санкт-Петербург, пр-кт Гражданский, д. 119 ЛИТЕР А, офис 8"
#     },
#     "contactInfo": {
#         "id": 0,
#         "name": "Олег",
#         "phone": "+78989321223",
#         "email": "user@example.com"
#     },
#     "bank": {
#         "id": 0,
#         "bankId": "123456789",
#         "correspondentAccount": "12345123451234512345",
#         "paymentAccount": "12345123451234512345"
#     },
#     "ActualAddreses": [
#         {"name": "195299, Г.Санкт-Петербург, пр-кт Гражданский, д. 119 ЛИТЕР А, офис 8"}
#     ]
# }