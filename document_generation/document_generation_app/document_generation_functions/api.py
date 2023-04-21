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
    # id_company = "b5deb54b-01bd-4bc3-87d0-21359b046e2a"
    # response = requests.get("http://secretochka.ru:48910/company-service/api/v1/manager/companyes/" + id_company).json()
    # actual_address = requests.get("http://secretochka.ru:48910/company-service/api/v1/companyes/" + id_company + "/info/actualaddresses").json()
    # response["ActualAddreses"] = actual_address
    response = {
        "id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
        "surname": 'Гайназаров',
        "name": 'Кайратбек',
        "patronymic": None,
        "inn": "4857522812",
        "kpp": "413401002",
        'temporary_registration': 'РОССИЯ, 188310, Ленинградская обл, Гатчинский р-н, Гатчина г, Авиатриссы Зверевой ул, Дом 20, Корпус 1, Квартира 5',
        "passport": {
            "series": 'АС',
            "number": '8651642',
            "date_of_issue": '2012-10-25',
            "registration_address": "г. Алматы, р-н. Бостандский"
        },
        "patent": {
            "series": 'AC',
            "number": '365451',
            "territory_of_action": 'Санкт-Петербург',
            "kind_of_activity": 'string',
            "date_of_issue": "2023-06-07",
            "expiration_date": '2023-06-08',
        },
        "bank": {
            "id": 0,
            "bankId": "12341234",
            "correspondentAccount": "12345123451234512345",
            "paymentAccount": "12345123451234512345"
        },
    }

    return response