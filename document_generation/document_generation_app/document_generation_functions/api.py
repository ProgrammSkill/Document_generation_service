import requests
import datefinder
import re

id_company = "b3d4e844-db00-40ed-865c-90179a363338"
id_employee = "38373a51-7260-493e-a7d2-5b3dad990d2f"

headers = {"Authorization": "Bearer eyJhbGciOiJSUzI1NiIsImtpZCI6IkEwQ0JEMzEzREEyOUE2Q0E0N0I2ODE3MTNEQ0I5RTIxIiwidHlwIjoiYXQrand0In0.eyJpc3MiOiJodHRwOi8vc2VjcmV0b2Noa2EucnU6NDg5MTEiLCJuYmYiOjE2ODMyOTYxNDMsImlhdCI6MTY4MzI5NjE0MywiZXhwIjoxNjgzMjk5NzQzLCJhdWQiOlsiY29tcGFueV9zZXJ2aWNlIiwiZG9jX3NlcnZpY2UiXSwic2NvcGUiOlsiY29tcGFueV9zZXJ2aWNlLnJlYWQiLCJjb21wYW55X3NlcnZpY2Uud3JpdGUiLCJkb2Nfc2VydmljZS5yZWFkIiwiZG9jX3NlcnZpY2Uud3JpdGUiLCJlbWFpbCIsIm9wZW5pZCIsInByb2ZpbGUiLCJvZmZsaW5lX2FjY2VzcyJdLCJhbXIiOlsicHdkIl0sImNsaWVudF9pZCI6IndlYi5jbGllbnQiLCJzdWIiOiI0OTViYzQyOS04NzBhLTQ1MTItODc1Mi1kMzhlYjUxZWQ0YTIiLCJhdXRoX3RpbWUiOjE2ODMyOTYxNDMsImlkcCI6ImxvY2FsIiwiQ29tcGFueUNvdW50IjoiMTAwMCIsIlVzZXJDb3VudCI6IjEwMDAiLCJFbXBsb3llZUNvdW50IjoiMTAwMCIsIk1lbW9yeVNpemUiOiIxMDAwIiwiRG9jR2VuTGV2ZWwiOiJmdWxsIiwiRG9jR2VuQ291bnQiOiIxMDAwMCIsImh0dHA6Ly9zY2hlbWFzLnhtbHNvYXAub3JnL3dzLzIwMDUvMDUvaWRlbnRpdHkvY2xhaW1zL25hbWUiOiI0OTViYzQyOS04NzBhLTQ1MTItODc1Mi1kMzhlYjUxZWQ0YTIiLCJqdGkiOiI3N0YyNjBGRTE0OTY2NzkyODhFNjhBRTZGRjBDRUMzNSJ9.NqYon625Dre5X9uKCz_3DFFljEs_lwiQDEPvPdPxuRkzJYnh32NPgsuYH6ZQFXzT2jMK_9Oluch71RVeDoUBKFvFG7pax_leR2s4P-EJ9hdm2btzBLGZmj560Z1rrgwp49ZNAPCJdM7N2POyK-nE9h10PoyiPNJR7JV8STkeLk2U9MZFaOUuxUlYFuV9xA3Wg6pwcv3MibyhLHws3787Fa1Erw9E38mEJifeIomnuPCvsW10vLv7uxANMn1GaX2Wzz6cJa_vuikmg54g_L6GLil_DNiP3MB8RY_csxWLNS2zE3EQmXvTusyeDFXbPIQ5PPIu4_NLs_h5U_60l4KLuA"}

def CompanyAPI():
    # id_company = "b3d4e844-db00-40ed-865c-90179a363338"
    # response = requests.get("http://secretochka.ru:48903/api/v1/manager/companyes/" + id_company,
    #                         headers=headers).json()
    # actual_address = requests.get("http://secretochka.ru:48903/api/v1/companyes/" + id_company +
    #                               "/info/actualaddresses", headers=headers).json()
    # response["ActualAddreses"] = actual_address
    response = {
        "id": "876cf8f7-bcdf-46b3-9017-99977f601ebc",
        "inn": "7804525530",
        "kpp": "780401001",
        "name": "Капитал Кадры",
        "organizationalForm": "ООО",
        "ogrn": "1117746358608",
        "okved": "123123123",
        "legalAddress": {
            "city": "Санкт-Петербург",
            "street": "пр-кт Гражданский",
            "house": "119  ЛИТЕР А, офис 8"
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
        "ActualAddreses": [
            {
                "city": "Санкт-Петербург",
                "street": "пр-кт Гражданский",
                "house": "119  ЛИТЕР А, офис 8"
            }
        ]
    }

    return response

print(CompanyAPI())

def IndividualAPI():
    # response = requests.get("http://secretochka.ru:48903/v1/forservice/employees/" + id_employee, headers=headers).json()
    # passport = requests.get("http://secretochka.ru:48907/api/v1/" + id_company + "/employees/" + id_employee +
    #                         "/documents/info/Passport?page=0&pageSize=10",  headers=headers).json()[0]
    # response['passport'] = passport




    # response = {
    #     "id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
    #     "surname": 'Гайназаров',
    #     "name": 'Кайратбек',
    #     "patronymic": None,
    #     "inn": "4857522812",
    #     "kpp": "413401002",
    #     'temporary_registration': 'РОССИЯ, 188310, Ленинградская обл, Гатчинский р-н, Гатчина г, Авиатриссы Зверевой ул, Дом 20, Корпус 1, Квартира 5',
    #     "passport": {
    #         "serias": 'АС',
    #         "number": '8651642',
    #         "date_of_issue": '2012-10-25',
    #         "registration_address": "г. Алматы, р-н. Бостандский"
    #     },
    #     "patent": {
    #         "serias": 'AC',
    #         "number": '365451',
    #         "territory_of_action": 'Санкт-Петербург',
    #         "kind_of_activity": 'string',
    #         "date_of_issue": "2023-06-07",
    #         "expiration_date": '2023-06-08',
    #     },
    #     "bank": {
    #         "id": 0,
    #         "bankId": "12341234",
    #         "correspondentAccount": "12345123451234512345",
    #         "paymentAccount": "12345123451234512345"
    #     },
    # }

    response = {
          "id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
          "fio": {
            "firstName": "Кайратбек",
            "secondName": "Гайназаров",
            "patronymic": None
          },
          "positionId": 1,
          "position": {
            "name": "Слесарь"
          },
        "gender": 0,
        "citizenship": "Киргизия",
        "birthday": "1992-01-01T12:59:35.649Z",
        "regAddress": {
            "city": "Ташкент",
            "street": "ул. Рихсилий	",
            "house": "64"
        },
        "passport": {
            "serias": 'AC231',
            "number": '865161242',
            "dateIssue": '2008-05-04T21:09:09.383+00:00',
            "registration_address": "г. Алматы, р-н. Бостандский"
        },
        "patent": {
            "serias": 'AC',
            "number": '8651642',
            "territory_of_action": 'Санкт-Петербург',
            "kind_of_activity": 'string',
            "date_of_issue": "2023-06-07",
            "expiration_date": '2023-06-08',
        },
        'temporary_registration': 'РОССИЯ, 188310, Ленинградская обл, Гатчинский р-н, Гатчина г, Авиатриссы Зверевой ул, Дом 20, Корпус 1, Квартира 5',

    }

    return response
print(IndividualAPI())

# {
#     "id": "1fa25f64-5717-4562-b3fc-2c963f66afa6",
#     "inn": "7804525530",
#     "kpp": "780401001",
#     "name": "Капитал Кадры",
#     "organizationalForm": "ООО",
#     "ogrn": "1117746358608",
#     "okved": "123456",
#     "legalAddress": {
#         "city": "Санкт-Петербург",
#         "street": "пр-кт Гражданский",
#         "house": "119  ЛИТЕР А, офис 8"
#
#     },
#     "ActualAddreses": [
#         {
#             "city": "Санкт-Петербург",
#             "street": "пр-кт Гражданский",
#             "house": "119  ЛИТЕР А, офис 8"
#         }
#     ],
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
#     }
# }


# {
#   "inn": "7804525530",
#   "kpp": "780401001",
#   "name": "Капитал Кадры",
#   "organizationalForm": "ООО",
#   "ogrn": "1117746358608",
#   "okved": "123456",
#   "legalAddress": {
#     "city": "Санкт-Петербург",
#     "street": "пр-кт Гражданский",
#     "house": "119  ЛИТЕР А, офис 8"
#   },
#   "actualAddreses": [
#     {
#       "city": "Санкт-Петербург",
#       "street": "пр-кт Гражданский",
#       "house": "119  ЛИТЕР А, офис 8"
#     }
#   ],
#   "bank": {
#     "bankId": "123456789",
#     "correspondentAccount": "12345123451234512345",
#     "paymentAccount": "12345123451234512345"
#   }
# }