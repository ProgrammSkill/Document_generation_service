import requests
import datefinder
import re
from dadata import Dadata

# Migrascope External API

id_company = "063fa310-f766-44c8-a973-4fa576d7edc7"
id_employee = "d6aefc6a-51eb-4270-ae4b-d95ad4a5fd0a"


headers = {"Authorization": "Bearer eyJhbGciOiJSUzI1NiIsImtpZCI6IkJBNzdFNjBCODY0OTlBQ0UwMjI5MDBFNUYwMzE0Q0JCIiwidHlwIjoiYXQrand0In0.eyJpc3MiOiJodHRwOi8vc2VjcmV0b2Noa2EucnU6NDg5MTEiLCJuYmYiOjE2ODU3MTcxMzksImlhdCI6MTY4NTcxNzEzOSwiZXhwIjoxNjg1NzIwNzM5LCJhdWQiOlsiY29tcGFueV9zZXJ2aWNlIiwiZG9jX3NlcnZpY2UiXSwic2NvcGUiOlsiY29tcGFueV9zZXJ2aWNlLnJlYWQiLCJjb21wYW55X3NlcnZpY2Uud3JpdGUiLCJkb2Nfc2VydmljZS5yZWFkIiwiZG9jX3NlcnZpY2Uud3JpdGUiLCJlbWFpbCIsIm9wZW5pZCIsInByb2ZpbGUiLCJvZmZsaW5lX2FjY2VzcyJdLCJhbXIiOlsicHdkIl0sImNsaWVudF9pZCI6IndlYi5jbGllbnQiLCJzdWIiOiIwOWM3YTNiNS0zZmFhLTQ5YmItOGIzNS1kMzFjNWQxNzBjNTMiLCJhdXRoX3RpbWUiOjE2ODU3MTcxMzksImlkcCI6ImxvY2FsIiwiQ29tcGFueUNvdW50IjoiMTAwMDAiLCJVc2VyQ291bnQiOiIxMDAwMCIsIkVtcGxveWVlQ291bnQiOiIxMDAwMCIsIk1lbW9yeVNpemUiOiIxMDAwMCIsIkRvY0dlbkxldmVsIjoiZnVsbCIsIkRvY0dlbkNvdW50IjoiMTAwMDAiLCJvd25lciI6IjA2M2ZhMzEwLWY3NjYtNDRjOC1hOTczLTRmYTU3NmQ3ZWRjNyIsImh0dHA6Ly9zY2hlbWFzLnhtbHNvYXAub3JnL3dzLzIwMDUvMDUvaWRlbnRpdHkvY2xhaW1zL25hbWUiOiIwOWM3YTNiNS0zZmFhLTQ5YmItOGIzNS1kMzFjNWQxNzBjNTMiLCJqdGkiOiJFQjYzRTM0ODY5NkRGODQ1MjBFRjQ3NzM3ODZCQzJCNiJ9.SbCTCvcm1MYjkUvY5ip0bQAHEgO0kfOg0w7oxah_1hTJMC9FK1SYiV3GCJFA88Mqt7slmmyMxX2-3sYumyhnwSIxaKFomabqTB1ElIWEUL6JH1XrPN3l0xVCr_HEm9BZ5AvrEn6mMHWxzpIB_TjnZOJJsPSRSwJL4mVIrKdIb2DbNIBntaNkaY4V0OsBoGWFGd1txSD9BWmuLeRagDniD9Vb3qpB-kQ8zux-loNVlQUJ7lRjq0kWr6VHFvCTfErSuIpombDwFmiuJfJy602ar0gRR8X_mC2q8UqzMX95AVb2IH6_RlOJUZCldzl0BnRHjOzMqbFqGLvQqo7BP7az8A"}

def CompanyAPI():
    response = requests.get("http://secretochka.ru:48903/api/v1/manager/companyes/" + id_company,
                            headers=headers).json()
    actual_address = requests.get("http://secretochka.ru:48903/api/v1/companyes/" + id_company +
                                  "/info/actualaddresses", headers=headers).json()
    response["ActualAddreses"] = actual_address
    director = requests.get("http://secretochka.ru:48903/v1/forservice/companys/" + id_company + "/employees?position=director",
                            headers=headers).json()[0]
    response['director'] = director
    bankId = response['bank']['bankId']
    token = "de4144a486735962685109c249279deb926a1331"
    dadata = Dadata(token)
    bankInfo = dadata.find_by_id("bank", bankId)
    response['bank']['nameBank'] = bankInfo[0]['value']
    response['bank']['city'] = bankInfo[0]['data']['payment_city']
    return response

print(CompanyAPI())

def IndividualAPI():
    response = requests.get("http://secretochka.ru:48903/v1/forservice/employees/" + id_employee, headers=headers).json()
    try: #If an error occurs when using a patent and passport, then this means that the employee does not have documents
        passport = requests.get("http://secretochka.ru:48907/api/v1/" + id_company + "/employees/" + id_employee +
                            "/documents/info/Passport?page=0&pageSize=10", headers=headers).json()[0]
        response['passport'] = passport
    except:
        response['passport'] = None

    try:
        patent = requests.get("http://secretochka.ru:48907/api/v1/" + id_company + "/employees/" + id_employee +
                                "/documents/work/Patent?page=0&pageSize=10", headers=headers).json()[0]
        response['patent'] = patent
    except:
        response['patent'] = None

    return response

print(IndividualAPI())