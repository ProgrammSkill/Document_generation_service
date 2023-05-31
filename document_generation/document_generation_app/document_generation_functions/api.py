import requests
import datefinder
import re
from dadata import Dadata

# Migrascope External API

id_company = "063fa310-f766-44c8-a973-4fa576d7edc7"
id_employee = "d6aefc6a-51eb-4270-ae4b-d95ad4a5fd0a"

# weq = requests.get("https://proverkacheka.nalog.ru:9999/v1/inns/*/kkts/*/fss/996044041%D0%98913143/tickets/121838?fiscalSign=0696672240&sendToEmail=no")
# print(weq)

headers = {"Authorization": "Bearer eyJhbGciOiJSUzI1NiIsImtpZCI6IkQ5RjdBRjlBRDIyRDU0QTlDMjZCMEUyRTA0QzJBNUY3IiwidHlwIjoiYXQrand0In0.eyJpc3MiOiJodHRwOi8vc2VjcmV0b2Noa2EucnU6NDg5MTEiLCJuYmYiOjE2ODUxMTA3NDMsImlhdCI6MTY4NTExMDc0MywiZXhwIjoxNjg1MTE0MzQzLCJhdWQiOlsiY29tcGFueV9zZXJ2aWNlIiwiZG9jX3NlcnZpY2UiXSwic2NvcGUiOlsiY29tcGFueV9zZXJ2aWNlLnJlYWQiLCJjb21wYW55X3NlcnZpY2Uud3JpdGUiLCJkb2Nfc2VydmljZS5yZWFkIiwiZG9jX3NlcnZpY2Uud3JpdGUiLCJlbWFpbCIsIm9wZW5pZCIsInByb2ZpbGUiLCJvZmZsaW5lX2FjY2VzcyJdLCJhbXIiOlsicHdkIl0sImNsaWVudF9pZCI6IndlYi5jbGllbnQiLCJzdWIiOiIwOWM3YTNiNS0zZmFhLTQ5YmItOGIzNS1kMzFjNWQxNzBjNTMiLCJhdXRoX3RpbWUiOjE2ODUxMTA3NDMsImlkcCI6ImxvY2FsIiwiQ29tcGFueUNvdW50IjoiMTAwMDAiLCJVc2VyQ291bnQiOiIxMDAwMCIsIkVtcGxveWVlQ291bnQiOiIxMDAwMCIsIk1lbW9yeVNpemUiOiIxMDAwMCIsIkRvY0dlbkxldmVsIjoiZnVsbCIsIkRvY0dlbkNvdW50IjoiMTAwMDAiLCJvd25lciI6IjA2M2ZhMzEwLWY3NjYtNDRjOC1hOTczLTRmYTU3NmQ3ZWRjNyIsImh0dHA6Ly9zY2hlbWFzLnhtbHNvYXAub3JnL3dzLzIwMDUvMDUvaWRlbnRpdHkvY2xhaW1zL25hbWUiOiIwOWM3YTNiNS0zZmFhLTQ5YmItOGIzNS1kMzFjNWQxNzBjNTMiLCJqdGkiOiJDNDg5MzgzMjQyNTIxODlDQzEyMEQyRjg3NTZFNTBEMCJ9.l2iqNt2P2iwpc99vW4xDRZt2EVA7dmJYtpzQe9zthdA7GKLFhkSdFCYobwrJcM5izoGae6-6gf_J_0ag2lpdUfKBSOXJXVuuC7mxTlmkl0V0Ls6hIfp7BSsCj1r6k5C7NGx4CqiWAYvXKChFWVNYZToS27KwdHNW8uCmc4AboW4AOP6j42BBqisPI5bWd7-Df9GuKW98Hqlzcz6pZF_iU_F_0sVscpAHSDifYP-QjC0eAn1npjwkRUnFr_6pLnjsiclqqZ22AY0CWrx6nfdUSTU90Mgb7Xi6BNQ0JfuX7EYxj8nZWyMKiUyur-OrpEgEAfujXM5sq-k1fwTa2R7o2w"}

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