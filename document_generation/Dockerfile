FROM python:3.10.11

WORKDIR /usr/src/Document_generation_service

COPY ./requirements.txt /usr/src/requirements.txt

RUN pip install -r /usr/src/requirements.txt

COPY . /usr/src/Document_generation_service

EXPOSE 8000
CMD ["python", "manage.py", "runserver", "0.0.0.0:80000"]