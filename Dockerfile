FROM tiangolo/uvicorn-gunicorn-fastapi:python3.8

RUN apt update
RUN apt install -y locales-all

COPY ./requirements.txt /
RUN pip install -r /requirements.txt

COPY ./app /app

