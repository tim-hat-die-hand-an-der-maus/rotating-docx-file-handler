FROM python:3.12-slim

WORKDIR /usr/src/app

ADD requirements.txt requirements.txt

RUN pip install --no-cache-dir -r requirements.txt
