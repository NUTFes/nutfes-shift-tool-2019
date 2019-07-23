FROM python:3.6
ENV PYTHONUNBUFFERED 1
RUN apt-get update
RUN apt-get install -y vim
RUN mkdir -p /var/www/nutfes-shift-tool
ENV APP_PATH /var/www/nutfes-shift-tool
WORKDIR $APP_PATH
ADD requirements.txt $APP_PATH
RUN pip install mojimoji
RUN pip install -r requirements.txt
