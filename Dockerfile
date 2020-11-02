FROM python:3.7-buster
MAINTAINER rashbirkohli@gmail.com

#ENV TZ=Europe/Kiev
#RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

#RUN apt-get update && \
#    apt upgrade -y
#    apt install software-properties-common -y && \
#    add-apt-repository ppa:deadsnakes/ppa -y && \
#    apt update && \
#    apt install python3.8 -y && \
#    apt install python3-pip -y

RUN apt-get update && apt-get install -y r-base

RUN mkdir /var/webapp

COPY ./webapp/requirements.txt /var/webapp

#WORKDIR /var

#RUN apt-get install wget -y && \
#    apt-get install make -y

#RUN wget http://download.redis.io/releases/redis-6.0.8.tar.gz && \
#    tar xzf redis-6.0.8.tar.gz

#WORKDIR /var/redis-6.0.8

#RUN make

WORKDIR /var/webapp

#RUN apt-get install r-base -y

RUN Rscript -e "install.packages('lubridate')"

RUN pip install --no-cache-dir -r requirements.txt

#pip install package --no-cache-dir

#RUN python -m spacy download en_core_web_sm

#RUN pip install https://blackstone-model.s3-eu-west-1.amazonaws.com/en_blackstone_proto-0.0.1.tar.gz

#RUN src/redis-server

#RUN celery -A webappproject worker -l info --loglevel=debug

#CMD python3 manage.py runserver 0.0.0.0:8000



