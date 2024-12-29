FROM ruby:3.1.3

WORKDIR /app

RUN apt-get update && apt-get install -y cron && apt-get install -y tzdata
ENV TZ=Asia/Tokyo
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime

ENV GOOGLE_APPLICATION_CREDENTIALS '/app/application_default_credentials.json'
ENV PATH="/usr/local/bundle/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:${PATH}"

COPY Gemfile Gemfile.lock ./
RUN bundle install
