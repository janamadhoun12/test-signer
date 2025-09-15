FROM apify/actor-python:3.11
RUN apt-get update && apt-get install -y libreoffice
COPY . ./
RUN pip install --no-cache-dir -r requirements.txt