FROM python:3.12
COPY mapper.py /
COPY requirements.txt /tmp
WORKDIR /
RUN python3 -m pip install -r /tmp/requirements.txt

ENTRYPOINT [ "/mapper.py" ]