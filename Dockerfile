FROM python:3
ADD email_rulz.py /
ADD ./rulz/rulz.db /rulz/rulz.db
RUN pip3 install imap-tools
CMD [ "python3","./email_rulz.py" ]
