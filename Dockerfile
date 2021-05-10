FROM python:3
ADD email_rulz.py /
ADD ./rulz/rulz.db /rulz/rulz.db
RUN pip3 install imap-tools
CMD [ "python3","./email_rulz.py" ]
ENV IMAP_SERVER=mail.server.com
ENV IMAP_LOGIN=email@address.com
ENV IMAP_PWD=email_password
ENV EMAIL_ADDRESS=my_email@address.com