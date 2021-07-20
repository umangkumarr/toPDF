# toPDF
Convert doc/docx file(s) to pdf.

# Dependencies:
    1. LibreOffice
    2. Python3

# How to Use:
Add relative/absolute path of the folder containing doc files to line 5 in the code:

    path=''
Run this command in the new terminal to start LibreOffice as a daemon service:

    soffice --nologo --headless --nofirststartwizard --accept='socket,host=127.0.0.1,port=2220,tcpNoDelay=1;urp;StarOffice.Service'

