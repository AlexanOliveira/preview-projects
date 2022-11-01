from ast import Pass
from imap_tools import MailBox, A
from pathlib import Path
from ftplib import FTP
from datetime import datetime as dt, timedelta
import time
import os
import unidecode
import re as regex
import html


utc = dt.today() - timedelta(hours=3)

if utc.hour >= 0 and utc.hour < 8:
   exit()


with MailBox('imap.gmail.com').login('###@gmail.com', '####') as mailbox:
   i = 0

   txt = Path(__file__).with_name('lastEmail.txt')

   d= utc.date()
   all = ''

   try:
      all = list(mailbox.fetch(A(date=d), mark_seen=False))
   except Exception:
      time.sleep(30.0)
      all = list(mailbox.fetch(A(date=d), mark_seen=False))

   all.sort(key=lambda x: x.uid, reverse=True)

   ftp = FTP('host.com.br', 'user', 'passwd')

   for msg in all:
      i += 1

      date = str(msg.date)

      msg_utc = int(date[20:22]) - 3
      msg_hour = int(date[11:13])

      if msg_utc > 0:
         msg_hour += abs(msg_utc)
      elif  msg_utc < 0:
         msg_hour -= abs(msg_utc)

      msg_hour = str(msg_hour).zfill(2)
      date = date[0:11] + msg_hour + date[13:19]

      this_email = msg.from_ + ' ' + str(date)[0:19].replace(':', '.')

      if i == 1:
         last_email = this_email

      if this_email == txt.open('r').read():
         log = Path(__file__).with_name('cron.log')
         with log.open("rb") as file:
            file.seek(-64, 2)

            last_input = file.readline().decode('latin1')
            check = 'AutenLab BOT: VocÃª nÃ£o recebeu nenhum novo E-mail'

            if not check in last_input:
               print('\n', utc.strftime("%d/%m/%Y %H:%M:%S"),' ♦ AutenLab BOT: Você não recebeu nenhum novo E-mail ♦ \n' )
         break


      body = 'De: ' + msg.from_ \
           + '\nAssunto: ' + msg.subject \
           + '\nData: ' + str(date) \
           + '\nBody: ' + '\n\n' + (msg.text or html.unescape(regex.sub('<.*?>', '', msg.html))) + '\n'

      with open(Path(__file__).with_name('body.txt'), 'w', encoding='utf-8') as f:
         f.write(body)
         f.close()

      ftp.cwd('/Arquivos E-mail')

      try:
         ftp.mkd(this_email)
      except Exception:
         continue

      ftp.cwd(this_email)

      file = open(Path(__file__).with_name('body.txt'), 'rb', )
      ftp.storbinary('STOR ' + 'body.txt', file)
      file.close()

      print ('\n' + 20 * "=====" + '\nEmail #' + str(i) + ' Date: ' + str(date) + ' | From: ' + msg.from_ + ' ⇒ ' + msg.subject[0:40])

      if not msg.attachments:
         print('File: 0 attachments')

      for att in msg.attachments:
         if not att.filename or len(att.filename) < 3:
            print('File: 0 attachments')
            continue

         filename = att.filename.replace('\r\n', '').replace('\n', '')
         filename = unidecode.unidecode(filename)

         print('File: ', filename)

         file = open(filename, 'wb')
         file.write(att.payload)

         sendFile = open(file.name, 'rb')
         ftp.storbinary('STOR ' + filename, sendFile)

         sendFile.close()
         file.close()
         os.remove(file.name)


   ftp.quit()

   if all:
      with txt.open('w') as f:
         f.write(last_email)
         f.close()
   else:
      log = Path(__file__).with_name('cron.log')
      with log.open("rb") as file:
         file.seek(-68, 2)

         last_input = file.readline().decode('latin1')
         check = 'AutenLab BOT: VocÃª ainda nÃ£o recebeu nenhum E-mail hoje'

         if not check in last_input:
            print('\n', utc.strftime("%d/%m/%Y %H:%M:%S"),' ♦ AutenLab BOT: Você ainda não recebeu nenhum E-mail hoje ♦ \n' )