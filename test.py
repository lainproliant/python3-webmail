import sys

from getpass import getpass
from webmail.client import *

client = MailClient ()
client.connect ("lain.proliant@gmail.com", getpass ())
client.set_mailbox ("INBOX", True)

print ("%d new emails." % client.fetch_unread_count ())

ids = client.fetch_unread_ids ()
print (ids)

for id in ids:
   message = client.fetch_message (id)
   text_part = message.text_part
   
   sys.stdout.buffer.write (("%s\t%s" % (id, message.get_subject ())).encode ())
   print ()
   print ("-"*70)
   sys.stdout.buffer.write (text_part.get_payload ())
   print ()
