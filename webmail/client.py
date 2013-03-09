#-------------------------------------------------------------------
# webmail.client
#
# Client abstraction and utilites for interacting with
# an IMAP email server.
#
# Author: Lee Supe
# Date: March 7th, 2013
#
#-------------------------------------------------------------------

import imaplib
import pyzmail

#-------------------------------------------------------------------
class MailClientException (Exception):
   def __init__ (self, message):
      Exception.__init__ (self, message)

#-------------------------------------------------------------------
class MailClient (object):
   """
      An object for accessing and manipulating messages in an
      IMAP email inbox.
   """

   #----------------------------------------------------------------
   def __init__ (self):
      self.imap = None
      self.mailbox = None

   #----------------------------------------------------------------
   def connect (self, username, password,
                hostname = "imap.gmail.com",
                port = 993, ssl = True):
      if ssl:
         self.imap = imaplib.IMAP4_SSL (hostname, port)
         self.imap.login (username, password)
      else:
         raise NotImplementedError ("Non-SSL IMAP connection not yet implemented.")


   #----------------------------------------------------------------
   def fetch_message_body (self, id):
      """
         Fetch the raw RFC822 body for the given message UID.
      """

      status, response = self.imap.uid ('fetch', id, '(RFC822)')
      return response[0][1]

   #----------------------------------------------------------------
   def fetch_message (self, id):
      """
         Fetch the given message from the server and parse it
         into a PyzMessage structure.
      """

      return pyzmail.PyzMessage.factory (self.fetch_message_body (id))

   #----------------------------------------------------------------
   def fetch_unread_count (self):
      """
         Fetch the unread message count of the current mailbox.
         If no mailbox is specified, it may fetch the unread mail
         count for all mailboxes in the account.
      """

      status, response = self.imap.status (self.get_mailbox (), '(UNSEEN)')
      unread_count = int (response[0].decode ().split ()[2].strip (').,]'))
      return unread_count

   #----------------------------------------------------------------
   def fetch_unread_ids (self):
      """
         Fetch a string list of the IDs of all messages in the
         current mailbox marked as unread.
      """
      status, id_pairs = self.imap.uid ('search', '(UNSEEN)')
      ids = []

      for id_pair in id_pairs:
         ids.extend ((id_pair).decode ().split ())

      return ids 
   
   #----------------------------------------------------------------
   def set_mailbox (self, mailbox, readonly = False):
      if not self.is_connected:
         raise MailClientException ("Cannot set mailbox, not connected.")

      self.mailbox = mailbox
      self.imap.select (self.mailbox, readonly)

   #----------------------------------------------------------------
   def get_mailbox (self):
      return self.mailbox

   #----------------------------------------------------------------
   def is_connected (self):
      if self.imap is None:
         return False
      else:
         return True

