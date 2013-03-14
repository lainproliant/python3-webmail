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
IMAP_DATE_FORMAT = "%d-%b-%Y"

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
      if not self.is_connected ():
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

#-------------------------------------------------------------------
class IMAPQuery (object):
   """
      An object providing contextual methods to
      build an IMAP search query.
   """

   #----------------------------------------------------------------
   def __init__ (self, client, phrases = []):
      self.client = client
      self.phrases = phrases

   #----------------------------------------------------------------
   def __str__ (self):
      sb = []

      for phrase in self.phrases:
         sb.extend (phrase)

      return ' '.join (sb)
   
   #----------------------------------------------------------------
   def extend (self, *args):
      """
         Extend this query with the search parameters
         in the given query, and return a new query object.
      """

      return IMAPQuery (self.client,
            self.phrases + [args])

   #----------------------------------------------------------------
   def extend_query (self, query):
      """
         Extend this query with all of the phrases from
         the given query.
      """

      return IMAPQuery (self.client,
            self.phrases + query.phrases)

   #----------------------------------------------------------------
   def all (self):
      return self.extend ("ALL")

   #----------------------------------------------------------------
   def answered (self):
      return self.extend ("ANSWERED")

   #----------------------------------------------------------------
   def mail_bcc (self, addr):
      return self.extend ("BCC", addr)

   #----------------------------------------------------------------
   def before (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ArgumentException ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("BEFORE", dstr)

   #----------------------------------------------------------------
   def body (self, s):
      return self.extend ("BODY", s)

   #----------------------------------------------------------------
   def mail_cc (self, addr):
      return self.extend ("CC", addr)

   #----------------------------------------------------------------
   def contains (self, s):
      q = IMAPQuery (self.client)
      return self.or_q (q.subject (s), q.text (s))

   #----------------------------------------------------------------
   def deleted (self):
      return self.extend ("DELETED")

   #----------------------------------------------------------------
   def draft (self):
      return self.extend ("DRAFT")

   #----------------------------------------------------------------
   def flagged (self):
      return self.extend ("FLAGGED")

   #----------------------------------------------------------------
   def mail_from (self, s):
      return self.extend ("FROM", s)

   #----------------------------------------------------------------
   def header (self, header, s):
      return self.extend ("HEADER", header, s)

   #----------------------------------------------------------------
   def keyword (self, s):
      return self.extend ("KEYWORD", s)

   #----------------------------------------------------------------
   def larger (self, n):
      return self.extend ("LARGER", str (n))

   #----------------------------------------------------------------
   def new (self):
      return self.extend ("NEW")

   #----------------------------------------------------------------
   def not_q (self, query):
      """
         Extend the query with each phrase from the
         given query, prepended with "NOT".
      """
      
      q = self

      for phrase in query.phrases:
         q = q.extend ("NOT", *phrase)

      return q

   #----------------------------------------------------------------
   def old (self):
      return self.extend ("OLD")

   #----------------------------------------------------------------
   def on (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ArgumentException ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("ON", dstr)

   #----------------------------------------------------------------
   def or_q (self, queryA, queryB):
      """
         Extend the query with "OR", then the single phrases
         from each query.  An exception is thrown if either
         query contains more than one phrase.
      """

      q = self.extend ("OR")

      if len (queryA.phrases) != 1 or len (queryB.phrases) != 1:
         raise ArgumentException ("There must be only one phrase in each query parameter to IMAPQuery.or_q.")

      for phrase in queryA.phrases:
         q = q.extend (*phrase)

      for phrase in queryB.phrases:
         q = q.extend (*phrase)

      return q

   #----------------------------------------------------------------
   def recent (self):
      return self.extend ("RECENT")

   #----------------------------------------------------------------
   def seen (self):
      return self.extend ("SEEN")

   #----------------------------------------------------------------
   def sent_before (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ArgumentException ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("SENTBEFORE", dstr)

   #----------------------------------------------------------------
   def sent_on (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ArgumentException ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("SENTON", dstr)

   #----------------------------------------------------------------
   def sent_since (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ArgumentException ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("SENTSINCE", dstr)

   #----------------------------------------------------------------
   def since (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ArgumentException ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("SINCE", dstr)

   #----------------------------------------------------------------
   def smaller (self, n):
      return self.extend ("SMALLER", str (n))

   #----------------------------------------------------------------
   def subject (self, s):
      return self.extend ("SUBJECT", s)

   #----------------------------------------------------------------
   def text (self, s):
      return self.extend ("TEXT", s)

   #----------------------------------------------------------------
   def mail_to (self, s):
      return self.extend ("TO", s)

   #----------------------------------------------------------------
   def uid (self, n):
      return self.extend ("UID", str (n))

   #----------------------------------------------------------------
   def unanswered (self):
      return self.extend ("UNANSWERED")

   #----------------------------------------------------------------
   def undeleted (self):
      return self.extend ("UNDELETED")

   #----------------------------------------------------------------
   def undraft (self):
      return self.extend ("UNDRAFT")

   #----------------------------------------------------------------
   def unflagged (self):
      return self.extend ("UNDRAFT")

   #----------------------------------------------------------------
   def unkeyword (self, s):
      return self.extend ("UNKEYWORD", s)

   #----------------------------------------------------------------
   def unseen (self):
      return self.extend (unseen)

