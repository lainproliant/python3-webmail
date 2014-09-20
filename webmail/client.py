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
class MailClient ():
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
      if response[0] is None:
         return None

      return response[0][1]

   #----------------------------------------------------------------
   def fetch_message_size (self, id):
      """
         Fetch the size of the given message in bytes.
      """

      status, response = self.imap.uid ('fetch', id, '(RFC822.SIZE)')
      if response[0] is None:
         return None

      return int (response[0].split ()[-1][:-1])

   #----------------------------------------------------------------
   def fetch_message_headers (self, id):
      """
         Fetch the size of the given message in bytes.
      """

      status, response = self.imap.uid ('fetch', id, '(RFC822.HEADER)')

      if response[0] is None:
         return None

      return pyzmail.PyzMessage.factory (response[0][1])

   #----------------------------------------------------------------
   def fetch_message (self, id):
      """
         Fetch the given message from the server and parse it
         into a PyzMessage structure.
      """

      body = self.fetch_message_body (id)

      if body is None:
         return None

      return pyzmail.PyzMessage.factory (self.fetch_message_body (id))

   #----------------------------------------------------------------
   def fetch_unread_ids (self):
      """
         Fetch a string list of the UIDs of all messages in the
         current mailbox marked as unread.
      """
      status, id_pairs = self.imap.uid ('search', '(UNSEEN)')
      ids = []

      for id_pair in id_pairs:
         ids.extend ((id_pair).decode ().split ())

      return ids

   #----------------------------------------------------------------
   def search (self, q):
      """
         Search the IMAP inbox with a query constructed from
         the given IMAPQuery object.  Returns a list of UIDs
         for messages matching the given criterion.
      """

      status, id_pairs = self.imap.uid ('search', str (q))
      ids = []

      for id_pair in id_pairs:
         ids.extend ((id_pair).decode ().split ())

      return ids

   #----------------------------------------------------------------
   def flag (self, uid, *flags):
      """
         Mark a message in an IMAP inbox with the given flags.
      """

      self.imap.uid ('store', uid, '+FLAGS', *flags)

   #----------------------------------------------------------------
   def unflag (self, uid, *flags):
      """
         Remove the given flags from a message in an IMAP inbox.
      """

      self.imap.uid ('store', uid, '-FLAGS', *flags)

   #----------------------------------------------------------------
   def set_mailbox (self, mailbox, readonly = False):
      if not self.is_connected ():
         raise MailClientException ("Cannot set mailbox, not connected.")

      self.mailbox = mailbox
      status, message = self.imap.select ("\"%s\"" % self.mailbox, readonly)

      if status == 'NO':
         raise MailClientException ("Could not change mailboxes: %s" % message)

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
class IMAPQuery ():
   """
      An object providing contextual methods to
      build an IMAP search query.
   """

   #----------------------------------------------------------------
   def __init__ (self, phrases = []):
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

      return IMAPQuery (self.phrases + [args])

   #----------------------------------------------------------------
   def extend_query (self, query):
      """
         Extend this query with all of the phrases from
         the given query.
      """

      return IMAPQuery (self.phrases + query.phrases)

   #----------------------------------------------------------------
   def all (self):
      return self.extend ("ALL")

   #----------------------------------------------------------------
   def answered (self):
      return self.extend ("ANSWERED")

   #----------------------------------------------------------------
   def bcc (self, addr):
      return self.extend ("BCC", addr)

   #----------------------------------------------------------------
   def before (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ValueError ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("BEFORE", dstr)

   #----------------------------------------------------------------
   def body (self, s):
      return self.extend ("BODY", "\"%s\"" % s)

   #----------------------------------------------------------------
   def cc (self, addr):
      return self.extend ("CC", addr)

   #----------------------------------------------------------------
   def contains (self, s):
      q = IMAPQuery ()
      return self.or_q (q.subject (s), q.text (s))

   #----------------------------------------------------------------
   def deleted (self):
      return self.extend ("DELETED")

   #----------------------------------------------------------------
   def draft (self):
      return self.extend ("DRAFT")

   #----------------------------------------------------------------
   def gmail_search (self, s):
      return self.extend ("X-GM-RAW", "\"%s\"" % s)

   #----------------------------------------------------------------
   def flagged (self):
      return self.extend ("FLAGGED")

   #----------------------------------------------------------------
   def from_q (self, s):
      return self.extend ("FROM", "\"%s\"" % s)

   #----------------------------------------------------------------
   def header (self, header, s):
      return self.extend ("HEADER", header, "\"%s\"" % s)

   #----------------------------------------------------------------
   def keyword (self, s):
      return self.extend ("KEYWORD", "\"%s\"" % s)

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
         raise ValueError ("Missing dstr or dt parameter.")

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
         raise ValueError ("There must be only one phrase in each query parameter to IMAPQuery.or_q.")

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
         raise ValueError ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("SENTBEFORE", dstr)

   #----------------------------------------------------------------
   def sent_on (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ValueError ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("SENTON", dstr)

   #----------------------------------------------------------------
   def sent_since (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ValueError ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("SENTSINCE", dstr)

   #----------------------------------------------------------------
   def since (self, dstr = None, dt = None):
      if dstr is None and dt is None:
         raise ValueError ("Missing dstr or dt parameter.")

      if dt is not None:
         dstr = dt.strftime (IMAP_DATE_FORMAT)

      return self.extend ("SINCE", dstr)

   #----------------------------------------------------------------
   def smaller (self, n):
      return self.extend ("SMALLER", str (n))

   #----------------------------------------------------------------
   def subject (self, s):
      return self.extend ("SUBJECT", "\"%s\"" % s)

   #----------------------------------------------------------------
   def text (self, s):
      return self.extend ("TEXT", "\"%s\"" % s)

   #----------------------------------------------------------------
   def to_q (self, s):
      return self.extend ("TO", "\"%s\"" % s)

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
      return self.extend ("UNKEYWORD", "\"%s\"" % s)

   #----------------------------------------------------------------
   def unseen (self):
      return self.extend ("UNSEEN")

