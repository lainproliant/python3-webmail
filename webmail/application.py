#-------------------------------------------------------------------
# webmail.application
#
# Implementation of a command line webmail interface. 
#
# Author: Lee Supe
# Date: March 11th, 2013
#-------------------------------------------------------------------

import datetime
import email.utils
import getopt
import getpass
import mimetypes
import os
import re
import subprocess
import sys
import time
import unicodedata

from string import Template
from tempfile import NamedTemporaryFile
from textwrap import TextWrapper

from parsedatetime import parsedatetime as pdt

from .client import *
from .data import parse_json

#-------------------------------------------------------------------
DEFAULT_CONFIG = {
      'account':                 'default',

      'imap_hostname':           'imap.gmail.com',
      'imap_port':               993,
      'imap_ssl':                True,

      'imap_username':           None,
      'imap_password':           None,
      'imap_mailbox':            'INBOX',

      'cache_dir':               '~/.webmail/',
      'cache_enabled':           True,
      'download_threshold':      100000,

      'smtp_hostname':           'smtp.gmail.com',
      'smtp_port':               587,
      'smtp_mode':               'tls',
      'smtp_text_enc':           'us-ascii',

      'verbose':                 0,
      'supress':                 False,
      'interactive':             True,
      'line_width':              120,
      'line_format':             '[$status] $uid <>    <$sender_name> $date',
      'st_date_format_recent':   '%b %d %l:%M%P',
      'st_date_format_far':      '%b %d %Y %l:%M%P',

      'date_format':             '%B %d %Y, %l:%M%P',
      'normalize_enabled':       False,
      'normalize_form':          'NFD',
      'print_encoding':          'utf8',
      'print_encoding_rule':     'ignore',
      'file_encoding':           'utf8',

      'mime:text/plain':         "PRINT",
      'mime:text/html':          'chromium --user-data-dir=/tmp %s',
      'mime:application/pdf':    'evince %s',
      'mime:image/*':            'geeqie %s',
      'mime:video/*':            'vlc %s',
      
      'debug':                   False,

      'default':                 {}
}

DEFAULT_CONFIG_FILENAMES = ["/etc/webmail.json", "~/.webmail.json"]

_G_SHORTOPTS = 'c:i:l:sa:NC'
_G_LONGOPTS = ['config=', 'imap-host=', 'imap-port=',
               'imap-user=', 'imap-password=',
               'inbox=', 'limit=', 'supress',
               'account=', 'debug', 'no-prompt',
               'no-cache']

#-------------------------------------------------------------------
class ThresholdExceeded (Exception):
   def __init__ (self):
      Exception.__init__ (self, "A threshold was exceeded.")

#-------------------------------------------------------------------
class MailpartHandler (object):
   """
      An object for opening and saving mailparts.
   """
   
   def __init__ (self, config, part, part_num):
     self.config = config
     self.part = part
     self.part_num = part_num

   def get_file_extension (self):
      """
         Determine an appropriate file extension for the content
         based on the filename or guess the extension based on 
         the content MIME type.
      """

      if self.part.filename is None:
         return mimetypes.guess_extension (self.part.type)

      else:
         return os.path.splitext (self.part.filename)[1]

   def open (self):
      """
         Open the mailpart for viewing by the user.  This could
         result in the execution of a subprocess or simply printing
         the decoded mailpart to standard out.

         The item is opened and handled based on its MIME type,
         and configured via 'mime:<type>' mapped to a command
         or the word "PRINT", e.g.:

         {
            ...
            'mime:text/plain':      "PRINT",
            'mime:text/html':       "cat %s | w3m",
            'mime:image/*':         "feh %s",
            ,,,
         }
      """

      # See if there is a handler for the specific MIME type.
      # Then, check for a generic handler based on the general
      # MIME type (e.g., 'image/*').

      handler_key = 'mime:%s' % self.part.type
      handler = None
      
      if handler_key not in self.config:
         handler_key = 'mime:%s/*' % (self.part.type.split ('/')[0])
         if handler_key not in self.config:
            raise Exception ("No file handler configured to open mimetype: %s." % self.part.type)

      handler = self.config [handler_key]
      
      if handler == 'PRINT':
         tw = TextWrapper (
               break_long_words = False,
               replace_whitespace = False)

         charset = self.part.charset

         if charset is None:
            print ("Warning: mailpart %d does not specify a charset, assuming 'utf8'." % self.part_num)
            charset = 'utf8'

         print ()
         print ('-' * self.config ['line_width'])
         lines = tw.wrap (self.part.get_payload ().decode (charset))
         for line in lines:
            print (line)

      else:
         # Save the mailpart payload to a file for further processing.
         tmpfile = NamedTemporaryFile (suffix = self.get_file_extension (), delete = False)
         tmpfile.write (self.part.get_payload ())
         tmpfile.close ()

         handler_cmd = handler % tmpfile.name

         if self.config ['debug']:
            print ("Running: %s" % handler_cmd)

         retcode = subprocess.call (handler_cmd, shell = True)
         os.unlink (tmpfile.name)
         
         if retcode != 0:
            raise Exception ("Error executing MIME handler (%d): %s" % (retcode, handler_cmd))
   
#-------------------------------------------------------------------
class BaseCommand (object):
   """
      The base command object.  Loads settings from config files
      in a standard way and coordinates other global settings.
   """

   PARAMETERS = 0

   #----------------------------------------------------------------
   def __init__ (self, argv, shortopts, longopts, config):
      self.argv = argv
      self.argv_copy = argv[:]
      self.shortopts = _G_SHORTOPTS + shortopts
      self.longopts = _G_LONGOPTS + longopts
      self.optional_config_files = DEFAULT_CONFIG_FILENAMES 
      self.specific_config_files = []
      self.config = DEFAULT_CONFIG
      self.config.update (config)
      
      opts, args = getopt.getopt (self.argv, self.shortopts, self.longopts)
      self.process_config (opts, args)

   #----------------------------------------------------------------
   def process_config (self, opts, args):
      for opt, val in opts:
         if opt in ['-c', '--config']:
            self.specific_config_files.append (str (val))

      self.optional_config_files = [os.path.abspath (os.path.expanduser (x)) for x in self.optional_config_files]
      self.specific_config_files = [os.path.abspath (os.path.expanduser (x)) for x in self.specific_config_files]

      for filename in self.optional_config_files:
         try:
            self.config.update (parse_json (filename))
         except FileNotFoundError as e:
            pass
         
      for filename in self.specific_config_files:
         self.config.update (parse_json (filename))
      
      for opt, val in opts:
         if opt in ['-v', '--verbose']:
            self.config ['verbose'] += 1
         elif opt in ['--imap-user']:
            self.config ['imap_username'] = str (val)
         elif opt in ['--imap-password']:
            self.config ['imap_password'] = str (val)
         elif opt in ['--imap-host']:
            self.config ['imap_hostname'] = str (val)
         elif opt in ['--imap-port']:
            self.config ['imap_port'] = int (val)
         elif opt in ['-i', '--inbox']:
            self.config ['inbox'] = str (val)
         elif opt in ['-s', '--supress']:
            self.config ['supress'] = True
         elif opt in ['-a', '--account']:
            self.config ['account'] = str (val)
         elif opt in ['--debug']:
            self.config ['debug'] = True
         elif opt in ['-N', '--no-prompt']:
            self.config ['interactive'] = False
         elif opt in ['-C', '--no-cache']:
            self.config ['cache_enabled'] = False

      self.process_account_settings (self.config ['account'])

   #----------------------------------------------------------------
   def process_account_settings (self, account):
      """
         Merge settings from an account in the config map.
      """

      if not account in self.config:
         raise Exception ("No settings for account \"%s\" found in the config map." % account)

      self.config.update (self.config [account])

   #----------------------------------------------------------------
   def normalize (self, s):
      """
         Normalize the given Unicode string for printing to a terminal based
         on settings in the config dictionary.
         
         Config Settings:
            normalize_enabled:
               Whether or not normalization should be performed.
            normalize_form:
               The form passed to unicodedata.normalize
            print_encoding:
               The Unicode encoding to use for printing to consoles.
            print_encoding_rule:
               The replacement rule for encoding, see str.encode.
      """
      if self.config ['normalize_enabled']:
         return unicodedata.normalize (self.config ['normalize_form'], s).encode (
               self.config ['print_encoding'], self.config ['print_encoding_rule']).decode ()
      else:
         return s

   #----------------------------------------------------------------
   def cache_filename_for_uid (self, uid):
      """
         Determines the absolute path of a cache file for the given UID.
      """

      cache_dir = os.path.join (
            os.path.expanduser (self.config ['cache_dir']),
            self.config ['account'])
      filename = os.path.join (cache_dir, "%s.webmail" % uid)
      return filename

   #----------------------------------------------------------------
   def cache_has_message (self, uid):
      """
         Determine if the given message is cached in the message cache.

         uid:
            The message to check for in the cache.
         Config Settings:
            cache_enabled:
               Whether emails should be written to a directory under cache_dir
               upon being fetched for later recall.  If this is False, this
               method will always return False.
            cache_dir:
               The directory under which emails are written.  Emails will be written
               to a directory under cache_dir based on the account from which they
               were read.

      """
      
      if not self.config ['cache_enabled']:
         return False

      filename = self.cache_filename_for_uid (uid)
      return os.path.exists (filename)

   #----------------------------------------------------------------
   def cache_fetch_message (self, uid):
      """
         Load a message item from the cache if it exists.  Returns None if
         the file does not exist in the cache or the cache is disabled.

         uid:
            The message to load from cache.
         
         Config Settings:
            cache_enabled:
               Whether emails should be written to a directory under cache_dir
               upon being fetched for later recall.
            cache_dir:
               The directory under which emails are written.  Emails will be written
               to a directory under cache_dir based on the account from which they
               were read.
      """

      if not self.config ['cache_enabled']:
         return None

      try:
         filename = self.cache_filename_for_uid (uid)
         raw_message = None

         with open (filename, 'rb') as in_file:
            raw_message = in_file.read ()
         
         message = pyzmail.PyzMessage.factory (raw_message) 
         return message
   
      except FileNotFoundError as e:
         return None

   #----------------------------------------------------------------
   def fetch_message (self, client, uid):
      """
         Loads a message from the cache directory or downloads
         the message from the server if not found or if the cache
         is not enabled.

         client:
            A MailClient used to fetch messages.
         uid:
            The UID of the message to be fetched.

         Config Settings:
            cache_enabled:
               If this setting is False, fetch_message will not attempt to
               load the message from the email cache, nor will it save
               messages to the email cache.
      """
      
      message = None

      if self.config ['cache_enabled']:
         message = self.cache_fetch_message (uid)

      if not message:
         message = client.fetch_message (uid)
         if message is not None and self.config ['cache_enabled']:
            self.cache_save_message (uid, message)

      return message

   #----------------------------------------------------------------
   def fetch_message_headers (self, client, uid, threshold = None):
      """
         Loads message headers from the IMAP server and caches
         messages if the cache is enabled and the threshold
         is not exceeded.

         client:
            A MailClient used to fetch messages.
         uid:
            The UID of the message to be fetched.
         threshold:
            If specified, the maximum size in bytes of a message
            to be fetched.  If the message is larger it will
            not be cached.
         
         Config Settings:
            cache_enabled:
               If this setting is False, fetch_message will not attempt to
               load the message from the email cache.
      """
      
      size = client.fetch_message_size (uid)

      if size is None:
         # The message doesn't exist on the IMAP server.
         return None

      elif self.config ['cache_enabled'] and (threshold is None or size < threshold):
         if not self.cache_has_message (uid):
            message = client.fetch_message (uid)
            if message is not None:
               self.cache_save_message (uid, message)

      return client.fetch_message_headers (uid)

   #----------------------------------------------------------------
   def cache_save_message (self, uid, message):
      """
         Save the given message to the email cache.
         
         uid:
            The UID of the message to be saved.
         message:
            A PyzMessage object representing the message.

         Config Settings:
            cache_dir:
               The directory under which emails are written.
            file_encoding:
               The encoding format used to store emails.
      """
      try:
         cache_dir = os.path.join (
               os.path.expanduser (self.config ['cache_dir']),
               self.config ['account'])

         if not os.path.isdir (cache_dir):
            try:
               os.makedirs (cache_dir)
            except OSError as e:
               raise Exception ("Unable to create cache directory: %s" % str (e))

         filename = os.path.join (cache_dir, "%s.webmail" % uid)
         with open (filename, 'wb') as out_file:
            out_file.write (message.as_string ().encode (self.config ['file_encoding']))
         
      except Exception as e:
         print ("Could not save message %s to cache: %s" % (uid, e), file = sys.stderr)
         raise e

   #----------------------------------------------------------------
   def perform_imap_login (self):
      """
         Attempt to perform a login with the current settings,
         prompting the user if necessary for a username
         and password.
      """

      while not self.config ['imap_username'] or \
            not self.config ['imap_username'].strip ():
         if not self.config ['interactive']:
            raise Exception ("No imap_username configured.")
         sys.stdout.write ("Username: ")
         self.config ['imap_username'] = input ()

      while not self.config ['imap_password']:
         if not self.config ['interactive']:
            raise Exception ("No imap_password configured.")
         self.config ['imap_password'] = getpass.getpass ()

      client = MailClient ()
      client.connect (
            self.config ['imap_username'],
            self.config ['imap_password'],
            self.config ['imap_hostname'],
            self.config ['imap_port'],
            self.config ['imap_ssl'])

      return client

   #----------------------------------------------------------------
   def print_message_status (self, client, uid):
      """
         Print a line containing information about the message.

         client:
            A logged-in MailClient object with a mailbox selected.
         uid:
            The UID of the message for which status will be displayed.

         Config Settings:
            download_threshold:
               If a message does not exceed this threshold size in bytes,
               it will be downloaded and cached locally.  Otherwise,
               IMAP queries will be used to fetch the information
               necessary to display the message's status.
            st_date_format_recent:
               The formatting of dates in the status line, as accepted by
               datetime.datetime.strftime ().
            line_format:
               A format string used to display the status line.
               This format string may have any of the following:
                  uid:
                     The UID of the message.
                  date:
                     The date of the message, as formatted by 'st_date_format_recent'.
                  sender_name:
                     The name (e.g., "Lee Supe") of the sender.
                  sender_addr:
                     The address (e.g., "lain.proliant@gmail.com") of the sender.
                  status:
                     A single character representing the status of the message.
                     '!' for unread messages, blank for read messages.
                  subject:
                     The subject of the message.  This may also be specified
                     as "<>", in which case the subject will be displayed and
                     sized to fill all remaining space not occupied by other fields.
            line_width:
               The maximum length of the message summary, or "subject" line
               for visibility in console printing.
            print_encoding:
               The encoding used to output text.  This should usually remain 'utf8'
               unless you are printing to a console that does not support unicode.
               Setting this property to 'ascii' will disable Unicode output and
               cause "..." to be printed instead of an ellipsis.

      """
      
      line = None
      
      threshold = self.config ['download_threshold']
      message = self.fetch_message_headers (client, uid, threshold = threshold)
      
      if message is None: 
         raise NotImplementedError ("Message partial fetching not yet implemented.")

      else:
         from_addresses = message.get_addresses ('from')
         sender_name = from_addresses[0][0]
         sender_addr = from_addresses[0][1]
         status = '!'
         subject = self.normalize (message.get_subject ())
         subject = re.sub ('\s+', ' ', subject)
         subject = re.sub ('\r|\n', '', subject)

         date_ts = email.utils.parsedate_tz (message.get_decoded_header ('Date'))
         date = datetime.datetime.fromtimestamp (email.utils.mktime_tz (
            date_ts))

      
      template = Template (self.config ['line_format'])
      pre_line = template.substitute (
            uid = uid,
            sender_name = sender_name,
            sender_addr = sender_addr,
            status = status,
            subject = subject,
            date = date.strftime (self.config ['st_date_format_recent'])) 
      
      ellipsis = '...'
      if self.config ['print_encoding'] != 'ascii':
         ellipsis = '\u2026'

      if '<>' in pre_line:
         if self.config ['line_width'] is not None:
            max_len = self.config ['line_width'] - len (pre_line) - 2
            if len (subject) > max_len:
               subject = subject [0:max_len - len (ellipsis)] + ellipsis
            else:
               subject = subject.ljust (max_len)

         line = pre_line.replace ('<>', subject) 

      else:
         line = pre_line

      if self.config ['line_width'] is not None:
         if len (line) > self.config ['line_width']:
            line = line [0:(self.config ['line_width'] - len (ellipsis))] + ellipsis
      
      print (line)

   #----------------------------------------------------------------
   def run (self):
      raise NotImplementedException ("Abstract base method called.")

#-------------------------------------------------------------------
class BaseQueryCommand (BaseCommand):
   """
      A base class for commands which need to deal with
      queries against an IMAP mailbox.
   """

   #----------------------------------------------------------------
   def __init__ (self, argv, shortopts, longopts, config):
      SHORTOPTS = ''
      LONGOPTS = ['all', 'answered', 'bcc=', 'before=',
                  'body=', 'cc=', 'contains=', 'deleted',
                  'draft', 'flagged', 'from=', 'keyword=',
                  'larger=', 'new', 'not', 'old', 'on=',
                  'or', 'recent', 'seen', 'sent-before=',
                  'sent-on=', 'sent-since=', 'since=',
                  'smaller=', 'subject=', 'text=',
                  'to=', 'uid=', 'unanswered',
                  'undeleted', 'undraft', 'unflagged',
                  'unkeyword=', 'unseen', 'gmail=']

      self.query = IMAPQuery ()

      BaseCommand.__init__ (self, argv, shortopts + SHORTOPTS, longopts + LONGOPTS, config)

   #----------------------------------------------------------------
   def process_config (self, opts, args):
      BaseCommand.process_config (self, opts, args)

      q = self.query
      
      or_found = False
      not_found = False
      not_flag = False
      
      or_stack = []
      
      n = 0

      for opt, val in opts:
         n += 1
         
         if opt == '--all':
            q = q.all ()
         elif opt == '--answered':
            q = q.answered ()
         elif opt == '--bcc':
            q = q.bcc (str (val))
         elif opt == '--before':
            dstr = self.human_to_imap_date (str (val))
            q = q.before (dstr)
         elif opt == '--body':
            q = q.body (str (val))
         elif opt == '--cc':
            q = q.cc (str (val))
         elif opt == '--contains':
            q = q.contains (str (val))
         elif opt == '--deleted':
            q = q.deleted ()
         elif opt == '--draft':
            q = q.draft ()
         elif opt == '--gmail':
            q = q.gmail_search (str (val))
         elif opt == '--flagged':
            q = q.flagged ()
         elif opt == '--from':
            q = q.from_q (str (val))
         elif opt == '--keyword':
            q = q.keyword (str (val))
         elif opt == '--larger':
            q = q.larger (int (val))
         elif opt == '--new':
            q = q.new ()
         elif opt == '--not':
            not_found = True
         elif opt == '--old':
            q = q.old ()
         elif opt == '--on':
            dstr = self.human_to_imap_date (str (val))
            q = q.on (dstr)
         elif opt == '--or':
            or_found = True
         elif opt == '--recent':
            q = q.recent ()
         elif opt == '--seen':
            q = q.seen ()
         elif opt == '--sent-before':
            dstr = self.human_to_imap_date (str (val))
            q = q.sent_before (dstr)
         elif opt == '--sent-on':
            dstr = self.human_to_imap_date (str (val))
            q = q.sent_on (dstr)
         elif opt == '--sent-since':
            dstr = self.human_to_imap_date (str (val))
            q = q.sent_since (dstr)
         elif opt == '--since':
            dstr = self.human_to_imap_date (str (val))
            q = q.since (dstr)
         elif opt == '--smaller':
            q = q.smaller (int (val))
         elif opt == '--subject':
            q = q.subject (str (val))
         elif opt == '--text':
            q = q.text (str (val))
         elif opt == '--to':
            q = q.to_q (str (val))
         elif opt == '--uid':
            q = q.uid (str (val))
         elif opt == '--unanswered':
            q = q.unanswered ()
         elif opt == '--undeleted':
            q = q.undeleted ()
         elif opt == '--undraft':
            q = q.undraft ()
         elif opt == '--unflagged':
            q = q.unflagged ()
         elif opt == '--unkeyword':
            q = q.unkeyword (str (val))
         elif opt == '--unseen':
            q = q.unseen ()
         
         if or_found:
            or_found = False
            or_stack.append (q.phrases.pop ())
         
         elif not_found:
            not_found = False
            not_flag = True

         elif not_flag:
            not_flag = False
            q = q.not_q (IMAPQuery ([q.phrases.pop ()]))

            if or_stack:
               or_stack.append (q.phrases.pop ())
               q = q.or_q (
                     IMAPQuery ([or_stack.pop ()]),
                     IMAPQuery ([or_stack.pop ()]))

         elif or_stack:
            or_stack.append (q.phrases.pop ())
            q = q.or_q (
                  IMAPQuery ([or_stack.pop ()]),
                  IMAPQuery ([or_stack.pop ()]))

      self.query = q

   #----------------------------------------------------------------
   def human_to_imap_date (self, s):
      cal = pdt.Calendar ()
      ts = time.mktime (time.struct_time (cal.parse (s)[0]))
      dt = datetime.datetime.fromtimestamp (ts)
      return dt.strftime ("%d-%b-%Y")

#-------------------------------------------------------------------
class SearchMailCommand (BaseQueryCommand):
   """
      Search the mailbox or check for new messages, with the option
      to perform operations on search results.

      Usage:
         webmail --search <options / queries / operations>
      
      Options:
         -u, --username <username>     (config: imap_username)
            The IMAP username.  User will be prompted if missing.

         -p, --password <password>     (config: imap_password)
            The IMAP password.  User will be prompted if missing.

         -H, --host <hostname>         (config: imap_hostname)
            The IMAP server to connect to. Default is 'imap.gmail.com'.

         -P, --port <port>             (config: imap_port)
            The IMAP server port. Default is '993'.

         -i, --inbox <inbox>           (config: imap_mailbox)
            The IMAP mailbox from which to read.  Default is 'INBOX'.

         --no-ssl                      (config: imap_ssl)
            Disables IMAP SSL/TLS.  Not recommended, enabled by default.

         -h, --help
            Prints this help message.

      Queries:
         This command supports search queries.  For help regarding
         search queries, type "webmail help queries".

      Operations:
         --flag FLAG
            Flag each message with the given flag.
         
         --unflag FLAG
            Remove the given flag from each message.

      Notes:
         This is the default command.  For a list of available commands,
         type "webmail help".
   """

   #----------------------------------------------------------------
   def __init__ (self, argv):
      SHORTOPTS   = 'su:p:l:H:i:p'
      LONGOPTS    = ['username=', 'password=', 'limit=', 'host=', 'inbox=', 'port=', 'no-ssl',
            'flag=', 'unflag=', 'print']
      
      self.operations = []

      BaseQueryCommand.__init__ (
            self, argv, SHORTOPTS, LONGOPTS, {
               'limit':       None,
            })

   #----------------------------------------------------------------
   def process_config (self, opts, args):
      BaseQueryCommand.process_config (self, opts, args)
      
      for opt, val in opts:
         if opt in ['-u', '--username']:
            self.config ['imap_username'] = str (val)
         elif opt in ['-p', '--password']:
            self.config ['imap_password'] = str (val)
         elif opt in ['-l', '--limit']:
            self.config ['limit'] = int (val)
         elif opt in ['-H', '--host']:
            self.config ['imap_hostname'] = str (val)
         elif opt in ['-i', '--inbox']:
            self.config ['imap_mailbox'] = str (val)
         elif opt in ['-P', '--port']:
            self.config ['imap_port'] = int (val)
         elif opt in ['--no-ssl']:
            self.config ['imap_ssl'] = False
         elif opt in ['--flag', '--unflag']:
            self.operations.append ((opt, val))
         elif opt in ['--print']:
            self.operations.append ((opt, val))
            self.config ['supress'] = False

   #----------------------------------------------------------------
   def run (self):
      client = self.perform_imap_login ()
      
      client.set_mailbox (self.config ['imap_mailbox'], True)

      message = "%d message(s) found."
      
      # If no query is specified, display new messages.
      if len (self.query.phrases) < 1:
         self.query = self.query.unseen ()
         message = "%d new message(s)."
      
      if self.config ['debug']:
         print ("Query: %s" % str (self.query))

      uids = client.search (self.query)
      uids.reverse ()
      
      print (message % len (uids))

      if self.config ['limit'] is not None:
         uids = uids [:self.config ['limit']]

      if not self.operations:
         self.operations.append (('--print', None))
      
      if len (uids) > 0:
         for opt, val in self.operations:
            if opt == '--print' and not self.config ['supress']:
               client.set_mailbox (self.config ['imap_mailbox'], True)

               for uid in uids:
                  self.print_message_status (client, uid)

            if opt == '--flag':
               client.set_mailbox (self.config ['imap_mailbox'], False)

               flag = str (val)
               if flag[0] != '\\':
                  flag = '\\' + str (val)

               client.flag (','.join (uids), flag)
               print ("%d message(s) flagged as %s." % (len (uids), flag))

            elif opt == '--unflag':
               client.set_mailbox (self.config ['imap_mailbox'], False)

               flag = str (val)
               if flag[0] != '\\':
                  flag = '\\' + str (val)

               client.unflag (','.join (uids), flag)
               print ("%s flag removed from %d messages." % (flag, len (uids)))

#-------------------------------------------------------------------
class ReadMailCommand (BaseCommand):
   """
      Read a mail message specified by UID.

      Usage:
         webmail read <options> UID
      
      Options:
         -u, --username <username>     (config: imap_username)
            The IMAP username.  User will be prompted if missing.

         -p, --password <password>     (config: imap_password)
            The IMAP password.  User will be prompted if missing.

         -H, --host <hostname>         (config: imap_hostname)
            The IMAP server to connect to. Default is 'imap.gmail.com'.

         -P, --port <port>             (config: imap_port)
            The IMAP server port. Default is '993'.

         -i, --inbox <inbox>           (config: imap_mailbox)
            The IMAP mailbox from which to read.  Default is 'INBOX'.

         -m, --mailpart <num>
            The mailpart to be inspected.

         --no-ssl                      (config: imap_ssl)
            Disables IMAP SSL/TLS.  Not recommended, enabled by default.

         --peek
            Do not mark the message as read. 

         -h, --help
            Prints this help message.

         UID
            Required. The UID of the message to read.

      For a list of available commands, type "webmail help".
   """

   PARAMETERS = 1

   #----------------------------------------------------------------
   def __init__ (self, argv):
      SHORTOPTS   = 'u:p:H:i:P:m:'
      LONGOPTS    = ['username=', 'password=', 'host=', 'inbox=', 'port=', 'no-ssl', 'mailpart']
      
      self.message_uid = None
      self.message_part = None

      BaseCommand.__init__ (
            self, argv, SHORTOPTS, LONGOPTS, {})

   #----------------------------------------------------------------
   def process_config (self, opts, args):
      BaseCommand.process_config (self, opts, args)
      
      if len (args) < 1:
         raise Exception ("No message UID specified to read.")

      self.message_uid = args.pop (0)

      if args:
         self.message_part = int (args.pop (0))

      for opt, val in opts:
         if opt in ['-u', '--username']:
            self.config ['imap_username'] = str (val)
         elif opt in ['-p', '--password']:
            self.config ['imap_password'] = str (val)
         elif opt in ['-l', '--limit']:
            self.config ['limit'] = int (val)
         elif opt in ['-h', '--host']:
            self.config ['imap_hostname'] = str (val)
         elif opt in ['-i', '--inbox']:
            self.config ['imap_mailbox'] = str (val)
         elif opt in ['-P', '--port']:
            self.config ['imap_port'] = int (val)
         elif opt in ['--no-ssl']:
            self.config ['imap_ssl'] = False
         elif opt in ['-m', '--mailpart']:
            self.message_part = int (val)

   #----------------------------------------------------------------
   def run (self):
      client = self.perform_imap_login ()
      # The readonly flag should be false once we are actually
      # reading messages, so they are flagged as read.
      client.set_mailbox (self.config ['imap_mailbox'])

      message = self.fetch_message (client, self.message_uid)
      if message is None:
         raise Exception ("No message exists with UID %s." % self.message_uid)
      
      client.flag (self.message_uid, '\\Seen')

      from_addr = message.get_address ('from')
      to_addrs = message.get_addresses ('to')
      cc_addrs = message.get_addresses ('cc')

      date_ts = email.utils.parsedate_tz (message.get_decoded_header ('Date'))
      date = datetime.datetime.fromtimestamp (email.utils.mktime_tz (
         date_ts))
      
      print ("Date: %s" % date.strftime (self.config ['date_format']))
      print ("Subject: %s" % self.normalize (message.get_subject ()))
      print ("From: %s <%s>" % from_addr)
      print ("To: %s" % (', '.join (["%s <%s>" % x for x in to_addrs])))
      if cc_addrs:
         print ("CC: %s" % (', '.join (["%s <%s>" % x for x in cc_addrs])))
      print ()
      n = 0
      for part in message.mailparts:
         mailpart_str = "mailpart: %d" % n
         
         if part.disposition is not None:
            mailpart_str = "%s (%s: %s)" % (mailpart_str, part.disposition, part.type)
         else:
            mailpart_str = "%s (%s)" % (mailpart_str, part.type)

         if part.filename is not None:
            mailpart_str = "%s [%s]" % (mailpart_str, part.filename) 
         
         print (mailpart_str)
         n += 1

      print ()
      
      if self.message_part is None:
         if not self.config ['supress']:
            n = 0
            for part in message.mailparts:
               if part.type == 'text/plain' and part.is_body:
                  handler = MailpartHandler (self.config, part, n)
                  handler.open ()
               n += 1

      else:
         part = message.mailparts [self.message_part]
         handler = MailpartHandler (self.config, part, self.message_part)
         handler.open ()
      
#-------------------------------------------------------------------
class CountMailCommand (SearchMailCommand):
   """
      A simple version of the search command which prints a
      number representing the number of results instead of
      a listing of the results.  Useful for scripting.
   """

   #----------------------------------------------------------------
   def run (self):
      client = self.perform_imap_login ()
      
      client.set_mailbox (self.config ['imap_mailbox'], True)

      message = "%d"
      
      # If no query is specified, display new messages.
      if len (self.query.phrases) < 1:
         self.query = self.query.unseen ()
      
      uids = client.search (self.query)
      uids.reverse ()
      
      print (message % len (uids))

#-------------------------------------------------------------------
COMMAND_MAP = {
      "--search":      SearchMailCommand,
      "--count":       CountMailCommand,
      "--read":        ReadMailCommand
}

#-------------------------------------------------------------------
if __name__ == "__main__":
   cmd = SearchMailCommand
   acct_name = None

   argv = sys.argv [1:]

   for arg in argv[:]:
      if arg in COMMAND_MAP:
         cmd = COMMAND_MAP [arg]

         for x in range (cmd.PARAMETERS):
            param = None
            x = argv.index (arg) + 1
            if x < len (argv) and len (argv[x]) > 0 and argv[x][0] != '-':
               param = argv[x]
               argv.pop (x)
               argv.append (param)

         argv.remove (arg)

   try:
      app = cmd (argv)
      app.run ()

   except Exception as e:
      import traceback
      print ("Fatal error: %s" % e, file=sys.stderr)

      if '--debug' in argv:
         print ('-' * 70, file=sys.stderr)
         traceback.print_exc (file=sys.stderr) 
         sys.exit (1)

