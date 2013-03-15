#!/usr/bin/python
# -*- coding: utf-8 -*-
#
#-------------------------------------------------------------------
# webmail.application
#
# Implementation of a command line webmail interface. 
#
# Author: Lee Supe
# Date: March 11th, 2013
#
#-------------------------------------------------------------------

import datetime
import email.utils
import getopt
import getpass
import os
import re
import sys
import time
import unicodedata

from string import Template

from parsedatetime import parsedatetime as pdt

from .client import *
from .data import parse_json

#-------------------------------------------------------------------
DEFAULT_CONFIG = {
      'imap_hostname':           'imap.gmail.com',
      'imap_port':               993,
      'imap_ssl':                True,

      'imap_username':           None,
      'imap_password':           None,
      'imap_mailbox':            'INBOX',

      'cache_dir':               '~/.webmail/',
      'cache_enabled':           True,
      'cache_alt_name':          None,

      'smtp_hostname':           'smtp.gmail.com',
      'smtp_port':               587,
      'smtp_mode':               'tls',
      'smtp_text_enc':           'us-ascii',

      'verbose':                 0,
      'supress_summary':         False,
      'line_limit':              120,
      'line_format':             '[$status] $uid <>    <$sender_name>',
      'date_format':             'oranges',

      'normalize_enabled':       True,
      'normalize_form':          'NFD',
      'print_encoding':          'ascii',
      'print_encoding_rule':     'ignore',
      'file_encoding':           'utf8'
}

DEFAULT_CONFIG_FILENAMES = ["/etc/webmail.json", "~/.webmail.json"]

_G_SHORTOPTS = 'c:i:l:s'
_G_LONGOPTS = ['config=', 'imap-host=', 'imap-port=',
               'imap-user=', 'imap-password=',
               'inbox=', 'limit=', 'supress']

#-------------------------------------------------------------------
class BaseCommand (object):
   """
      The base command object.  Loads settings from config files
      in a standard way and coordinates other global settings.
   """

   #----------------------------------------------------------------
   def __init__ (self, argv, shortopts, longopts, config):
      self.argv = argv
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
            self.config ['supress_summary'] = True

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
   def cache_load_message (self, uid):
      """
         Load a message item from the cache if it exists.  Returns None if
         the file does not exist in the cache or the cache is disabled.
         
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
         filename = os.path.join (os.path.expanduser (self.config ['cache_dir']), "%s.webmail" % uid)
         raw_message = None

         with open (filename, 'rb') as in_file:
            raw_message = in_file.read ()
         
         message = pyzmail.PyzMessage.factory (raw_message) 
         return message
   
      except FileNotFoundError as e:
         return None

   #----------------------------------------------------------------
   def load_message (self, client, uid):
      """
         Loads an emessage from the cache directory or downloads
         the emessage from the server if not found or if the cache
         is not enabled.
         
         Config Settings:
            cache_enabled:
               If this setting is False, load_message will not attempt to
               load the message from the email cache.
      """

      if not self.config ['cache_enabled']:
         return client.fetch_message (uid)

      message = self.cache_load_message (uid)

      if not message:
         message = client.fetch_message (uid)
         self.cache_save_message (uid, message)

      return message

   #----------------------------------------------------------------
   def cache_save_message (self, uid, message):
      """
         Save the given message to the email cache.
         
         Config Settings:
            cache_dir:
               The directory under which emails are written.
            file_encoding:
               The encoding format used to store emails.
      """
      try:
         filename = os.path.join (os.path.expanduser (self.config ['cache_dir']), "%s.webmail" % uid)
         with open (filename, 'wb') as out_file:
            out_file.write (message.as_string ().encode (self.config ['file_encoding']))
         
      except Exception as e:
         print ("Could not save message %s to cache: %s" % (uid, e), file = sys.stderr)

   #----------------------------------------------------------------
   def print_message_status (self, uid, message):
      """
         Print a line containing information about the message.

         Config Settings:
            line_limit:
               The maximum length of the message summary, or "subject" line
               for visibility in console printing.
      """
      
      line = None

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
            date = date.strftime (self.config ['date_format'])) 
      
      ellipsis = '...'
      if self.config ['print_encoding'] != 'ascii':
         ellipsis = '\u2026'

      if '<>' in pre_line:
         if self.config ['line_limit'] is not None:
            max_len = self.config ['line_limit'] - len (pre_line) - 2
            if len (subject) > max_len:
               subject = subject [0:max_len - len (ellipsis)] + ellipsis
            else:
               subject = subject.ljust (max_len)

         line = pre_line.replace ('<>', subject) 

      else:
         line = pre_line

      if self.config ['line_limit'] is not None:
         if len (line) > self.config ['line_limit']:
            line = line [0:(self.config ['line_limit'] - len (ellipsis))] + ellipsis
      
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
                  'unkeyword=', 'unseen']

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
class CheckMailCommand (BaseQueryCommand):
   """
      Command to check the mailbox for new messages
      and print them in a list with uid and subject.
   """

   #----------------------------------------------------------------
   def __init__ (self, argv):
      SHORTOPTS   = 'su:p:l:h:i:p'
      LONGOPTS    = ['username=', 'password=', 'limit=', 'host=', 'inbox=', 'port=', 'no-ssl']

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
         elif opt in ['-h', '--host']:
            self.config ['imap_hostname'] = str (val)
         elif opt in ['-i', '--inbox']:
            self.config ['imap_mailbox'] = str (val)
         elif opt in ['-p', '--port']:
            self.config ['imap_port'] = int (val)
         elif opt in ['--no-ssl']:
            self.config ['imap_ssl'] = False

   #----------------------------------------------------------------
   def run (self):
      while not self.config ['imap_username'] or \
            not self.config ['imap_username'].strip ():
         sys.stdout.write ("Username: ")
         self.config ['imap_username'] = input ()

      while not self.config ['imap_password']:
         self.config ['imap_password'] = getpass.getpass ()

      client = MailClient ()
      client.connect (
            self.config ['imap_username'],
            self.config ['imap_password'],
            self.config ['imap_hostname'],
            self.config ['imap_port'],
            self.config ['imap_ssl'])

      client.set_mailbox (self.config ['imap_mailbox'], True)
      
      # If no query is specified, display new messages.
      if len (self.query.phrases) < 1:
         uids = client.fetch_unread_ids ()
         uids.reverse ()

         if not self.config ['supress_summary']:
            print ("%d new message(s)." % len (uids))

         if self.config ['limit'] is not None:
            uids = uids [:self.config ['limit']]
         
         for uid in uids:
            message = self.load_message (client, uid)
            self.print_message_status (uid, message)

      # If a query is specified, perform a search and return
      # the results
      else:
         print ("QUERY: %s" % str (self.query))

         uids = client.search (self.query)
         uids.reverse ()

         if not self.config ['supress_summary']:
            print ("%d message(s) found." % len (uids))

         if self.config ['limit'] is not None:
            uids = uids [:self.config ['limit']]

         for uid in uids:
            message = self.load_message (client, uid)
            self.print_message_status (uid, message)

#-------------------------------------------------------------------
if __name__ == "__main__":
   app = CheckMailCommand (sys.argv[1:])
   app.run ()

