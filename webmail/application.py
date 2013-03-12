#-------------------------------------------------------------------
# webmail.application
#
# Implementation of a command line webmail interface. 
#
# Author: Lee Supe
# Date: March 11th, 2013
#
#-------------------------------------------------------------------

import getopt
import getpass
import os
import sys

from .client import *
from .data import parse_json

#-------------------------------------------------------------------
DEFAULT_CONFIG = {
      'imap_hostname':     'imap.gmail.com',
      'imap_port':         993,
      'imap_ssl':          True,

      'imap_username':     None,
      'imap_password':     None,
      'imap_mailbox':      'INBOX',

      'smtp_hostname':     'smtp.gmail.com',
      'smtp_port':         587,
      'smtp_mode':         'tls',
      'smtp_text_enc':     'us-ascii',
      'verbose':           0
}

DEFAULT_CONFIG_FILENAMES = ["/etc/webmail.json", "~/.webmail.json"]

_G_SHORTOPTS = 'c:i:l:q'
_G_LONGOPTS = ['config=', 'imap-host=', 'imap-port=',
               'imap-user=', 'imap-password',
               'inbox=', 'limit=', 'quiet']

#-------------------------------------------------------------------
class BaseCommand (object):
   """
      The base command object.  Loads settings from config files
      in a standard way and coordinates other global settings.
   """

   def __init__ (self, argv, shortopts, longopts, config):
      self.argv = argv
      self.shortopts = _G_SHORTOPTS + shortopts
      self.longopts = _G_LONGOPTS = longopts
      self.optional_config_files = DEFAULT_CONFIG_FILENAMES 
      self.specific_config_files = []
      self.config = DEFAULT_CONFIG
      self.config.update (config)
      
      opts, args = getopt.getopt (self.argv, self.shortopts, self.longopts)
      self.process_configs (opts, args)

   def process_configs (self, opts, args):
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

   def process_account_settings (self, account):
      """
         Merge settings from an account in the config map.
      """

      if not account in self.config:
         raise Exception ("No settings for account \"%s\" found in the config map." % account)

      self.config.update (self.config [account])

       
   def run (self):
      raise NotImplementedException ("Abstract base method called.")

#-------------------------------------------------------------------
class CheckMailCommand (BaseCommand):
   """
      Command to check the mailbox for new messages
      and print them in a list with uid and subject.
   """

   def __init__ (self, argv):
      SHORTOPTS   = 'qu:p:l:h:i:p'
      LONGOPTS    = ['quiet', 'username=', 'password=', 'limit=', 'host=', 'inbox=', 'port=']

      BaseCommand.__init__ (
            self, argv, SHORTOPTS, LONGOPTS, {
               'quiet':    False,
               'limit':    None
            })

   def process_config (self, opts, args):
      BaseCommand.process_config (self, opts, args)

      for opt, val in opts:
         if opt in ['-l', '--limit']:
            self.config ['limit'] = int (val)

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

      if self.config ['quiet']:
         print ("%d" % client.fetch_unread_count ())

      else:
         print ("%d unread messages." % client.fetch_unread_count ())
         ids = client.fetch_unread_ids ()[self.config['limit']:]
         ids.reverse ()
         
         for id in ids:
            message = client.fetch_message (id)
            sys.stdout.buffer.write (("%s\t%s\n" % (id, message.get_subject ())).encode ())

#-------------------------------------------------------------------
if __name__ == "__main__":
   app = CheckMailCommand (sys.argv[1:])
   app.run ()

