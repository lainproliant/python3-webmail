#-----------------------------------------------------------------------------
# webmail.py
#
# A command line webmail tool.
#
# Author: Lee Supe
# Date: March 10th, 2013
#
#-----------------------------------------------------------------------------

import sys
import getopt

from webmail.client import *

#-----------------------------------------------------------------------------
SHORTOPTS = ""
LONGOPTS = []

#-----------------------------------------------------------------------------
def cmd_check (cfg, opts, args):
    

#-----------------------------------------------------------------------------
def die (message):
   """
      Print the given error message and abort with error.
   """
   
   print (message, file = sys.stderr)
   sys.exit (1)
   
#-----------------------------------------------------------------------------
def main (argv):
   opts, args = getopt.getopt (argv, SHORTOPTS, LONGOPTS)

   command = 'check'

   if len (args) > 0:
      command = args.pop (0)

   if not command in COMMANDS:
      die ("Unrecognized command: %s" % command)

   COMMANDS[command] (cfg, opts, args)

#-----------------------------------------------------------------------------
if __name__ == '__main__':
   main (sys.argv[1:])

