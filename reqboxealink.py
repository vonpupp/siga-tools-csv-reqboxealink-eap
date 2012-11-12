#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#   Project:			SIGA
#   Component Name:		reqboxealink
#   Language:			Python 2.7
#
#   License: 			GNU Public License
#       This file is part of the project.
#	This is free software: you can redistribute it and/or modify
#	it under the terms of the GNU General Public License as published by
#	the Free Software Foundation, either version 3 of the License, or
#	(at your option) any later version.
#
#	Distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
#       without even the implied warranty of MERCHANTABILITY or
#       FITNESS FOR A PARTICULAR PURPOSE.
#       See the GNU General Public License for more details.
#       <http://www.gnu.org/licenses/>
#
#   Author:			Albert De La Fuente (www.albertdelafuente.com)
#   E-Mail:			http://www.google.com/recaptcha/mailhide/d?k=01eb_9W_IYJ4Pm_Y9ALRIPug==&c=L15IEH_kstH8WRWfqnRyeW4IDQuZPzNDRB0KCzMTbHQ=
#
#   Description:		This script will import into a EA Com object (EAP file)
#                               the relationships defined on a .csv file
#
#   Limitations:		Error handling is correctly implemented, time constraints
#	The code is not clean and elegant as it should, again, time constraints
#   Database tables used:	None 
#   Thread Safe:	        No
#   Extendable:			No
#   Platform Dependencies:	Linux (openSUSE used)
#   Compiler Options:

"""
   EA Links tool. [see siga-tools-*-reqbox-*]
   
   This program will load relations from a csv file and apply them into a eap file

   Command Line Usage:
      reqboxealink {<option> <argument>}

   Options:
      -h, --help                          Print this help and exit.

   Examples:
      reqboxealink.py --eap test.eap --links links.csv
"""

import getopt
import win32com.client
import sys
import codecs
import csv

class ReqBoxLinker():
   """ Reqbox
   Attributes:
      - ea: Dispatch based COM object
   """
   ea = None
   relmatrix = []
    
   def __init__(self):
      # Public
      # Init structures
      pass
   
   def getdispatcher(self):
      try:
         # Dispatch
         ea = win32com.client.Dispatch('EA.Repository')
      except:
         print "failure dispatch"
         sys.exit(2)
        
   def loadeap(self):
      try:
         # Open file
         o.OpenFile2(self.eapfile, 1, 0)
         #o.OpenFile2('/cygdrive/c/siga-tools-ea-relation/SIGA.EAP', 1, 0)
         #o.OpenFile2('SIGA.EAP', 1, 0)
      except:
         print "failure opening file"
         sys.exit(2)
   
   def loadlinks(self):
      f = codecs.open(self.linksfile, encoding='utf-8', mode='r')
      reader = csv.reader(f, delimiter='\t')
      self.relmatrix = []
      reader.next()
      for row in reader:
         self.relmatrix += [row]

   def loadeapa(self):
      pass

def main(argv):
   try:
      opts, args = getopt.getopt(argv, 'he:l:', ['help', 'eap', 'links'])
   except getopt.GetoptError, msg:
      sys.stderr.write("reqboxealink: error: %s" % msg)
      sys.stderr.write("See 'reqboxealink --help'.\n")
      return 1
#    if len(args) is not 1:
#        sys.stderr.write("Not enough arguments. See 'reqbox --help'.\n")
#        return 1
    
   rl = ReqBoxLinker()
   
   for opt, arg in opts:
      if opt in ('-h', '--help'):
         sys.stdout.write(__doc__)
         return 0
      elif opt in ('-e', '--eap'):
         rl.eafile = arg
      elif opt in ('-l', '--links'):
         rl.linksfile = arg
      
      #rl.parseall = rb.parseall or opt in ('-a', '--export-all')
   
   rl.loadlinks()
   rl.getdispatcher()
   rl.loadeap()
   
if __name__ == "__main__":
   sys.exit(main(sys.argv[1:]))