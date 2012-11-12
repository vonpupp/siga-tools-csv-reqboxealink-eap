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
   EA Links tool. [see siga-tools-ea]
   
   This program will load from a doc file an implied hierarchy or relations and will produce several ouputs

   Command Line Usage:
      reqbox {<option> <argument>}

   Options:
      -h, --help                          Print this help and exit.
      -a, --export-all                    Parse all
      --parse-v1 / --parse-v2             Prints items in that level

   Examples:
       reqbox.py -a --parse-v2 ./data/LFv14.ms.default.fixed.txt

   Note:
      Please note that the csv files dir is hardcoded into the application
      to make command line arguments easier.
       
       The csv files are:
         - in-rfn-objects.csv
         - in-rgn-objects.csv
         - in-rnf-objects.csv
"""
import win32com.client
import sys

class ReqBoxLinker():
   """ Reqbox
   Attributes:
      - ea: Dispatch based COM object
   """
   ea = None
    
   def __init__(self):
      # Public
      # Init structures
      try:
         # Dispatch
         ea = win32com.client.Dispatch('EA.Repository')
      except:
         print "failure dispatch"
         sys.exit(2)
        
   def open(self, filename):
      try:
         # Open file
         o.OpenFile2(filename,1,0)
         #o.OpenFile2('/cygdrive/c/siga-tools-ea-relation/SIGA.EAP',1,0)
         #o.OpenFile2('SIGA.EAP',1,0)
      except:
         print "failure opening file"
         sys.exit(2)
   
   def loadlinks(self, filename):
      pass

   def loadlinks(self, filename):
      pass

def main(argv):
   try:
      optlist, args = getopt.getopt(argv[1:], 'hv:e:l:', ['help', 'verbose',
         'eap', 'links'])
   except getopt.GetoptError, msg:
      sys.stderr.write("reqboxealink: error: %s" % msg)
      sys.stderr.write("See 'reqboxealink --help'.\n")
      return 1
#    if len(args) is not 1:
#        sys.stderr.write("Not enough arguments. See 'reqbox --help'.\n")
#        return 1
    
   rl = ReqBoxLinker()
   
   for opt, optarg in optlist:
      if opt in ('-h', '--help'):
         sys.stdout.write(__doc__)
         return 0
      elif opt in ('-v', '--verbose'):
         pass
      
      rb.eafile = args[0]
      rb.linksfile = args[1]
      rb.parseall = rb.parseall or opt in ('-a', '--export-all')
      rb.parsefun = rb.parseall or rb.parsefun or opt in ('-f', '--export-fun')
      rb.parserfi = rb.parseall or rb.parserfi or opt in ('-i', '--export-rfi')
      rb.parserfn = rb.parseall or rb.parserfn or opt in ('-r', '--export-rfn')
      rb.parsernf = rb.parseall or rb.parsernf or opt in ('-n', '--export-rnf')
      rb.parsergn = rb.parseall or rb.parsergn or opt in ('-g', '--export-rgn')
      rb.parseext = rb.parseall or rb.parseext or opt in ('-e', '--export-ext')
      rb.parseinc = rb.parseall or rb.parseinc or opt in ('-n', '--export-inc')
      rb.parseimp = rb.parseall or rb.parseimp or opt in ('-m', '--export-imp')
      rb.inobjects = opt in ('-o', '--in-objects')
   
   rl.loadlinks()
   rl.open('C:\siga-tools-ea-relation\SIGA.EAP')