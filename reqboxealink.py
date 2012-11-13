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
   eap = None
   repo = None
   eapfile = None
   linksfile = None
   relmatrix = []
   sdict = {}
   ddict = {}
    
   def __init__(self):
      # Public
      # Init structures
      pass

   def loadlinks(self):
      f = codecs.open(self.linksfile, encoding='utf-8', mode='r')
      reader = csv.reader(f, delimiter='\t')
      self.relmatrix = []
      #reader.next()
      for row in reader:
         self.relmatrix += [row]
      
   def sourcens(self):
      result = self.relmatrix[0][0]
      return result

   def destinationns(self):
      result = self.relmatrix[0][1]
      return result

   def getdispatcher(self):
      try:
         # Dispatch
         self.eap = win32com.client.Dispatch('EA.Repository')
      except:
         print("COM dispatch error: EA.Repository")
         sys.exit(2)
        
   def loadeap(self):
#      try:
         # Open file
         self.eap.OpenFile2(self.eapfile, 1, 0)
         #self.eap.OpenFile2('/cygdrive/c/siga-tools-ea-relation/SIGA.EAP', 1, 0)
         #self.eap.OpenFile2('SIGA.EAP', 1, 0)
      #except:
      #   print "failure opening file"
      #   sys.exit(2)
   
   def nspackagefetch(self, nsstring, dictstruc):
      # Example:
      # SIGA stable|Biblioteca REQ..|Requisitos FI|Comum
      
      ns = nsstring.split('|')
      print("  Opening model: %s" % ns[0]);
      package = self.eap.Models.GetByName(ns[0]);
      for i in range(1, len(ns)):
         print("  Opening package: %s" % ns[i]);
         package = package.Packages.GetByName(ns[i])
         
      for i in xrange(0, package.Elements.Count):
         element = package.Elements.GetAt(i)
         name = element.Name
         #name = element.Name.decode('latin1')
         print("    Mapping element: %s" % name.encode('utf-8'));
         alias = element.Alias
         dictstruc[alias] = element
      pass
   
   def linksourcetodestination(self):
      for i in range(1, len(self.relmatrix)):
         salias = self.relmatrix[i][0]
         dalias = self.relmatrix[i][1]
         selement = None
         delement = None
         if salias in self.sdict:
            selement = self.sdict[self.relmatrix[i][0]]
         if dalias in self.ddict:
            delement = self.ddict[self.relmatrix[i][1]]
         if selement is None:
            print("  MISSING: %s (NOT FOUND) -> %s" % (salias, dalias))
         else:
            if delement is None:
               print("  MISSING: %s -> %s (NOT FOUND)" % (salias, dalias))
            else:
               linkname = self.relmatrix[i][2]
               linktype = self.relmatrix[i][3]
               if linkname is 'auto' or linkname is '':
                  linkname = "rel-" + self.relmatrix[i][0] + "-" + self.relmatrix[i][1]
               
               linksubtype = None
               # TODO: Refactor this
               if 'Extends' == linktype:
                  linksubtype = linktype
                  linktype = 'UseCase'
               elif 'Implements' == linktype:
                  pass
               elif 'Includes' == linktype:
                  linksubtype = linktype
                  linktype = 'UseCase'
                  
               print("  Linking: %s -> %s [%s, %s]" % (salias, dalias, linkname, linktype))
               linkconnection = selement.Connectors.AddNew(linkname, linktype)
               linkconnection.SupplierID = delement.ElementID
               #linkconnection.SetSupplierID(delement.ElementID)
               
               # TODO: Refactor this
               if 'Extends' == linksubtype:
                  linkconnection.Subtype = linksubtype
               elif 'Implements' == linksubtype:
                  pass
               elif 'Includes' == linksubtype:
                  linkconnection.Subtype = linksubtype
               
               linkconnection.Update()
               selement.Refresh()

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
         rl.eapfile = arg
      elif opt in ('-l', '--links'):
         rl.linksfile = arg
      
      #rl.parseall = rb.parseall or opt in ('-a', '--export-all')
   
   print("Loading: " + rl.linksfile + " relationship matrix...")
   rl.loadlinks()
   
   print("Loading: COM dispatch EA.Repository")
   rl.getdispatcher()
   
   print("Loading: " + rl.eapfile + " repository file")
   rl.loadeap()
   
   print("Loading: " + rl.sourcens() + " source namespace")
   rl.nspackagefetch(rl.sourcens(), rl.sdict)
   
   print("Loading: " + rl.destinationns() + " destination namespace")
   rl.nspackagefetch(rl.destinationns(), rl.ddict)

   print("Linking...")
   rl.linksourcetodestination()
   
if __name__ == "__main__":
   sys.exit(main(sys.argv[1:]))