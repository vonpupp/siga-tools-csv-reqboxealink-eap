'''
List all IDs and associated package/diagram/element
Usage:
[shighlight]python[/shighlight] eaid.py full_path_to_eap
'''

def expand_diags(diagrams, indent):
   for i in xrange(0, diagrams.Count):
      diagram = diagrams.GetAt(i)
      print ("  " * indent) + "Diagram ID: " + `diagram.DiagramID` + "\t" + diagram.Name

def expand_conns(connectors, indent):
   for i in xrange(0, connectors.Count):
      connector = connectors.GetAt(i)
      print ("  " * indent) + "Connector ID: " + `connector.ConnectorID` + "\t" + connector.Name + "\t" + connector.Type + "\t" + connector.MetaType + "\t" + connector.Subtype
      
def expand_elems(elements, indent):
   for i in xrange(0, elements.Count):
      element = elements.GetAt(i)
      print ("  " * indent) + "Element ID: " + `element.ElementID` + "\t" + element.Name
      connectors = element.Connectors
      if connectors != None:
         expand_conns(connectors, indent+1)
      
def expand_pkgs(packages, indent):
   for i in xrange(0, packages.Count):
      package = packages.GetAt(i)
      if 'p' in opts:
         print ("  " * indent) + "Package ID: " + `package.PackageID` + "\t" + package.Name
      if 'd' in opts:
         expand_diags(package.Diagrams, indent + 1)
      if 'e' in opts:
         expand_elems(package.Elements, indent + 1)
      innerpkg = package.Packages
      if innerpkg.Count > 0:
         expand_pkgs(innerpkg, indent + 1)


opts = ['p', 'd', 'e']

import sys
filepath = sys.argv[1]

import win32com.client
ea = win32com.client.Dispatch('EA.Repository')

if ea == None:
   print "COM dispatch error: EA.Repository"
else:
   #if(ea.OpenFile(filepath)): #raises pywintypes.com_error
   ea.OpenFile2('C:\siga-tools-ea-relation\LIST.EAP',1,0)
   #if(ea.OpenFile2(filepath, "", "")):
   expand_pkgs(ea.Models, 0)
   #else:
      #sys.exit(ea.GetLastError()) 

