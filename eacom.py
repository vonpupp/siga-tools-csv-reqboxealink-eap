#!/bin/python

#  EaDumpRepository
#
# Dump a Ea Repository (first level only)
#
import win32com.client
import sys

#---Dump functions------------------------------------------------
def DumpRepository(Earepository):
   print "-- Repostitory  --"
   print "Instance GUID: ", Earepository.InstanceGUID
   print "ConnectionString: ", Earepository.ConnectionString
   print "Library version: ", Earepository.LibraryVersion
   print
   print "Model count: ", Earepository.Models.Count
   print "Terms count: ", Earepository.Terms.Count
   print "Issues count: ", Earepository.Issues.Count
   print "Author count: ", Earepository.Authors.Count
   print "Client count: ", Earepository.Clients.Count
   print "Task count: ", Earepository.Tasks.Count
   print "Datatypes count: ", Earepository.Datatypes.Count
   print "Recource count: ", Earepository.Resources.Count
   print "Stereotype count: ", Earepository.Stereotypes.Count
   print "PropertyType count: ", Earepository.PropertyTypes.Count
   print

def DumpAuthors(EaRepository):
   print "-- Authors  --"
   for i in range(EaRepository.Authors.Count):
      x = EaRepository.Authors.GetAt(i)
      print "Author #", i
      print "Name: %s  Roles: %s" % (x.Name, x.Roles)
      print "ObjectType: ", x.ObjectType
      print "Notes: ", x.Notes
      print
   print

def DumpClients(EaRepository):
   print "-- Clients  --"
   for i in range(EaRepository.Clients.Count):
      x = EaRepository.Clients.GetAt(i)
      print "Client #", i
      print "Name: %s  Organization: %s" % (x.Name, x.Organization)
      print "Phone 1: %s  Phone2: %s" % (x.Phone1, x.Phone2)
      print "Mobile: %s  Fax: %s" % (x.Mobile, x.Fax)
      print "Email: %s  Roles: %s"% (x.Email, x.Roles)
      print "Notes: ", x.Notes
      print
   print


def DumpDatatypes(EaRepository):
   print "-- Datatypes  --"
   for i in range(EaRepository.Datatypes.Count):
      x = EaRepository.Datatypes.GetAt(i)
      print "Datatype #", i
      print "Name: %s  Type: %s  Product: %s" % (x.Name, x.Type, x.Product)
      print "Size: %s  GenericType: %s" % (x.Size, x.GenericType)
      print "Size: %s  MaxLen: %s  MaxPrec: %s" % (x.Size, x.Maxlen, x.MaxPrec)
      print "DefaultLen: %s  DefaultPrec: %s  DefaultScale: %s" % (x.Defaultlen, x.DefaultPrec, x.DefaultScale)	 
      print "UserDefined: %s  HasLength: %s" % (x.UserDefined, x.HasLength)
      print
   print

def DumpIssues(EaRepository):
   print "-- Issues  --"
   for i in range(EaRepository.Issues.Count):
      x = EaRepository.Issues.GetAt(i)
      print "Issue #", i
      print "Name: %s  Category: %s" % (x.Name, x.Category)
      print "Date: %s  Owner: %s"% (x.Date, x.Owner)
      print "IssueID: %s  ObjectType: %s" % (x.IssueID, x.ObjectType)
      print "Notes: ", x.Notes
      print
   print

def DumpModels(EaRepository):
   print "-- Models  (a package) --"
   for i in range(EaRepository.Models.Count):
      x = EaRepository.Models.GetAt(i)
      print "Model #", i
      print "Name: %s  PackageID: %s PackageGUID %s" % (x.Name, x.PackageID, x.PackageGUID)
#	   print "Created: %s  Modified: %s  Version: %s" % (x.Created, x.Modified, x.Version)
      print "IsNamespace: %s IsControlled: %s" % (x.IsNamespace, x.IsControlled)
      print "IsProtected %s  IsModel: %s" % (x.IsProtected, x.Ismodel)
      print "UseDTD: %s  LogXML: %s  XMLPath: %s" % (x.UseDTD, x.LogXML, x.XMLPath)
#	   print "LastLoadDate: %s  LastSaveDate: %s" % (x.LastLoadDate, x.LastSaveDate)
      print "Owner: %s  CodePath: %s" % (x.Owner, x.CodePath)
      print "UMLVersion: %s  TreePos: %s"% (x.UMLVersion, x.TreePos)
      print "Element: %s   IsVersionControlled: %s" % (x.Element, x.IsVersionControlled)
      print "BatchLoad: %s  BatchSave: %s" % (x.BatchLoad, x.BatchSave)
      print "Notes: ", x.Notes
      print "Package count: ", x.Packages.Count
      print "Element count: ", x.Elements.Count
      print "Diagram count: ", x.Diagrams.Count
      print "Connector count: ", x.Connectors.Count
      print
   print

def DumpPropertyTypes(EaRepository):
   print "-- PropertyTypes  --"
   for i in range(EaRepository.PropertyTypes.Count):
      x = EaRepository.PropertyTypes.GetAt(i)
      print "PropertyType #", i
      print "Tag: %s  Description: %s" % (x.Tag, x.Description)
      print "Detail: ", x.Detail
      print
   print

def DumpResources(EaRepository):
   print "-- Resources  --"
   for i in range(EaRepository.Resources.Count):
      x = EaRepository.Resources.GetAt(i)
      print "Resource #", i
      print "Name: %s  Organization: %s" % (x.Name, x.Organization)
      print "Phone1: %s  Phone2: %s" % (x.Phone1, x.Phone2)
      print "Mobile: %s  Fax: %s  Email: %s"% (x.Mobile, x.Fax, x.Email)
      print "Roles: %s" % (x.Roles)	 
      print "Notes: ", x.Notes
      print
   print

def DumpStereotypes(EaRepository):
   print "-- Stereotypes  --"
   for i in range(EaRepository.Stereotypes.Count):
      x = EaRepository.Stereotypes.GetAt(i)
      print "Stereotype #", i
#	   print "Name: %s  AppliesTo: %s  Style: %s" % (x.Name, x.AppliesTo, x.Style)
      print "Notes: ", x.Notes
      print
   print

def DumpTasks(EaRepository):
   print "-- Tasks  --"
   for i in range(EaRepository.Tasks.Count):
      x = EaRepository.Tasks.GetAt(i)
      print "Task #", i
      print "Name: %s  Priority: %s" % (x.Name, x.Priority)
      print "Status: %s  Owner: %s" % (x.Status, x.Owner)
      print "StartDate: %s  EndDate: %s" % (x.Startdate, x.EndDate)
      print "Phase: %s  Percent: %s" % (x.Phase, x.Percent)
      print "TotalTime: %s  ActualTime: %s" % (x.TotalTime, x.ActualTime)
      print "AssignedTo: %s  Type: %s" % (x.AssignedTo, x.Type)	 
      print "History: ", x.History
      print "Notes: ", x.Notes
      print
   print

def DumpTerms(EaRepository):
   print "-- Terms  --"
   for i in range(EaRepository.Terms.Count):
      x = EaRepository.Terms.GetAt(i)
      print "Term #", i
      print "Term: %s  Type: %s  TermID: %s" % (x.Type, x.Term, x.TermID)
      print "Meaning: ", x.Meaning
      print
   print
 

#---Dispatch---------------------------------------------------------
try:
   o = win32com.client.Dispatch('EA.Repository')
except:
   print "failure dispatch"
   sys.exit(2)
#---Open file-----------------------------------------------------
try:
   o.OpenFile2('C:\siga-tools-ea-relation\SIGA.EAP',1,0)
   #o.OpenFile2('/cygdrive/c/siga-tools-ea-relation/SIGA.EAP',1,0)
   #o.OpenFile2('SIGA.EAP',1,0)
except:
   print "failure opening file"
   sys.exit(2)
  
#---Start dumping--------------------------------------------------
DumpRepository(o)
DumpModels(o)
DumpTerms(o)
DumpIssues(o)
DumpAuthors(o)
DumpClients(o)
DumpTasks(o)
DumpDatatypes(o)
DumpResources(o)
DumpStereotypes(o)
DumpPropertyTypes(o)

print "End of PythonScript \n\n"

