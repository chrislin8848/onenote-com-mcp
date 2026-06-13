# -*- coding: utf-8 -*-
# ruff: noqa
#
# VENDORED makepy module — the fully-generated early-bound wrapper for the OneNote 15.0 type
# library ({0EA692EE-BB50-4E3C-AEF0-356D91732725}, ver 1.1). DO NOT hand-edit; it is committed
# verbatim from `gencache` output (generated on the Windows VM, 2026-06-12).
#
# WHY this is vendored instead of regenerated at runtime (frozen-deploy bug, 2026-06-12):
#   `gencache.EnsureModule(...)` defaults to ON-DEMAND generation — the per-method dispids are
#   filled in lazily on the FIRST call (e.g. GetHierarchy), which calls `LoadRegTypeLib(clsid,
#   1, 1)`. On a cold consumer PC (no warm gen_py cache) that lazy load raised
#   TYPE_E_LIBNOTREGISTERED (0x8002801D): OneNote registers its typelib under HKCR\TypeLib\
#   {0EA692EE...}\1.1\0\Win32 ONLY (no Win64 subkey), so a 64-bit process can't resolve it from
#   the registry. The dev VM never hit this because its gen_py cache was permanently warm. By
#   importing THIS fully-baked module, every method's dispid is already present — no type library
#   is loaded at runtime at all, so the bug cannot occur on any employee PC running OneNote 15.0.
#   `Win32ComBackend.app` imports this lazily (keeps the guarded-import invariant) and falls back
#   to runtime regeneration if it is ever absent. See docs/com-api-reference.md.
#
# TO REGENERATE (only if the OneNote COM interface ever changes): on a Windows box with OneNote,
#   python -c "from win32com.client import gencache, selecttlb; t=max((x for x in selecttlb.EnumTlbs() if 'onenote' in x.desc.lower()), key=lambda x:(int(x.major,16),int(x.minor,16))); gencache.EnsureModule(t.clsid,0,int(t.major,16),int(t.minor,16))"
#   then copy %TEMP%\gen_py\<pyver>\0EA692EE-...x0x1x1.py here (CRLF→LF), keeping this header.
#
# Created by makepy.py version 0.5.01
# By python version 3.12.13 (main, Jun  2 2026, 22:47:20) [MSC v.1944 64 bit (AMD64)]
# From type library '{0EA692EE-BB50-4E3C-AEF0-356D91732725}'
# On Fri Jun 12 15:32:23 2026
'Microsoft OneNote 15.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x30c0df0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{0EA692EE-BB50-4E3C-AEF0-356D91732725}')
MajorVersion = 1
MinorVersion = 1
LibraryFlags = 8
LCID = 0x0

class constants:
	cftFolder                     =2          # from enum CreateFileType
	cftNone                       =0          # from enum CreateFileType
	cftNotebook                   =1          # from enum CreateFileType
	cftSection                    =3          # from enum CreateFileType
	dlBottom                      =4          # from enum DockLocation
	dlDefault                     =-1         # from enum DockLocation
	dlLeft                        =1          # from enum DockLocation
	dlNone                        =0          # from enum DockLocation
	dlRight                       =2          # from enum DockLocation
	dlTop                         =3          # from enum DockLocation
	hrAppInModalUI                =-2147213264 # from enum Error
	hrBinaryObjectDoesNotExist    =-2147213297 # from enum Error
	hrConvertFailed               =-2147213267 # from enum Error
	hrCreatingSection             =-2147213310 # from enum Error
	hrDisabledByPolicy            =-2147213284 # from enum Error
	hrFileAlreadyExists           =-2147213286 # from enum Error
	hrFileDoesNotExist            =-2147213306 # from enum Error
	hrFolderDoesNotExist          =-2147213288 # from enum Error
	hrFutureContentLoss           =-2147213278 # from enum Error
	hrGroupDoesNotExist           =-2147213295 # from enum Error
	hrImportLNTThumbnailFailed    =-2147213270 # from enum Error
	hrInsertingFile               =-2147213290 # from enum Error
	hrInsertingHtml               =-2147213303 # from enum Error
	hrInsertingImage              =-2147213305 # from enum Error
	hrInsertingInk                =-2147213304 # from enum Error
	hrInsertingOutlineText        =-2147213299 # from enum Error
	hrInvalidLinkedNoteThumbnail  =-2147213271 # from enum Error
	hrInvalidLinkedNoteUri        =-2147213272 # from enum Error
	hrInvalidName                 =-2147213289 # from enum Error
	hrInvalidQuery                =-2147213287 # from enum Error
	hrInvalidSelection            =-2147213268 # from enum Error
	hrInvalidXML                  =-2147213311 # from enum Error
	hrInvalidXMLSchema            =-2147213280 # from enum Error
	hrLastModifiedDateDidNotMatch =-2147213296 # from enum Error
	hrLegacySection               =-2147213282 # from enum Error
	hrMalformedXML                =-2147213312 # from enum Error
	hrMergeFailed                 =-2147213281 # from enum Error
	hrNavigatingToPage            =-2147213302 # from enum Error
	hrNoActiveSelection           =-2147213293 # from enum Error
	hrNoFriendlyNameForLinkedNote =-2147213273 # from enum Error
	hrNoShortNameForLinkedNote    =-2147213274 # from enum Error
	hrNotYetSynchronized          =-2147213283 # from enum Error
	hrNotebookDoesNotExist        =-2147213291 # from enum Error
	hrObjectDoesNotExist          =-2147213292 # from enum Error
	hrOpeningSection              =-2147213309 # from enum Error
	hrPageDoesNotExist            =-2147213307 # from enum Error
	hrPageDoesNotExistInGroup     =-2147213294 # from enum Error
	hrPageObjectDoesNotExist      =-2147213298 # from enum Error
	hrPageReadOnly                =-2147213300 # from enum Error
	hrPublishFormatUnsupportedForLabels=-2147213263 # from enum Error
	hrRecordingInProgress         =-2147213276 # from enum Error
	hrRecycleBinEditFailed        =-2147213266 # from enum Error
	hrSectionDoesNotExist         =-2147213308 # from enum Error
	hrSectionEncryptedAndLocked   =-2147213285 # from enum Error
	hrSectionReadOnly             =-2147213301 # from enum Error
	hrTimeOut                     =-2147213277 # from enum Error
	hrUnknownLinkedNoteState      =-2147213275 # from enum Error
	hrUnreadDisabledForNotebook   =-2147213269 # from enum Error
	flContacts                    =1          # from enum FilingLocation
	flEMail                       =0          # from enum FilingLocation
	flMeetings                    =3          # from enum FilingLocation
	flPrintOuts                   =5          # from enum FilingLocation
	flTasks                       =2          # from enum FilingLocation
	flWebContent                  =4          # from enum FilingLocation
	fltCurrentPage                =2          # from enum FilingLocationType
	fltCurrentSectionNewPage      =1          # from enum FilingLocationType
	fltNamedPage                  =4          # from enum FilingLocationType
	fltNamedSectionNewPage        =0          # from enum FilingLocationType
	heNone                        =0          # from enum HierarchyElement
	heNotebooks                   =1          # from enum HierarchyElement
	hePages                       =8          # from enum HierarchyElement
	heSectionGroups               =2          # from enum HierarchyElement
	heSections                    =4          # from enum HierarchyElement
	hsChildren                    =1          # from enum HierarchyScope
	hsNotebooks                   =2          # from enum HierarchyScope
	hsPages                       =4          # from enum HierarchyScope
	hsSections                    =3          # from enum HierarchyScope
	hsSelf                        =0          # from enum HierarchyScope
	npsBlankPageNoTitle           =2          # from enum NewPageStyle
	npsBlankPageWithTitle         =1          # from enum NewPageStyle
	npsDefault                    =0          # from enum NewPageStyle
	nfoLocal                      =1          # from enum NotebookFilterOutType
	nfoNetwork                    =2          # from enum NotebookFilterOutType
	nfoNoWacUrl                   =8          # from enum NotebookFilterOutType
	nfoWeb                        =4          # from enum NotebookFilterOutType
	piAll                         =7          # from enum PageInfo
	piBasic                       =0          # from enum PageInfo
	piBinaryData                  =1          # from enum PageInfo
	piBinaryDataFileType          =5          # from enum PageInfo
	piBinaryDataSelection         =3          # from enum PageInfo
	piFileType                    =4          # from enum PageInfo
	piSelection                   =2          # from enum PageInfo
	piSelectionFileType           =6          # from enum PageInfo
	pfEMF                         =6          # from enum PublishFormat
	pfHTML                        =7          # from enum PublishFormat
	pfMHTML                       =2          # from enum PublishFormat
	pfOneNote                     =0          # from enum PublishFormat
	pfOneNote2007                 =8          # from enum PublishFormat
	pfOneNotePackage              =1          # from enum PublishFormat
	pfPDF                         =3          # from enum PublishFormat
	pfWord                        =5          # from enum PublishFormat
	pfXPS                         =4          # from enum PublishFormat
	rrtFiling                     =1          # from enum RecentResultType
	rrtLinks                      =3          # from enum RecentResultType
	rrtNone                       =0          # from enum RecentResultType
	rrtSearch                     =2          # from enum RecentResultType
	slBackUpFolder                =0          # from enum SpecialLocation
	slDefaultNotebookFolder       =2          # from enum SpecialLocation
	slUnfiledNotesSection         =1          # from enum SpecialLocation
	tcsCollapsed                  =1          # from enum TreeCollapsedStateType
	tcsExpanded                   =0          # from enum TreeCollapsedStateType
	xs2007                        =0          # from enum XMLSchema
	xs2010                        =1          # from enum XMLSchema
	xs2013                        =2          # from enum XMLSchema
	xs2026                        =3          # from enum XMLSchema
	xsCurrent                     =2          # from enum XMLSchema
	xsPreRelease                  =3          # from enum XMLSchema

from win32com.client import DispatchBaseClass
class IApplication(DispatchBaseClass):
	'IApplication Interface'
	CLSID = IID('{452AC71A-B655-4967-A208-A4CC39DD7949}')
	coclass_clsid = IID('{DC67E480-C3CB-49F8-8232-60B0C2056C8E}')

	def CloseNotebook(self, bstrNotebookID=defaultNamedNotOptArg, force=False):
		return self._oleobj_.InvokeTypes(1610743813, LCID, 1, (24, 0), ((8, 1), (11, 49)),bstrNotebookID
			, force)

	def CreateNewPage(self, bstrSectionID=defaultNamedNotOptArg, pbstrPageID=pythoncom.Missing, npsNewPageStyle=0):
		return self._ApplyTypes_(1610743812, 1, (24, 0), ((8, 1), (16392, 2), (3, 49)), 'CreateNewPage', None,bstrSectionID
			, pbstrPageID, npsNewPageStyle)

	def DeleteHierarchy(self, bstrObjectID=defaultNamedNotOptArg, dateExpectedLastModified=(1899, 12, 30, 0, 0, 0, 5, 364, 0), deletePermanently=False):
		return self._oleobj_.InvokeTypes(1610743811, LCID, 1, (24, 0), ((8, 1), (7, 49), (11, 49)),bstrObjectID
			, dateExpectedLastModified, deletePermanently)

	def DeletePageContent(self, bstrPageID=defaultNamedNotOptArg, bstrObjectID=defaultNamedNotOptArg, dateExpectedLastModified=(1899, 12, 30, 0, 0, 0, 5, 364, 0), force=False):
		return self._oleobj_.InvokeTypes(1610743818, LCID, 1, (24, 0), ((8, 1), (8, 1), (7, 49), (11, 49)),bstrPageID
			, bstrObjectID, dateExpectedLastModified, force)

	def FindMeta(self, bstrStartNodeID=defaultNamedNotOptArg, bstrSearchStringName=defaultNamedNotOptArg, pbstrHierarchyXmlOut=pythoncom.Missing, fIncludeUnindexedPages=False
			, xsSchema=2):
		return self._ApplyTypes_(1610743825, 1, (24, 0), ((8, 1), (8, 1), (16392, 2), (11, 49), (3, 49)), 'FindMeta', None,bstrStartNodeID
			, bstrSearchStringName, pbstrHierarchyXmlOut, fIncludeUnindexedPages, xsSchema)

	def FindPages(self, bstrStartNodeID=defaultNamedNotOptArg, bstrSearchString=defaultNamedNotOptArg, pbstrHierarchyXmlOut=pythoncom.Missing, fIncludeUnindexedPages=False
			, fDisplay=False, xsSchema=2):
		return self._ApplyTypes_(1610743824, 1, (24, 0), ((8, 1), (8, 1), (16392, 2), (11, 49), (11, 49), (3, 49)), 'FindPages', None,bstrStartNodeID
			, bstrSearchString, pbstrHierarchyXmlOut, fIncludeUnindexedPages, fDisplay, xsSchema
			)

	def GetBinaryPageContent(self, bstrPageID=defaultNamedNotOptArg, bstrCallbackID=defaultNamedNotOptArg, pbstrBinaryObjectB64Out=pythoncom.Missing):
		return self._ApplyTypes_(1610743817, 1, (24, 0), ((8, 1), (8, 1), (16392, 2)), 'GetBinaryPageContent', None,bstrPageID
			, bstrCallbackID, pbstrBinaryObjectB64Out)

	def GetHierarchy(self, bstrStartNodeID=defaultNamedNotOptArg, hsScope=defaultNamedNotOptArg, pbstrHierarchyXmlOut=pythoncom.Missing, xsSchema=2):
		return self._ApplyTypes_(1610743808, 1, (24, 0), ((8, 1), (3, 1), (16392, 2), (3, 49)), 'GetHierarchy', None,bstrStartNodeID
			, hsScope, pbstrHierarchyXmlOut, xsSchema)

	def GetHierarchyParent(self, bstrObjectID=defaultNamedNotOptArg, pbstrParentID=pythoncom.Missing):
		return self._ApplyTypes_(1610743814, 1, (24, 0), ((8, 1), (16392, 2)), 'GetHierarchyParent', None,bstrObjectID
			, pbstrParentID)

	def GetHyperlinkToObject(self, bstrHierarchyID=defaultNamedNotOptArg, bstrPageContentObjectID=defaultNamedNotOptArg, pbstrHyperlinkOut=pythoncom.Missing):
		return self._ApplyTypes_(1610743823, 1, (24, 0), ((8, 1), (8, 1), (16392, 2)), 'GetHyperlinkToObject', None,bstrHierarchyID
			, bstrPageContentObjectID, pbstrHyperlinkOut)

	def GetPageContent(self, bstrPageID=defaultNamedNotOptArg, pbstrPageXmlOut=pythoncom.Missing, pageInfoToExport=0, xsSchema=2):
		return self._ApplyTypes_(1610743815, 1, (24, 0), ((8, 1), (16392, 2), (3, 49), (3, 49)), 'GetPageContent', None,bstrPageID
			, pbstrPageXmlOut, pageInfoToExport, xsSchema)

	def GetSpecialLocation(self, slToGet=defaultNamedNotOptArg, pbstrSpecialLocationPath=pythoncom.Missing):
		return self._ApplyTypes_(1610743826, 1, (24, 0), ((3, 1), (16392, 2)), 'GetSpecialLocation', None,slToGet
			, pbstrSpecialLocationPath)

	def GetWebHyperlinkToObject(self, bstrHierarchyID=defaultNamedNotOptArg, bstrPageContentObjectID=defaultNamedNotOptArg, pbstrHyperlinkOut=pythoncom.Missing):
		return self._ApplyTypes_(1610743836, 1, (24, 0), ((8, 1), (8, 1), (16392, 2)), 'GetWebHyperlinkToObject', None,bstrHierarchyID
			, bstrPageContentObjectID, pbstrHyperlinkOut)

	def MergeFiles(self, bstrBaseFile=defaultNamedNotOptArg, bstrClientFile=defaultNamedNotOptArg, bstrServerFile=defaultNamedNotOptArg, bstrTargetFile=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610743827, LCID, 1, (24, 0), ((8, 1), (8, 1), (8, 1), (8, 1)),bstrBaseFile
			, bstrClientFile, bstrServerFile, bstrTargetFile)

	def MergeSections(self, bstrSectionSourceId=defaultNamedNotOptArg, bstrSectionDestinationId=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610743833, LCID, 1, (24, 0), ((8, 1), (8, 1)),bstrSectionSourceId
			, bstrSectionDestinationId)

	def NavigateTo(self, bstrHierarchyObjectID=defaultNamedNotOptArg, bstrObjectID='', fNewWindow=False):
		return self._ApplyTypes_(1610743819, 1, (24, 32), ((8, 1), (8, 49), (11, 49)), 'NavigateTo', None,bstrHierarchyObjectID
			, bstrObjectID, fNewWindow)

	def NavigateToUrl(self, bstrUrl=defaultNamedNotOptArg, fNewWindow=False):
		return self._oleobj_.InvokeTypes(1610743820, LCID, 1, (24, 0), ((8, 1), (11, 49)),bstrUrl
			, fNewWindow)

	def OpenHierarchy(self, bstrPath=defaultNamedNotOptArg, bstrRelativeToObjectID=defaultNamedNotOptArg, pbstrObjectID=pythoncom.Missing, cftIfNotExist=0):
		return self._ApplyTypes_(1610743810, 1, (24, 0), ((8, 1), (8, 1), (16392, 2), (3, 49)), 'OpenHierarchy', None,bstrPath
			, bstrRelativeToObjectID, pbstrObjectID, cftIfNotExist)

	def OpenPackage(self, bstrPathPackage=defaultNamedNotOptArg, bstrPathDest=defaultNamedNotOptArg, pbstrPathOut=pythoncom.Missing):
		return self._ApplyTypes_(1610743822, 1, (24, 0), ((8, 1), (8, 1), (16392, 2)), 'OpenPackage', None,bstrPathPackage
			, bstrPathDest, pbstrPathOut)

	def Publish(self, bstrHierarchyID=defaultNamedNotOptArg, bstrTargetFilePath=defaultNamedNotOptArg, pfPublishFormat=0, bstrCLSIDofExporter=''):
		return self._ApplyTypes_(1610743821, 1, (24, 32), ((8, 1), (8, 1), (3, 49), (8, 49)), 'Publish', None,bstrHierarchyID
			, bstrTargetFilePath, pfPublishFormat, bstrCLSIDofExporter)

	# Result is of type IQuickFilingDialog
	def QuickFiling(self):
		ret = self._oleobj_.InvokeTypes(1610743828, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'QuickFiling', '{1D12BD3F-89B6-4077-AA2C-C9DC2BCA42F9}')
		return ret

	def SetFilingLocation(self, flToSet=defaultNamedNotOptArg, fltToSet=defaultNamedNotOptArg, bstrFilingSectionID=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610743830, LCID, 1, (24, 0), ((3, 1), (3, 1), (8, 1)),flToSet
			, fltToSet, bstrFilingSectionID)

	def SyncHierarchy(self, bstrHierarchyID=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610743829, LCID, 1, (24, 0), ((8, 1),),bstrHierarchyID
			)

	def UpdateHierarchy(self, bstrChangesXmlIn=defaultNamedNotOptArg, xsSchema=2):
		return self._oleobj_.InvokeTypes(1610743809, LCID, 1, (24, 0), ((8, 1), (3, 49)),bstrChangesXmlIn
			, xsSchema)

	def UpdatePageContent(self, bstrPageChangesXmlIn=defaultNamedNotOptArg, dateExpectedLastModified=(1899, 12, 30, 0, 0, 0, 5, 364, 0), xsSchema=2, force=False):
		return self._oleobj_.InvokeTypes(1610743816, LCID, 1, (24, 0), ((8, 1), (7, 49), (3, 49), (11, 49)),bstrPageChangesXmlIn
			, dateExpectedLastModified, xsSchema, force)

	_prop_map_get_ = {
		"COMAddIns": (104, 2, (9, 0), (), "COMAddIns", None),
		"Dummy1": (102, 2, (11, 0), (), "Dummy1", None),
		"LanguageSettings": (105, 2, (9, 0), (), "LanguageSettings", None),
		# Method 'Windows' returns object of type 'Windows'
		"Windows": (100, 2, (9, 0), (), "Windows", '{6D4B9C3E-CC05-493F-85E2-43D1006DF96A}'),
	}
	_prop_map_put_ = {
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

class IOneNoteEvents:
	'IOneNoteEvents Interface'
	CLSID = CLSID_Sink = IID('{E2E1511D-502D-4BD0-8B3A-8A89A05CDCAE}')
	coclass_clsid = IID('{DC67E480-C3CB-49F8-8232-60B0C2056C8E}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		        1 : "OnNavigate",
		        2 : "OnHierarchyChange",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnNavigate(self):
#	def OnHierarchyChange(self, bstrActivePageID=defaultNamedNotOptArg):


class IQuickFilingDialog(DispatchBaseClass):
	'IQuickFilingUI Interface'
	CLSID = IID('{1D12BD3F-89B6-4077-AA2C-C9DC2BCA42F9}')
	coclass_clsid = None

	def AddButton(self, bstrText=defaultNamedNotOptArg, allowedElements=defaultNamedNotOptArg, allowedReadOnlyElements=defaultNamedNotOptArg, fDefault=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(10, LCID, 1, (24, 0), ((8, 1), (3, 1), (3, 1), (11, 1)),bstrText
			, allowedElements, allowedReadOnlyElements, fDefault)

	def AddInitialEditor(self, initialEditor=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(17, LCID, 1, (24, 0), ((8, 0),),initialEditor
			)

	def ClearInitialEditors(self):
		return self._oleobj_.InvokeTypes(18, LCID, 1, (24, 0), (),)

	def Run(self, piCallback=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(11, LCID, 1, (24, 0), ((9, 1),),piCallback
			)

	def SetRecentResults(self, recentResults=defaultNamedNotOptArg, fShowCurrentSection=defaultNamedNotOptArg, fShowCurrentPage=defaultNamedNotOptArg, fShowUnfiledNotes=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(8, LCID, 1, (24, 0), ((3, 1), (11, 1), (11, 1), (11, 1)),recentResults
			, fShowCurrentSection, fShowCurrentPage, fShowUnfiledNotes)

	def ShowCreateNewNotebook(self):
		return self._oleobj_.InvokeTypes(16, LCID, 1, (24, 0), (),)

	def ShowSharingHyperlink(self):
		return self._oleobj_.InvokeTypes(19, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"CheckboxState": (3, 2, (11, 0), (), "CheckboxState", None),
		"CheckboxText": (2, 2, (8, 0), (), "CheckboxText", None),
		"Description": (1, 2, (8, 0), (), "Description", None),
		"ParentWindowHandle": (6, 2, (21, 0), (), "ParentWindowHandle", None),
		"Position": (7, 2, (36, 0), (), "Position", None),
		"PressedButton": (13, 2, (19, 0), (), "PressedButton", None),
		"SelectedItem": (12, 2, (8, 0), (), "SelectedItem", None),
		"Title": (0, 2, (8, 0), (), "Title", None),
		"TreeDepth": (5, 2, (3, 0), (), "TreeDepth", None),
		"WindowHandle": (4, 2, (21, 0), (), "WindowHandle", None),
	}
	_prop_map_put_ = {
		"CheckboxState": ((3, LCID, 4, 0),()),
		"CheckboxText": ((2, LCID, 4, 0),()),
		"Description": ((1, LCID, 4, 0),()),
		"NotebookFilterOut": ((15, LCID, 4, 0),()),
		"ParentWindowHandle": ((6, LCID, 4, 0),()),
		"Position": ((7, LCID, 4, 0),()),
		"Title": ((0, LCID, 4, 0),()),
		"TreeCollapsedState": ((14, LCID, 4, 0),()),
		"TreeDepth": ((5, LCID, 4, 0),()),
	}
	# Default property for this class is 'Title'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "Title", None))
	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

class IQuickFilingDialogCallback(DispatchBaseClass):
	'IQuickFilingDialogCallback Interface'
	CLSID = IID('{627EA7B4-95B5-4980-84C1-9D20DA4460B1}')
	coclass_clsid = None

	def OnDialogClosed(self, dialog=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610743808, LCID, 1, (24, 0), ((9, 1),),dialog
			)

	_prop_map_get_ = {
	}
	_prop_map_put_ = {
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

class Window(DispatchBaseClass):
	'Window Interface'
	CLSID = IID('{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}')
	coclass_clsid = None

	def NavigateTo(self, bstrHierarchyObjectID=defaultNamedNotOptArg, bstrObjectID=''):
		return self._ApplyTypes_(9, 1, (24, 32), ((8, 1), (8, 49)), 'NavigateTo', None,bstrHierarchyObjectID
			, bstrObjectID)

	def NavigateToUrl(self, bstrUrl=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(16, LCID, 1, (24, 0), ((8, 1),),bstrUrl
			)

	def SetDockedLocation(self, DockLocation=defaultNamedNotOptArg, ptMonitor=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(17, LCID, 1, (24, 0), ((3, 1), (36, 1)),DockLocation
			, ptMonitor)

	_prop_map_get_ = {
		"Active": (11, 2, (11, 0), (), "Active", None),
		# Method 'Application' returns object of type 'IApplication'
		"Application": (14, 2, (9, 0), (), "Application", '{452AC71A-B655-4967-A208-A4CC39DD7949}'),
		"CurrentNotebookId": (4, 2, (8, 0), (), "CurrentNotebookId", None),
		"CurrentPageId": (1, 2, (8, 0), (), "CurrentPageId", None),
		"CurrentSectionGroupId": (3, 2, (8, 0), (), "CurrentSectionGroupId", None),
		"CurrentSectionId": (2, 2, (8, 0), (), "CurrentSectionId", None),
		"DockedLocation": (13, 2, (3, 0), (), "DockedLocation", None),
		"FullPageView": (10, 2, (11, 0), (), "FullPageView", None),
		"SideNote": (15, 2, (11, 0), (), "SideNote", None),
		"WindowHandle": (0, 2, (21, 0), (), "WindowHandle", None),
	}
	_prop_map_put_ = {
		"Active": ((11, LCID, 4, 0),()),
		"DockedLocation": ((13, LCID, 4, 0),()),
		"FullPageView": ((10, LCID, 4, 0),()),
	}
	# Default property for this class is 'WindowHandle'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (21, 0), (), "WindowHandle", None))
	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

class Windows(DispatchBaseClass):
	'List of our Windows Interface'
	CLSID = IID('{6D4B9C3E-CC05-493F-85E2-43D1006DF96A}')
	coclass_clsid = None

	# Result is of type Window
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((19, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Item', '{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}')
		return ret

	_prop_map_get_ = {
		"Count": (1, 2, (19, 0), (), "Count", None),
		# Method 'CurrentWindow' returns object of type 'Window'
		"CurrentWindow": (3, 2, (9, 0), (), "CurrentWindow", '{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}'),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((19, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}')
		return ret

	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,2,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, '{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (19, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __bool__(self):
		return True

from win32com.client import CoClassBaseClass
# This CoClass is known by the name 'OneNote.Application.14'
class Application(CoClassBaseClass): # A CoClass
	# Application Class
	CLSID = IID('{D7FAC39E-7FF1-49AA-98CF-A1DDD316337E}')
	coclass_sources = [
		IOneNoteEvents,
	]
	default_source = IOneNoteEvents
	coclass_interfaces = [
		IApplication,
	]
	default_interface = IApplication

# This CoClass is known by the name 'OneNote.Application.15'
class Application2(CoClassBaseClass): # A CoClass
	# Application2 Class
	CLSID = IID('{DC67E480-C3CB-49F8-8232-60B0C2056C8E}')
	coclass_sources = [
		IOneNoteEvents,
	]
	default_source = IOneNoteEvents
	coclass_interfaces = [
		IApplication,
	]
	default_interface = IApplication

IApplication_vtables_dispatch_ = 1
IApplication_vtables_ = [
	(( 'GetHierarchy' , 'bstrStartNodeID' , 'hsScope' , 'pbstrHierarchyXmlOut' , 'xsSchema' , 
			 ), 1610743808, (1610743808, (), [ (8, 1, None, None) , (3, 1, None, None) , (16392, 2, None, None) , (3, 49, '2', None) , ], 1 , 1 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'UpdateHierarchy' , 'bstrChangesXmlIn' , 'xsSchema' , ), 1610743809, (1610743809, (), [ (8, 1, None, None) , 
			 (3, 49, '2', None) , ], 1 , 1 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'OpenHierarchy' , 'bstrPath' , 'bstrRelativeToObjectID' , 'pbstrObjectID' , 'cftIfNotExist' , 
			 ), 1610743810, (1610743810, (), [ (8, 1, None, None) , (8, 1, None, None) , (16392, 2, None, None) , (3, 49, '0', None) , ], 1 , 1 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'DeleteHierarchy' , 'bstrObjectID' , 'dateExpectedLastModified' , 'deletePermanently' , ), 1610743811, (1610743811, (), [ 
			 (8, 1, None, None) , (7, 49, '(1899, 12, 30, 0, 0, 0, 5, 364, 0)', None) , (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'CreateNewPage' , 'bstrSectionID' , 'pbstrPageID' , 'npsNewPageStyle' , ), 1610743812, (1610743812, (), [ 
			 (8, 1, None, None) , (16392, 2, None, None) , (3, 49, '0', None) , ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'CloseNotebook' , 'bstrNotebookID' , 'force' , ), 1610743813, (1610743813, (), [ (8, 1, None, None) , 
			 (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'GetHierarchyParent' , 'bstrObjectID' , 'pbstrParentID' , ), 1610743814, (1610743814, (), [ (8, 1, None, None) , 
			 (16392, 2, None, None) , ], 1 , 1 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'GetPageContent' , 'bstrPageID' , 'pbstrPageXmlOut' , 'pageInfoToExport' , 'xsSchema' , 
			 ), 1610743815, (1610743815, (), [ (8, 1, None, None) , (16392, 2, None, None) , (3, 49, '0', None) , (3, 49, '2', None) , ], 1 , 1 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'UpdatePageContent' , 'bstrPageChangesXmlIn' , 'dateExpectedLastModified' , 'xsSchema' , 'force' , 
			 ), 1610743816, (1610743816, (), [ (8, 1, None, None) , (7, 49, '(1899, 12, 30, 0, 0, 0, 5, 364, 0)', None) , (3, 49, '2', None) , (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'GetBinaryPageContent' , 'bstrPageID' , 'bstrCallbackID' , 'pbstrBinaryObjectB64Out' , ), 1610743817, (1610743817, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16392, 2, None, None) , ], 1 , 1 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'DeletePageContent' , 'bstrPageID' , 'bstrObjectID' , 'dateExpectedLastModified' , 'force' , 
			 ), 1610743818, (1610743818, (), [ (8, 1, None, None) , (8, 1, None, None) , (7, 49, '(1899, 12, 30, 0, 0, 0, 5, 364, 0)', None) , (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'NavigateTo' , 'bstrHierarchyObjectID' , 'bstrObjectID' , 'fNewWindow' , ), 1610743819, (1610743819, (), [ 
			 (8, 1, None, None) , (8, 49, "''", None) , (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 144 , (3, 32, None, None) , 0 , )),
	(( 'NavigateToUrl' , 'bstrUrl' , 'fNewWindow' , ), 1610743820, (1610743820, (), [ (8, 1, None, None) , 
			 (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'Publish' , 'bstrHierarchyID' , 'bstrTargetFilePath' , 'pfPublishFormat' , 'bstrCLSIDofExporter' , 
			 ), 1610743821, (1610743821, (), [ (8, 1, None, None) , (8, 1, None, None) , (3, 49, '0', None) , (8, 49, "''", None) , ], 1 , 1 , 4 , 0 , 160 , (3, 32, None, None) , 0 , )),
	(( 'OpenPackage' , 'bstrPathPackage' , 'bstrPathDest' , 'pbstrPathOut' , ), 1610743822, (1610743822, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16392, 2, None, None) , ], 1 , 1 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'GetHyperlinkToObject' , 'bstrHierarchyID' , 'bstrPageContentObjectID' , 'pbstrHyperlinkOut' , ), 1610743823, (1610743823, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16392, 2, None, None) , ], 1 , 1 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'FindPages' , 'bstrStartNodeID' , 'bstrSearchString' , 'pbstrHierarchyXmlOut' , 'fIncludeUnindexedPages' , 
			 'fDisplay' , 'xsSchema' , ), 1610743824, (1610743824, (), [ (8, 1, None, None) , (8, 1, None, None) , 
			 (16392, 2, None, None) , (11, 49, 'False', None) , (11, 49, 'False', None) , (3, 49, '2', None) , ], 1 , 1 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'FindMeta' , 'bstrStartNodeID' , 'bstrSearchStringName' , 'pbstrHierarchyXmlOut' , 'fIncludeUnindexedPages' , 
			 'xsSchema' , ), 1610743825, (1610743825, (), [ (8, 1, None, None) , (8, 1, None, None) , (16392, 2, None, None) , 
			 (11, 49, 'False', None) , (3, 49, '2', None) , ], 1 , 1 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'GetSpecialLocation' , 'slToGet' , 'pbstrSpecialLocationPath' , ), 1610743826, (1610743826, (), [ (3, 1, None, None) , 
			 (16392, 2, None, None) , ], 1 , 1 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'MergeFiles' , 'bstrBaseFile' , 'bstrClientFile' , 'bstrServerFile' , 'bstrTargetFile' , 
			 ), 1610743827, (1610743827, (), [ (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'QuickFiling' , 'ppiDialog' , ), 1610743828, (1610743828, (), [ (16393, 10, None, "IID('{1D12BD3F-89B6-4077-AA2C-C9DC2BCA42F9}')") , ], 1 , 1 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'SyncHierarchy' , 'bstrHierarchyID' , ), 1610743829, (1610743829, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'SetFilingLocation' , 'flToSet' , 'fltToSet' , 'bstrFilingSectionID' , ), 1610743830, (1610743830, (), [ 
			 (3, 1, None, None) , (3, 1, None, None) , (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'Windows' , 'ppONWindows' , ), 100, (100, (), [ (16393, 10, None, "IID('{6D4B9C3E-CC05-493F-85E2-43D1006DF96A}')") , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Dummy1' , 'pBool' , ), 102, (102, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 64 , )),
	(( 'MergeSections' , 'bstrSectionSourceId' , 'bstrSectionDestinationId' , ), 1610743833, (1610743833, (), [ (8, 1, None, None) , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'COMAddIns' , 'ppiComAddins' , ), 104, (104, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'LanguageSettings' , 'ppiLanguageSettings' , ), 105, (105, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'GetWebHyperlinkToObject' , 'bstrHierarchyID' , 'bstrPageContentObjectID' , 'pbstrHyperlinkOut' , ), 1610743836, (1610743836, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16392, 2, None, None) , ], 1 , 1 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
]

IQuickFilingDialog_vtables_dispatch_ = 1
IQuickFilingDialog_vtables_ = [
	(( 'Title' , 'bstrTitle' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Title' , 'bstrTitle' , ), 0, (0, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Description' , 'bstrDescription' , ), 1, (1, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Description' , 'bstrDescription' , ), 1, (1, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'CheckboxText' , 'bstrText' , ), 2, (2, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'CheckboxText' , 'bstrText' , ), 2, (2, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'CheckboxState' , 'pfChecked' , ), 3, (3, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'CheckboxState' , 'pfChecked' , ), 3, (3, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'WindowHandle' , 'pHWNDWindow' , ), 4, (4, (), [ (16405, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'TreeDepth' , 'pTreeDepth' , ), 5, (5, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'TreeDepth' , 'pTreeDepth' , ), 5, (5, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'ParentWindowHandle' , 'pHWNDParentWindow' , ), 6, (6, (), [ (16405, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'ParentWindowHandle' , 'pHWNDParentWindow' , ), 6, (6, (), [ (21, 1, None, None) , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'Position' , 'pPoint' , ), 7, (7, (), [ (16420, 10, None, "IID('{00000000-0000-0000-0000-000000000000}')") , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Position' , 'pPoint' , ), 7, (7, (), [ (36, 1, None, "IID('{00000000-0000-0000-0000-000000000000}')") , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'SetRecentResults' , 'recentResults' , 'fShowCurrentSection' , 'fShowCurrentPage' , 'fShowUnfiledNotes' , 
			 ), 8, (8, (), [ (3, 1, None, None) , (11, 1, None, None) , (11, 1, None, None) , (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'AddButton' , 'bstrText' , 'allowedElements' , 'allowedReadOnlyElements' , 'fDefault' , 
			 ), 10, (10, (), [ (8, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'Run' , 'piCallback' , ), 11, (11, (), [ (9, 1, None, "IID('{627EA7B4-95B5-4980-84C1-9D20DA4460B1}')") , ], 1 , 1 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'SelectedItem' , 'pbstrSelectedNodeID' , ), 12, (12, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'PressedButton' , 'pButtonIndex' , ), 13, (13, (), [ (16403, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'TreeCollapsedState' , ), 14, (14, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'NotebookFilterOut' , ), 15, (15, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'ShowCreateNewNotebook' , ), 16, (16, (), [ ], 1 , 1 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'AddInitialEditor' , 'initialEditor' , ), 17, (17, (), [ (8, 0, None, None) , ], 1 , 1 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'ClearInitialEditors' , ), 18, (18, (), [ ], 1 , 1 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'ShowSharingHyperlink' , ), 19, (19, (), [ ], 1 , 1 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
]

IQuickFilingDialogCallback_vtables_dispatch_ = 1
IQuickFilingDialogCallback_vtables_ = [
	(( 'OnDialogClosed' , 'dialog' , ), 1610743808, (1610743808, (), [ (9, 1, None, "IID('{1D12BD3F-89B6-4077-AA2C-C9DC2BCA42F9}')") , ], 1 , 1 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
]

Window_vtables_dispatch_ = 1
Window_vtables_ = [
	(( 'WindowHandle' , 'pHWNDWindow' , ), 0, (0, (), [ (16405, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'CurrentPageId' , 'pbstrPageObjectId' , ), 1, (1, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'CurrentSectionId' , 'pbstrSectionObjectId' , ), 2, (2, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'CurrentSectionGroupId' , 'pbstrSectionObjectId' , ), 3, (3, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'CurrentNotebookId' , 'pbstrNotebookObjectId' , ), 4, (4, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'NavigateTo' , 'bstrHierarchyObjectID' , 'bstrObjectID' , ), 9, (9, (), [ (8, 1, None, None) , 
			 (8, 49, "''", None) , ], 1 , 1 , 4 , 0 , 96 , (3, 32, None, None) , 0 , )),
	(( 'FullPageView' , 'pIsFullPageView' , ), 10, (10, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'FullPageView' , 'pIsFullPageView' , ), 10, (10, (), [ (11, 0, None, None) , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Active' , 'pIsActive' , ), 11, (11, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Active' , 'pIsActive' , ), 11, (11, (), [ (11, 0, None, None) , ], 1 , 4 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'DockedLocation' , 'pDockLocation' , ), 13, (13, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'DockedLocation' , 'pDockLocation' , ), 13, (13, (), [ (3, 0, None, None) , ], 1 , 4 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'ppiApp' , ), 14, (14, (), [ (16393, 10, None, "IID('{452AC71A-B655-4967-A208-A4CC39DD7949}')") , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'SideNote' , 'pIsSideNote' , ), 15, (15, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'NavigateToUrl' , 'bstrUrl' , ), 16, (16, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'SetDockedLocation' , 'DockLocation' , 'ptMonitor' , ), 17, (17, (), [ (3, 1, None, None) , 
			 (36, 1, None, "IID('{00000000-0000-0000-0000-000000000000}')") , ], 1 , 1 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
]

Windows_vtables_dispatch_ = 1
Windows_vtables_ = [
	(( 'Item' , 'Index' , 'Item' , ), 0, (0, (), [ (19, 1, None, None) , 
			 (16393, 10, None, "IID('{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Count' , 'Count' , ), 1, (1, (), [ (16403, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( '_NewEnum' , '_NewEnum' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 1024 , )),
	(( 'CurrentWindow' , 'ppCurrentWindow' , ), 3, (3, (), [ (16393, 10, None, "IID('{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}')") , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
]

RecordMap = {
	###'tagPOINT': '{00000000-0000-0000-0000-000000000000}', # Record disabled because it doesn't have a non-null GUID
}

CLSIDToClassMap = {
	'{1D12BD3F-89B6-4077-AA2C-C9DC2BCA42F9}' : IQuickFilingDialog,
	'{627EA7B4-95B5-4980-84C1-9D20DA4460B1}' : IQuickFilingDialogCallback,
	'{452AC71A-B655-4967-A208-A4CC39DD7949}' : IApplication,
	'{6D4B9C3E-CC05-493F-85E2-43D1006DF96A}' : Windows,
	'{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}' : Window,
	'{E2E1511D-502D-4BD0-8B3A-8A89A05CDCAE}' : IOneNoteEvents,
	'{D7FAC39E-7FF1-49AA-98CF-A1DDD316337E}' : Application,
	'{DC67E480-C3CB-49F8-8232-60B0C2056C8E}' : Application2,
}
CLSIDToPackageMap = {}
win32com.client.CLSIDToClass.RegisterCLSIDsFromDict( CLSIDToClassMap )
VTablesToPackageMap = {}
VTablesToClassMap = {
	'{1D12BD3F-89B6-4077-AA2C-C9DC2BCA42F9}' : 'IQuickFilingDialog',
	'{627EA7B4-95B5-4980-84C1-9D20DA4460B1}' : 'IQuickFilingDialogCallback',
	'{452AC71A-B655-4967-A208-A4CC39DD7949}' : 'IApplication',
	'{6D4B9C3E-CC05-493F-85E2-43D1006DF96A}' : 'Windows',
	'{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}' : 'Window',
}


NamesToIIDMap = {
	'IQuickFilingDialog' : '{1D12BD3F-89B6-4077-AA2C-C9DC2BCA42F9}',
	'IQuickFilingDialogCallback' : '{627EA7B4-95B5-4980-84C1-9D20DA4460B1}',
	'IApplication' : '{452AC71A-B655-4967-A208-A4CC39DD7949}',
	'Windows' : '{6D4B9C3E-CC05-493F-85E2-43D1006DF96A}',
	'Window' : '{8E8304B8-CBD1-44F8-B0E8-89C625B2002E}',
	'IOneNoteEvents' : '{E2E1511D-502D-4BD0-8B3A-8A89A05CDCAE}',
}

win32com.client.constants.__dicts__.append(constants.__dict__)

