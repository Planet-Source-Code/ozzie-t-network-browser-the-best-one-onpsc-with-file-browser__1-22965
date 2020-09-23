VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit



Public Enum NetResourceTypes
    Generic = 0
    Domain = 1
    Server = 2
    share = 3
    File = 4
    Group = 5
    Network = 6
    Root = 7
    ShareAdmin = 8
    Directory = 9
    Tree = 10
    NDSContainer = 11
    Printer = &HFF
End Enum

Private mvNetRes As NETRES2
Private mvGotChildren As Boolean
Private mvChildren As NetResources
Private mvAmRoot As Boolean
Private mvAmPrinter As Boolean

Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrcpyA Lib "KERNEL32" Alias "lstrcpy" (ByVal NewString As String, ByVal OldString As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, ByVal lpBuffer As Long, ByRef lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long

Private Type sNETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type
   
Private Type NETRES2 ' VB compatible NETRESOURCE structure
    dwScope As Long  ' Members mapped back to VB datatypes
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3
Private Const RESOURCE_CONTEXT = &H5

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_FIXED = &H0
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GMEM_LOCKCOUNT = &HFF
Private Const GMEM_MODIFY = &H80
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Const ERROR_MORE_DATA = 234

Private Const RESOURCEDISPLAYTYPE_GENERIC = 0
Private Const RESOURCEDISPLAYTYPE_DOMAIN = 1
Private Const RESOURCEDISPLAYTYPE_SERVER = 2
Private Const RESOURCEDISPLAYTYPE_SHARE = 3
Private Const RESOURCEDISPLAYTYPE_FILE = 4
Private Const RESOURCEDISPLAYTYPE_GROUP = 5
Private Const RESOURCEDISPLAYTYPE_NETWORK = 6
Private Const RESOURCEDISPLAYTYPE_ROOT = 7
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN = 8
Private Const RESOURCEDISPLAYTYPE_DIRECTORY = 9
Private Const RESOURCEDISPLAYTYPE_TREE = &HA
Private Const RESOURCEDISPLAYTYPE_NDSCONTAINER = &HB

Private Sub GetPrinters()
' API wrangling...
' Basically the same routine as GetChildren but tweaked to only return printer objects
' It also discards all non-share objects since we only want printers for this enumeration

' Initialise  collection and variables
Dim hEnum As Long, lpBuff As Long
Dim cbBuff As Long, cCount As Long
Dim p As Long, res As Long, i As Long
Dim EnumHTemp As Long
Dim reqBufferSize As Long
Dim nR As sNETRESOURCE  ' API friendly structure
Dim tempRes As NETRES2  ' VB friendly structure
Dim tChild As NetResource

' If this object is the Network root then we need to make a slight adjustment to the starting values
' of our API friendly NETRESOURCE structure
If mvAmRoot Then
    nR.dwUsage = RESOURCEUSAGE_CONNECTABLE
    nR.lpRemoteName = 0
End If

' Open a net enumeration
' Limit enumeration to connectable print resources (i.e. printer objects)
res = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_PRINT, RESOURCEUSAGE_CONNECTABLE, mvNetRes, hEnum)

' Check for errors
If res <> 0 Then
    ' Error returned when trying to open the enumeration
    ' Probably means we don't have access to see its children.
    ' See the MSDN for more details on possible errors.
    ' Currently no trapping is done here and the routine just exits with an empty children collection
    Exit Sub
End If

' Now begin to enumerate the collection
EnumHTemp = hEnum
' Allocate a default buffer for the NETRESOURCE structure returned from the enum routine, say 1K
cbBuff = 1024&
lpBuff = GlobalAlloc(GPTR, cbBuff)
Do
    EnumHTemp = hEnum
    cCount = &HFFFFFFFF ' Number of entries to return from enumeration - &HFFFFFFFF causes all objects to be returned
    res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
    If res = ERROR_MORE_DATA Then
        ' The enumeration has reported that the lpBuff is not big enough to hold all of the information in the
        ' NETRESOURCE structure. cbBuff has been updated to hold the required amount of space.
        GlobalFree lpBuff   ' Free the memory we're using for the current small buffer
        lpBuff = GlobalAlloc(GPTR, cbBuff)  ' Allocate a new space of the size requested by the enumeration
    Else
        If res = 0 Then     ' No error
            p = lpBuff
            ' cCount holds the number of NETRESOURCE structures returned in this pass
            ' (The enumeration returns as many as will fit into the buffer)
            For i = 1 To cCount ' Loop through the buffer, tackling each structure in turn
                CopyMemory nR, ByVal p, LenB(nR)    ' Copy the block of memory representing the structure into a local API friendly NETRESOURCE structure
                p = p + LenB(nR)    ' Step forward in the buffer by the length of the copied structure
                If nR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE Then
                    tempRes.dwDisplayType = nR.dwDisplayType
                    tempRes.dwScope = nR.dwScope
                    tempRes.dwType = nR.dwType
                    tempRes.dwUsage = nR.dwUsage
                    tempRes.lpComment = lStrCpy(nR.lpComment)
                    tempRes.lpLocalName = lStrCpy(nR.lpLocalName)
                    tempRes.lpProvider = lStrCpy(nR.lpProvider)
                    tempRes.lpRemoteName = lStrCpy(nR.lpRemoteName)
                    Set tChild = New NetResource
                    tChild.NRStruct = tempRes
                    tChild.IsPrinter = True ' I know this is a bit of a fudge, but I didn't think it worth the effort to write polymorphic classes for such a small matter
                    mvChildren.Add tChild
                End If
            Next
        End If
    End If
Loop Until cCount = 0
' Close the enum
WNetCloseEnum hEnum
' Free the memory
GlobalFree lpBuff

End Sub

Friend Property Let IsPrinter(pVal As Boolean)
mvAmPrinter = pVal
End Property

Private Function lStrCpy(lStrPointer As Long) As String
Dim TString As String
TString = String(255, Chr$(0))
lstrcpyA TString, lStrPointer
lStrCpy = Left(TString, InStr(TString, Chr$(0)) - 1)
End Function

Public Property Get Children() As NetResources
If Not mvGotChildren Then GetChildren
Set Children = mvChildren
End Property



Public Property Get Comment() As String
Comment = mvNetRes.lpComment
End Property

Private Sub GetChildren()
' API wrangling...

' Initialise my collection and variables
Set mvChildren = New NetResources
Dim hEnum As Long, lpBuff As Long
Dim cbBuff As Long, cCount As Long
Dim p As Long, res As Long, i As Long
Dim EnumHTemp As Long
Dim reqBufferSize As Long
Dim nR As sNETRESOURCE  ' API friendly structure
Dim tempRes As NETRES2  ' VB friendly structure
Dim tChild As NetResource

' If this object is the Network root then we need to make a slight adjustment to the starting values
' of our API friendly NETRESOURCE structure
If mvAmRoot Then
    nR.dwUsage = RESOURCEUSAGE_CONNECTABLE
    nR.lpRemoteName = 0
End If

' Open a net enumeration
res = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_DISK, 0, mvNetRes, hEnum)

' Check for errors
If res <> 0 Then
    ' Error returned when trying to open the enumeration
    ' Probably means we don't have access to see its children.
    ' See the MSDN for more details on possible errors.
    ' Currently no trapping is done here and the routine just exits with an empty children collection
    Exit Sub
End If

' Now begin to enumerate the collection
EnumHTemp = hEnum
' Allocate a default buffer for the NETRESOURCE structure returned from the enum routine, say 1K
cbBuff = 1024&
lpBuff = GlobalAlloc(GPTR, cbBuff)
Do
    EnumHTemp = hEnum
    cCount = &HFFFFFFFF ' Number of entries to return from enumeration - &HFFFFFFFF causes all objects to be returned
    res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
    If res = ERROR_MORE_DATA Then
        ' The enumeration has reported that the lpBuff is not big enough to hold all of the information in the
        ' NETRESOURCE structure. cbBuff has been updated to hold the required amount of space.
        GlobalFree lpBuff   ' Free the memory we're using for the current small buffer
        lpBuff = GlobalAlloc(GPTR, cbBuff)  ' Allocate a new space of the size requested by the enumeration
    Else
        If res = 0 Then     ' No error
            p = lpBuff
            ' cCount holds the number of NETRESOURCE structures returned in this pass
            ' (The enumeration returns as many as will fit into the buffer)
            For i = 1 To cCount ' Loop through the buffer, tackling each structure in turn
                CopyMemory nR, ByVal p, LenB(nR)    ' Copy the block of memory representing the structure into a local API friendly NETRESOURCE structure
                p = p + LenB(nR)    ' Step forward in the buffer by the length of the copied structure
                tempRes.dwDisplayType = nR.dwDisplayType    ' Begin copying the members of the API friendly structure to the VB friendly structure
                tempRes.dwScope = nR.dwScope
                tempRes.dwType = nR.dwType
                tempRes.dwUsage = nR.dwUsage
                tempRes.lpComment = lStrCpy(nR.lpComment)   ' String copies accomplished by using the lStrCpy routine
                tempRes.lpLocalName = lStrCpy(nR.lpLocalName)
                tempRes.lpProvider = lStrCpy(nR.lpProvider)
                tempRes.lpRemoteName = lStrCpy(nR.lpRemoteName)
                Set tChild = New NetResource    ' Create the new NetResource object that will be the new child
                tChild.NRStruct = tempRes   ' Pass the current VB friendly NETRESOURCE structure to tbe force populate method of the NetResource object
                mvChildren.Add tChild   ' Add the new object to my children collection
            Next
        End If
    End If
Loop Until cCount = 0
' Close the enum
WNetCloseEnum hEnum
' Free the memory
GlobalFree lpBuff

' In order to distinguish printers from other shares we need to enumerate them separately
GetPrinters

mvGotChildren = True

End Sub

Public Property Get LocalName() As String
LocalName = mvNetRes.lpLocalName

End Property


Friend Property Let NRStruct(RHS As NETRES2)
' Private force populate routine
' When a NetResource object it defaults to being the network root object
' The only way to change this is to call this routine and pass a VB friendly NETRES2 NETRESOURCE structure
' When this function is called correctly it populates the data for this NetResource and forces it to act as a child rather than
' a network root.
' When compiled as a COM DLL this function will not be visible to the user - it's intended for internal use only
mvNetRes = RHS
mvAmRoot = False
End Property



Public Property Get Provider() As String
Provider = mvNetRes.lpProvider
End Property

Public Property Get RemoteName() As String
RemoteName = mvNetRes.lpRemoteName
End Property


Public Property Get ResourceType() As NetResourceTypes
If Not mvAmPrinter Then ResourceType = mvNetRes.dwDisplayType Else ResourceType = Printer

End Property

Public Property Get ResourceTypeName() As String
' Provides a friendly name for the resource type as an alternative to using the enumerated "ResourceType" property
' This can be used to quicky bind NetResource objects to named images in an imagelist control (for example)
If mvAmPrinter Then
    ResourceTypeName = "Printer"
    Exit Property
End If
Select Case mvNetRes.dwDisplayType
    Case RESOURCEDISPLAYTYPE_GENERIC
        ResourceTypeName = "Generic"
    Case RESOURCEDISPLAYTYPE_DOMAIN
        ResourceTypeName = "Domain"
    Case RESOURCEDISPLAYTYPE_SERVER
        ResourceTypeName = "Server"
    Case RESOURCEDISPLAYTYPE_SHARE
        ResourceTypeName = "Share"
    Case RESOURCEDISPLAYTYPE_FILE
        ResourceTypeName = "File"
    Case RESOURCEDISPLAYTYPE_GROUP
        ResourceTypeName = "Group"
    Case RESOURCEDISPLAYTYPE_NETWORK
        ResourceTypeName = "Network"
    Case RESOURCEDISPLAYTYPE_ROOT
        ResourceTypeName = "Root"
    Case RESOURCEDISPLAYTYPE_SHAREADMIN
        ResourceTypeName = "AdminShare"
    Case RESOURCEDISPLAYTYPE_DIRECTORY
        ResourceTypeName = "Directory"
    Case RESOURCEDISPLAYTYPE_TREE
        ResourceTypeName = "Tree"
    Case RESOURCEDISPLAYTYPE_NDSCONTAINER
        ResourceTypeName = "NDSContainer"
End Select
End Property

Public Property Get ShortName() As String
' Return just the final part of the object's name (rather than a fully qualified path or context)
Dim i As Integer
i = InStrRev(mvNetRes.lpRemoteName, "\")
ShortName = Right(mvNetRes.lpRemoteName, Len(mvNetRes.lpRemoteName) - i)
End Property


Private Sub Class_Initialize()
mvAmRoot = True
End Sub


Private Sub Class_Terminate()
Set mvChildren = Nothing
End Sub


