Attribute VB_Name = "ModBrowse"
'*************************************
'            eBilling System
'             Version 1.0.0
'      Created by Vishnu Sivan
'          Date : 21-Aug-2018
'*************************************
'
'declare global variable and procedure
'it is used to show folder browser
'
'*************************************

Option Explicit

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'>>> Browsing type.
Public Enum BrowseType
    BrowseForFolders = &H1
    BrowseForComputers = &H1000
    BrowseForPrinters = &H2000
    BrowseForEverything = &H4000
End Enum

'>>> Folder Type
Public Enum FolderType
    CSIDL_BITBUCKET = 10
    CSIDL_CONTROLS = 3
    CSIDL_DESKTOP = 0
    CSIDL_DRIVES = 17
    CSIDL_FONTS = 20
    CSIDL_NETHOOD = 18
    CSIDL_NETWORK = 19
    CSIDL_PERSONAL = 5
    CSIDL_PRINTERS = 4
    CSIDL_PROGRAMS = 2
    CSIDL_RECENT = 8
    CSIDL_SENDTO = 9
    CSIDL_STARTMENU = 11
End Enum

Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, ListId As Long) As Long
 

'>>>> main function for broser
 Public Function BrowseFolders(hWndOwner As Long, sMessage As String, Browse As BrowseType, ByVal RootFolder As FolderType) As String
    Dim Nullpos As Integer
    Dim lpIDList As Long
    Dim res As Long
    Dim sPath As String
    Dim BInfo As BrowseInfo
    Dim RootID As Long

    SHGetSpecialFolderLocation hWndOwner, RootFolder, RootID
    BInfo.hWndOwner = hWndOwner
    BInfo.lpszTitle = lstrcat(sMessage, "")
    BInfo.ulFlags = Browse
    If RootID <> 0 Then BInfo.pIDLRoot = RootID
    lpIDList = SHBrowseForFolder(BInfo)
    If lpIDList <> 0 Then
        sPath = String(MAX_PATH, 0)
        res = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        Nullpos = InStr(sPath, vbNullChar)
        If Nullpos <> 0 Then
            sPath = Left(sPath, Nullpos - 1)
        End If
    End If
    BrowseFolders = sPath
End Function
 

