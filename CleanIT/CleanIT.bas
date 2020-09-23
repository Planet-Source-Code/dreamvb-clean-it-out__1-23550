Attribute VB_Name = "CleanIT"
Private Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Function CleanMediaPlayer()
Dim RegPath As String
Dim Icount As Integer
Dim URL_List As String
    On Error Resume Next
    RegPath = "Software\Microsoft\MediaPlayer\Player\RecentURLList"
    For Icount = 0 To 100
        URL_List = getstring(HKEY_CURRENT_USER, RegPath, "URL" & Icount)
        DeleteValue HKEY_CURRENT_USER, RegPath, "URL" & Icount
        If Len(URL_List) = 0 Then Exit For: Icount = 0
    Next
    URL_List = ""
    
End Function
Public Function CleanPaintBrushFileLst()
Dim RegPath As String
Dim Icount As Integer
Dim File_List As String

    On Error Resume Next
    RegPath = "Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List"
    File_List = getstring(HKEY_CURRENT_USER, RegPath, "File1")
    For Icount = 1 To 100
        DeleteValue HKEY_CURRENT_USER, RegPath, "File" & Icount
        If Len(File_List) = 0 Then Exit For: Icount = 0
    Next
    File_List = ""
    
End Function
Public Function CleanWordPadFileLst()
Dim RegPath As String
Dim Icount As Integer
Dim File_List As String

    On Error Resume Next
    RegPath = "Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad\Recent File List"
    File_List = getstring(HKEY_CURRENT_USER, RegPath, "File1")
    For Icount = 1 To 100
        DeleteValue HKEY_CURRENT_USER, RegPath, "File" & Icount
        If Len(File_List) = 0 Then Exit For: Icount = 0
    Next
    File_List = ""
    

    
End Function
Public Function CleanNetScapeCache()
Dim Path As String
    Path = AddSlash("C:\Program Files\Netscape\Users\default\cache")
    WipeFiles Path
    
End Function
Public Function CleanNetScapeUrls()
Dim Path As String
    Path = AddSlash("C:\Program Files\Netscape\Users\default")
    WipeFile Path & "prefs.js"
    WipeFile Path & "netscape.hst"
    
    
End Function
Public Function VBFileMenu()
Dim RegPath As String
Dim Icount As Integer
Dim File_Lst As String
Dim Item As String

    On Error Resume Next
    RegPath = "Software\Microsoft\Visual Basic\6.0\RecentFiles"
        For Icount = 1 To 100
            Item = Icount
            File_Lst = getstring(HKEY_CURRENT_USER, RegPath, Item)
            DeleteValue HKEY_CURRENT_USER, RegPath, Item
        Next
        File_Lst = ""
        
End Function
Public Function CleanWZipFile()
Dim RegPath As String
Dim Icount As Integer
Dim File_List As String

    On Error Resume Next
    RegPath = "Software\Nico Mak Computing\WinZip\filemenu"
    For Icount = 1 To 100
        File_List = getstring(HKEY_CURRENT_USER, RegPath, "filemenu" & Icount)
        DeleteValue HKEY_CURRENT_USER, RegPath, "filemenu" & Icount
        If Len(File_List) = 0 Then Exit For: Icount = 0
    Next
    File_List = ""
    
End Function
Public Function CleanWZipExtract()
Dim RegPath As String
Dim Icount As Integer
Dim ExtFile_List As String

    On Error Resume Next
    RegPath = "Software\Nico Mak Computing\WinZip\extract"
    For Icount = 1 To 100
        ExtFile_List = getstring(HKEY_CURRENT_USER, RegPath, "extract" & Icount)
        DeleteValue HKEY_CURRENT_USER, RegPath, "extract" & Icount
        If Len(ExtFile_List) = 0 Then Exit For: Icount = 0
    Next
    ExtFile_List = ""
    
End Function

Public Function EmptyBin(THwnd As Long)
    SHEmptyRecycleBin THwnd, vbNullString, 0
    SHUpdateRecycleBinIcon

End Function

Public Function CleanIEUrls()
Dim URL_List As String
Dim RegPath As String
Dim Icount As Integer
    On Error Resume Next
    RegPath = "Software\Microsoft\Internet Explorer\TypedURLs"
    For Icount = 1 To 100
        URL_List = getstring(HKEY_CURRENT_USER, RegPath, "url" & Icount)
        DeleteValue HKEY_CURRENT_USER, RegPath, "url" & Icount
        If Len(URL_List) = 0 Then Exit For: Icount = 0
    Next
        
End Function
Public Function CleanRunMnu()
Dim Mnu_List As String
Dim Menu_Item As String
Dim Icount As Integer
Dim RegPath As String
    RegPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    
    On Error Resume Next
    Mnu_List = getstring(HKEY_CURRENT_USER, RegPath, "MRUList")
    For Icount = 1 To Len(Mnu_List)
        Menu_Item = Mid(Mnu_List, Icount, 1)
        DeleteValue HKEY_CURRENT_USER, RegPath, Menu_Item
    Next
    DeleteValue HKEY_CURRENT_USER, RegPath, "MRUList"
    Icount = 0
    Mnu_List = ""
    Menu_Item = ""
    
End Function
Private Function WipeFiles(lzPath As String)
Dim Icount As Long
Dim TChar1 As String
Dim TChar2 As String
Dim T As Long

TFile = FreeFile
    On Error Resume Next
    x = Dir(lzPath, vbHidden)
    Do While x <> ""
        Open lzPath & x For Binary As #TFile
            FLen = LOF(TFile)
            For Icount = 0 To FLen
                    Randomize
                    TChar1 = Chr(Val(Int(160 * Rnd) + 1))
                    TChar2 = Chr(Val(Int(60 * Rnd) + 2))
                    Tchar3 = Chr(13)
                    T = Seek(TFile)
                    Put #TFile, TFile, TChar1
                    Put #TFile, T, TChar2
            Next
        Close #TFile
    Kill lzPath & x
    x = Dir
    Loop
    TChar1 = "": TChar2 = "": FLen = 0: Icount = 0: T = 0
    

End Function

Private Function WipeFile(Filename As String)
Dim Icount As Long
Dim TChar1 As String
Dim TChar2 As String
Dim T As Long

TFile = FreeFile
    On Error Resume Next
        Open Filename For Binary As #TFile
            FLen = LOF(TFile)
            For Icount = 0 To FLen
                    Randomize
                    TChar1 = Chr(Val(Int(160 * Rnd) + 1))
                    TChar2 = Chr(Val(Int(60 * Rnd) + 2))
                    Tchar3 = Chr(13)
                    T = Seek(TFile)
                    Put #TFile, TFile, TChar1
                    Put #TFile, T, TChar2
            Next
        Close #TFile
    Kill Filename
    TChar1 = "": TChar2 = "": FLen = 0: Icount = 0: T = 0
    

End Function

Function CleanTemp()
    WipeFiles GetTempFolder
    
End Function
Public Function CleanDocs()
    SHAddToRecentDocs 0, 0
    
End Function
Function AddSlash(lzPathName As String) As String
    If Right(lzPathName, 1) = "\" Then AddSlash = lzPathName Else AddSlash = lzPathName & "\"
    
End Function

Private Function GetTempFolder() As String
Dim StrTemp As String
    StrTemp = String(255, Chr(0))
    GetTempPath 255, StrTemp
    StrTemp = Left(StrTemp, InStr(StrTemp, Chr(0)) - 1)
    GetTempFolder = StrTemp
    StrTemp = ""
    
End Function
