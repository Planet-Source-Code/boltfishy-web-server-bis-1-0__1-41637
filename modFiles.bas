Attribute VB_Name = "modFiles"
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Public Root As String
Public Index As String
Public Port As Integer
Public Minimized As Integer
Public AutoStart As Integer
Public SaveLog As Integer
Public ErrorPath As String

Public Function FileExists(FilePath As String) As Boolean

    'find out if a file exists
    FileExists = Dir(FilePath) <> ""
    
End Function

Public Function GetFileName(FilePath As String) As String
    'return file name from a path
    
    Dim i As Integer
    On Error Resume Next

    For i = Len(FilePath) To 1 Step -1 'i to length of file going back
        If Mid(FilePath, i, 1) = "\" Then 'when it finds the \
            Exit For 'stop trying
        End If
    Next
     
    GetFileName = Mid(FilePath, i + 1)

End Function

Public Function GetPath(ByVal FilePath As String, Optional ByVal AddSlash As Boolean = False) As String
    'Retrieve path from a filepath
    
    Dim temp As String
    Dim i, X As Integer
    
    For X = 0 To Len(FilePath) - 1
        temp = temp & Mid(FilePath, Len(FilePath) - X, 1)
    Next X
    
    i = InStr(1, temp, "\")
    GetPath = Left(FilePath, Len(FilePath) - i + 1)
    
End Function
