Attribute VB_Name = "Module2"
Global Const MAX_PATH = 260
Global Const MAX_LENGTH = 260
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nsize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nsize As Long) As Long
Global FINFO As CFileInfo
Global FINFV As CFileVersionInfo
Global STRTEMP As String
Global LRET As Long

Function ArqDataM(ARQUIVO As String) As String
Set FINFO = New CFileInfo
FINFO.FullPathName = ARQUIVO
ArqDataM = FINFO.FormatFileDate(FINFO.ModifyTime)
End Function

Function ArqVersao(ARQUIVO As String) As String
ArqVersao = ""
Set FINFV = New CFileVersionInfo
FINFV.FullPathName = ARQUIVO
If FINFV.Available Then ArqVersao = FINFV.FileVersion
End Function

Function ArqSize(ARQUIVO As String) As String
ArqSize = 0
Set FINFO = New CFileInfo
FINFO.FullPathName = ARQUIVO
ArqSize = FINFO.FormatFileSize(FINFO.FileSize)
End Function

Sub Status(Texto As String)
FrmMenu.SBar.Panels(1).Text = Texto
End Sub
Sub Status2(Texto As String)
FrmMenu.SBar.Panels(2).Text = Texto
End Sub
Function GetWindowsDir() As Variant
STRTEMP = Space$(MAX_LENGTH)
LRET = GetWindowsDirectory(STRTEMP, MAX_LENGTH)
LRET = InStr(STRTEMP, Chr$(0))
GetWindowsDir = FixPath(Left$(STRTEMP, LRET - 1))
End Function
Function GetSystemDir() As Variant
STRTEMP = Space$(MAX_LENGTH)
LRET = GetSystemDirectory(STRTEMP, MAX_LENGTH)
LRET = InStr(STRTEMP, Chr$(0))
GetSystemDir = FixPath(Left(STRTEMP, LRET - 1))


End Function
Private Function FixPath(ByVal PassedPath As String) As String
  If Right$(PassedPath, 1) = "\" Then
    FixPath = PassedPath
  Else
    FixPath = PassedPath & "\"
  End If
End Function

