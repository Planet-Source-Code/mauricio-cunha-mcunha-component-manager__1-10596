Attribute VB_Name = "Module3"
Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" (lpFileOP As SHFILEOPSTRUCT) As Long

Type BrowseInfo
    hWndOwner As Long
    pIDLRoor As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hWnd As Long
  lpVerb As String
  lpFile As String
  lpParameters As String
  lpDirectory As String
  nShow As Long
  hInstApp As Long
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Public Function ShowProperties(filename As String, OwnerhWnd As Long) As Long
Dim SEI As SHELLEXECUTEINFO, R As Long

With SEI
  .cbSize = Len(SEI)
  .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
  .hWnd = OwnerhWnd
  .lpVerb = "properties"
  .lpFile = filename
  .lpParameters = vbNullChar
  .lpDirectory = vbNullChar
  .nShow = 0
  .hInstApp = 0
  .lpIDList = 0
End With
R = ShellExecuteEX(SEI)
ShowProperties = SEI.hInstApp
End Function

Sub ShowProps(hWnd As Long, strPath As String)
Dim R As Long
Dim fname As String
fname = strPath
R = ShowProperties(fname, hWnd)
If R <= 32 Then MsgBox "Ocorreu um erro ao exibir as propriedades do arquivo !", 16
End Sub

Public Function blnDeleteFilesToRecycleBin _
(ParamArray vntFilename() As Variant) As Boolean

On Error GoTo ErrorToRecycleBin

Dim intK As Integer
Dim strFiles As String
Dim udtShellFileOper As SHFILEOPSTRUCT
Dim lngResult As Long

For intK = LBound(vntFilename) To UBound(vntFilename)
strFiles = strFiles & vntFilename(intK) & vbNullChar
Next

strFiles = strFiles & vbNullChar

With udtShellFileOper
.wFunc = &H3
.pFrom = strFiles
.fFlags = &H40
End With

lngResult = SHFileOperation(udtShellFileOper)

blnDeleteFilesToRecycleBin = True
Exit Function

ErrorToRecycleBin:
blnDeleteFilesToRecycleBin = False
Exit Function

End Function


Public Function strChooseFolder(hWndOwner As Long, strPrompt As String) As String
Dim intNull As Integer
Dim lngIDList As Long
Dim lngResult As Long
Dim strPath As String
Dim udtBI As BrowseInfo

With udtBI
.hWndOwner = hWndOwner
.lpszTitle = lstrcat(strPrompt, "")
.ulFlags = BIF_RETURNONLYFSDIRS
End With

lngIDList = SHBrowseForFolder(udtBI)
If lngIDList Then
strPath = String$(MAX_PATH, 0)
lngResult = SHGetPathFromIDList(lngIDList, strPath)
Call CoTaskMemFree(lngIDList)
intNull = InStr(strPath, vbNullChar)
If intNull Then
strPath = Left$(strPath, intNull - 1)
End If
End If

strChooseFolder = strPath

End Function


