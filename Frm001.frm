VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMenu 
   Caption         =   "MCunha - Component manager"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6855
   HelpContextID   =   10
   Icon            =   "Frm001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImgTBar 
      Left            =   6120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":046A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":05CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":072A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":0F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":10E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDial 
      Left            =   360
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImgLista 
      Left            =   3120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":1246
            Key             =   "ctl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":13A2
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":193E
            Key             =   "exe"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm001.frx":1A9A
            Key             =   "else"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4485
      WhatsThisHelpID =   210
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8811
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3975
      HelpContextID   =   220
      Left            =   0
      TabIndex        =   1
      Top             =   360
      WhatsThisHelpID =   220
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7011
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImgLista"
      SmallIcons      =   "ImgLista"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   230
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      HelpContextID   =   230
      Style           =   1
      ImageList       =   "ImgTBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Load list"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear list"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save list"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Registry component"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Unregistry component"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete file"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Properties"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuArq 
      Caption         =   "&File"
      HelpContextID   =   240
      Begin VB.Menu MnuCarregar 
         Caption         =   "&Load list..."
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuLimpar 
         Caption         =   "&Clear list"
         Shortcut        =   ^L
      End
      Begin VB.Menu MnuSalvar 
         Caption         =   "&Save list"
         Shortcut        =   ^S
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRegistar 
         Caption         =   "R&egistry component"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuDesregistrar 
         Caption         =   "&Unregistry component"
         Shortcut        =   ^D
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExcluir 
         Caption         =   "&Delete file"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MnuPropriedades 
         Caption         =   "&Properties"
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSair 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu MnuAjuda 
      Caption         =   "&Help"
      HelpContextID   =   250
      Begin VB.Menu MnuAjudaTopicos 
         Caption         =   "&Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuAjudaProcura 
         Caption         =   "&Search help for..."
         Shortcut        =   +{F1}
      End
      Begin VB.Menu MnuAjudaLinha 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAjudaSobre 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Const MF_BYPOSITION = &H400

Public N As Integer
Public A As Long
Public B As Integer
Public J As Integer
Private Ctrls() As ControlInfo
Public i As Long
Public IMAGEM As Integer
Public TOTAL As Integer
Public ITM As ListItem
Public ARQUIVO As String
Public EXTENSAO As String
Public MOSTRA As String

Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Sub MnuAjudaProcura_Click()
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
  If Err Then
    MsgBox Err.Description, 16
  End If
End Sub

Private Sub MnuAjudaTopicos_Click()
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
  If Err Then
    MsgBox Err.Description, 16
  End If
End Sub



Private Sub Lista_DblClick()
If Lista.ListItems.Count > 0 Then MnuPropriedades_Click
End Sub

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MnuArq
End Sub

Private Sub MnuAjudaSobre_Click()
ShellAbout Me.hWnd, "- " & App.Title, "Developed by Mauricio Cunha" & vbCrLf & "http://www.mcunha98.cjb.net", Me.Icon
End Sub


Private Sub Form_Load()
'*** Code added by HelpWriter ***
    SetAppHelp Me.hWnd
'***********************************
Me.Caption = App.Title & " [Version " & App.Major & "." & Format(App.Revision, "00") & "]"

Lista.ColumnHeaders.Clear
Lista.ColumnHeaders.Add , , "Component", 2900
Lista.ColumnHeaders.Add , , "Version", 1700
Lista.ColumnHeaders.Add , , "File", 3000
Lista.ColumnHeaders.Add , , "Last modify", 3500
Lista.ColumnHeaders.Add , , "Library Type", 3600
Lista.ColumnHeaders.Add , , "Size", 1500
Lista.FullRowSelect = True
Lista.LabelEdit = lvwAutomatic
Lista.View = lvwReport

SystemMenu% = GetSystemMenu(hWnd, 0)
Res% = RemoveMenu(SystemMenu%, 6, MF_BYPOSITION)
Res% = RemoveMenu(SystemMenu%, 6, MF_BYPOSITION)

Habilitar False

End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub
Lista.Move 10, TBar.Height + 5, Me.ScaleWidth, Me.ScaleHeight - (TBar.Height + SBar.Height + 5)
End Sub


Private Sub Form_Unload(Cancel As Integer)
'*** Code added by HelpWriter ***
    QuitHelp
'***********************************
End
End Sub


Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With Lista
    If (ColumnHeader.Index - 1) = .SortKey Then
     .SortOrder = (.SortOrder + 1) Mod 2
    Else
     .Sorted = False
     .SortOrder = 0
     .SortKey = ColumnHeader.Index - 1
     .Sorted = True
    End If
End With
End Sub

Private Sub MnuCarregar_Click()
    Screen.MousePointer = 11
    Status "Enumerating controls, please wait..."
    
    EnumControls Ctrls
    TOTAL = UBound(Ctrls)
    
    For i = 0 To TOTAL
     Status "Now loading component " & i & " of " & TOTAL & ", please wait..."
     
     ARQUIVO = Ctrls(i).File
     EXTENSAO = Right(ARQUIVO, 3)
     
     Select Case LCase(EXTENSAO)
      Case "ocx": IMAGEM = 1
      Case "dll": IMAGEM = 2
      Case "exe": IMAGEM = 3
      Case Else: IMAGEM = 4
     End Select
     
        Set ITM = Lista.ListItems.Add(, , Left(Ctrls(i).Description, 50), IMAGEM, IMAGEM)
        With ITM
            .SubItems(1) = ArqVersao(ARQUIVO)
            .SubItems(2) = Ctrls(i).File
            .SubItems(3) = ArqDataM(ARQUIVO)
            .SubItems(4) = Ctrls(i).TYPELIB
            .SubItems(5) = ArqSize(ARQUIVO)
        End With
    Next
 Lista.SortOrder = lvwAscending
 Lista.SortKey = 0
 Lista.Sorted = True
 
 Screen.MousePointer = 0
 Status ""
 Status2 TOTAL & " components"
 
 If TOTAL > 0 Then Habilitar True Else Habilitar False
End Sub

Private Sub MnuDesregistrar_Click()
If Lista.ListItems.Count = 0 Then Exit Sub


For N = 1 To Lista.ListItems.Count
 If Lista.ListItems(N).Selected = True Then
  Lista.SelectedItem = Lista.ListItems(N)
  ARQUIVO = Lista.SelectedItem.SubItems(2)
  If ExistsFile(ARQUIVO) = True Then
    A = Shell(GetSystemDir & "regsvr32 /u " & ARQUIVO, vbNormalNoFocus)
    If A = 0 Then
     Status "Component " & ARQUIVO & " unregistry with sucess..."
    Else
     Status "Dont possible unregistry " & ARQUIVO & " !"
    End If
  End If
 End If
Next N
Status ""

End Sub

Private Sub MnuExcluir_Click()
Dim strPath As String, blnResult As Boolean

For N = 1 To Lista.ListItems.Count
 If Lista.ListItems(N).Selected = True Then
  Lista.SelectedItem = Lista.ListItems(N)
  ARQUIVO = Lista.SelectedItem.SubItems(2)
  If ExistsFile(ARQUIVO) = True Then
    If blnDeleteFilesToRecycleBin(ARQUIVO) Then Lista.SelectedItem.ForeColor = Me.BackColor
  End If
 End If
Next N
MnuCarregar_Click
End Sub

Private Sub MnuLimpar_Click()
Lista.ListItems.Clear
Status ""
Status2 ""
Habilitar False
End Sub

Private Sub MnuPropriedades_Click()
If Lista.ListItems.Count = 0 Then Exit Sub


For N = 1 To Lista.ListItems.Count
 If Lista.ListItems(N).Selected = True Then
  Lista.SelectedItem = Lista.ListItems(N)
  ARQUIVO = Lista.SelectedItem.SubItems(2)
  If ExistsFile(ARQUIVO) = True Then ShowProps Me.hWnd, ARQUIVO
 End If
Next N
End Sub

Private Sub MnuRegistar_Click()
If Lista.ListItems.Count = 0 Then Exit Sub


For N = 1 To Lista.ListItems.Count
 If Lista.ListItems(N).Selected = True Then
  Lista.SelectedItem = Lista.ListItems(N)
  ARQUIVO = Lista.SelectedItem.SubItems(2)
  If ExistsFile(ARQUIVO) = True Then
    A = Shell(GetSystemDir & "regsvr32 " & ARQUIVO, vbNormalNoFocus)
    If A = 0 Then
     Status "Component " & ARQUIVO & " registry with sucess..."
    Else
     Status "Dont possible registry " & ARQUIVO & " !"
    End If
  End If
 End If
Next N
Status ""
End Sub

Private Sub MnuSair_Click()
Unload Me
End Sub

Private Sub MnuSalvar_Click()
On Error Resume Next

If Lista.ListItems.Count = 0 Then Exit Sub

CDial.filename = ""
CDial.InitDir = App.Path
CDial.Filter = "Arquivos de texto (*.txt)|*.txt"
CDial.ShowSave

If CDial.filename = "" Then Exit Sub

J = Lista.ColumnHeaders.Count
Status "Saving files, please wait..."
MOSTRA = ""

Open CDial.filename For Output As FreeFile
 For A = 1 To J
  MOSTRA = MOSTRA & Lista.ColumnHeaders(A).Text & vbTab
 Next A
 Print #1, MOSTRA
  
 B = J - 1
 
 For N = 1 To Lista.ListItems.Count
  MOSTRA = ""
  Lista.SelectedItem = Lista.ListItems(N)
  MOSTRA = Lista.SelectedItem
   For A = 1 To B
     MOSTRA = MOSTRA & vbTab & Lista.SelectedItem.SubItems(A)
   Next A
  Print #1, MOSTRA
 Next N

Close #1

For N = 1 To Lista.ListItems.Count
 Lista.ListItems(N).Selected = False
Next N

Status ""
End Sub

Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
 Case 1: MnuCarregar_Click
 Case 2: MnuLimpar_Click
 Case 3: MnuSalvar_Click
 Case 4: MnuRegistar_Click
 Case 5: MnuDesregistrar_Click
 Case 6: MnuExcluir_Click
 Case 7: MnuPropriedades_Click
 Case 8: MnuSair_Click
End Select
 

End Sub
Sub Habilitar(Optional B1 As Boolean = False)
TBar.Buttons(1).Enabled = Not (B1)

For N = 2 To 7
TBar.Buttons(N).Enabled = B1
Next N

MnuCarregar.Enabled = Not (B1)
MnuLimpar.Enabled = B1
MnuSalvar.Enabled = B1
MnuRegistar.Enabled = B1
MnuDesregistrar.Enabled = B1
MnuExcluir.Enabled = B1
MnuPropriedades.Enabled = B1

End Sub
