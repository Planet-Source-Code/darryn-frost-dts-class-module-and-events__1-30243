VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   4950
      TabIndex        =   5
      Top             =   1770
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "Execute Package"
         Height          =   375
         Left            =   300
         TabIndex        =   7
         Top             =   450
         Width           =   3525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Execute the Included Sample DTS Class Module"
         Height          =   285
         Left            =   300
         TabIndex        =   6
         Top             =   210
         Width           =   3645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   180
      TabIndex        =   3
      Top             =   1800
      Width           =   4185
      Begin VB.CommandButton Command1 
         Caption         =   "Get Package From Server and Create Class Module"
         Height          =   495
         Left            =   180
         TabIndex        =   4
         Top             =   330
         Width           =   3855
      End
   End
   Begin MSComctlLib.ProgressBar prgDTS 
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3150
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgDest 
      Left            =   6600
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgSource 
      Left            =   6570
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgPkg 
      Left            =   6330
      Top             =   2010
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   $"Form1.frx":0000
      Height          =   1545
      Left            =   180
      TabIndex        =   8
      Top             =   120
      Width           =   8985
   End
   Begin VB.Label lblCurrentSet 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   3030
      Width           =   4185
   End
   Begin VB.Label lblRowCount 
      BackStyle       =   0  'Transparent
      Caption         =   "lblRowCount"
      Height          =   345
      Left            =   210
      TabIndex        =   1
      Top             =   3300
      Width           =   4155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents testDTSClass As ClassDTSScript
Attribute testDTSClass.VB_VarHelpID = -1

Private Sub command3_Click()
  Dim strSourceConString As String
  Dim strDestConString As String
  Dim strfilename As String
  Dim strdestfile As String
  Dim blnRet As Boolean
  
With dlgSource
  .InitDir = App.Path & "\Database"
  .DialogTitle = "Find Endpoint File"
  .Filter = "Access Databases (*.mdb)|*.mdb"
  .ShowOpen
End With

If Len(dlgSource.FileName) = 0 Then Exit Sub

strfilename = dlgSource.FileName
strSourceConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strfilename & ";UserID=Admin;Pwd="

With dlgDest
  .InitDir = App.Path & "\Database"
  .DialogTitle = "Find Destination Database"
  .Filter = "Access Databases (*.mdb)|*.mdb"
  .ShowOpen
End With

If Len(dlgDest.FileName) = 0 Then Exit Sub

strdestfile = dlgDest.FileName

strDestConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdestfile & ";UserID=Admin;Pwd="

Call testDTSClass.CreatePackage(strSourceConString, 1, strDestConString, 2)

blnRet = testDTSClass.ExecutePackage

lblCurrentSet.Caption = ""
lblRowCount.Caption = ""
If blnRet Then
  MsgBox "Execution successful"
Else
  MsgBox "Execution Failed"
End If

prgDTS.Value = 0

End Sub


Private Sub Command1_Click()
Dim strPkgName As String

'load package by name
strPkgName = InputBox("Enter the name of the package to load from the local SQL server's MSDB.")
    
If strPkgName = "" Then Exit Sub

  With dlgPkg
    .InitDir = App.Path & "\Packages"
    .DialogTitle = "Choose Directory to save Class Module"
    .Filter = "VB Class Modules (*.cls)|*.cls"
    .FileName = "ClassDTSScript.cls"
    .ShowSave
  End With
  
  If Len(dlgPkg.FileName) > 0 Then
    Call getPackage(strPkgName, "ClassDTSScript.cls")
  End If
  
End Sub


Private Sub cmdSave_Click()
 'Call testDTSClass.SavePackage
End Sub
'------------------------------------------------------------
'this function strips the file name from a path\file string
'------------------------------------------------------------
Private Function StripFileName(pFilePath As String) As String
  
  Dim lPos As Integer
  
  lPos = InStrRev(pFilePath, "\")
  
  StripFileName = Left$(pFilePath, lPos - 1)
  
End Function

Private Sub Form_Load()
  Set testDTSClass = New ClassDTSScript
  lblCurrentSet.Caption = ""
  lblRowCount.Caption = ""
End Sub

Private Sub testDTSClass_Currenttask(ByVal pCurrenttask As String)
lblCurrentSet.Caption = pCurrenttask
End Sub

Private Sub testDTSClass_PercentDone(ByVal percent As Integer)
  prgDTS.Value = percent
End Sub

Private Sub testDTSClass_RowsCopied(ByVal RowsCopied As String)
lblRowCount.Caption = RowsCopied
End Sub
