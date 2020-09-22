VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1560
      IntegralHeight  =   0   'False
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   120
      Width           =   8000
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1425
      ItemData        =   "Form1.frx":0004
      Left            =   120
      List            =   "Form1.frx":0006
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1800
      Width           =   8000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   9999
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Open"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo MS ' Error control
    Dim vFiles As Variant
    Dim lFile As Long
    With CommonDialog1
        .FileName = "" 'Clear the filename
        .CancelError = True 'Gives an error if cancel is pressed
        .DialogTitle = "Select File(s)..."
        .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Falgs, allows Multi select, Explorer style and hide the Read only tag
        .Filter = "All files (*.*)|*.*"
        .ShowOpen
        vFiles = Split(.FileName, Chr(0)) 'Splits the filename up in segments
    If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
    List1.AddItem .FileName
    List2.AddItem .FileTitle
    Else
    For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
    List1.AddItem vFiles(0) + "\" & vFiles(lFile)
    List2.AddItem vFiles(lFile)
    Next
    End If
    End With
MS:
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub
