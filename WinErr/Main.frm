VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinErr v1.2"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar P 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   8040
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog FileDlg 
      Left            =   8760
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save The Error List..."
      Filter          =   "Text Files (*.txt)|*.txt"
   End
   Begin VB.CommandButton ExportButton 
      Caption         =   "E&xport"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox ErrNo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox ErrText 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   10455
   End
   Begin MSComctlLib.ListView ErrList 
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Enter a error number (0-15999):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Here's the whole list..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, _
    ByVal lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, _
    ByVal nSize As Long, ByVal Arguments As Long) As Long

Dim i As Integer
Dim retval As Long
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Dim buffer As String
Dim ListObj As ListItem

Private Sub ErrList_ItemClick(ByVal Item As MSComctlLib.ListItem)
ErrText.Text = ErrList.ListItems(Item.Index).SubItems(1)
ErrNo.Text = Item.Text
End Sub

Private Sub ErrNo_Change()
ErrText.Text = ""
If ErrNo.Text = "" Then
    ErrText.Text = ""
End If
buffer = Space(255)
On Error GoTo out
i = CInt(ErrNo.Text)
    'SetLastError (i)
    retval = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, "", i, 0, buffer, Len(buffer), 0)
    buffer = Trim(buffer)
    If Len(buffer) > 2 Then
        buffer = Replace(buffer, vbCrLf, "")
        ErrText.Text = buffer
    End If
    Exit Sub
out:
ErrNo.Text = ""
End Sub

Private Sub ExportButton_Click()
buffer = ""
On Error Resume Next
FileDlg.CancelError = True
FileDlg.DefaultExt = ".txt"
FileDlg.ShowSave
If Err.Number = cdlCancel Then
    Exit Sub
End If
P.Max = ErrList.ListItems.Count
P.Value = 0
P.Visible = True
Open FileDlg.FileName For Binary As 1
For i = 1 To ErrList.ListItems.Count
    buffer = buffer & ErrList.ListItems(i).Text & vbCrLf & ErrList.ListItems(i).SubItems(1) & vbCrLf & vbCrLf
P.Value = i
DoEvents
Next
Put 1, , buffer
Close 1
MsgBox "Export complete!!", vbExclamation, "Done"
P.Value = 0
P.Visible = False
End Sub

Private Sub Form_Load()
Dim ColHead As ColumnHeader
ErrList.ColumnHeaders.Clear
Set ColHead = ErrList.ColumnHeaders.Add(1, , "Error Number", 1150)
Set ColHead = ErrList.ColumnHeaders.Add(2, , "Description", 9000)

For i = 0 To 15999
    buffer = Space(255)
    retval = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, "", i, 0, buffer, Len(buffer), 0)
    buffer = Trim(buffer)
    If Len(buffer) > 2 Then
        Dim j As Integer
        buffer = Replace(buffer, vbCrLf, "")
        Set ListObj = ErrList.ListItems.Add(, , i)
        ListObj.SubItems(1) = buffer
    End If
Next
End Sub
