VERSION 5.00
Begin VB.Form frmDBConnInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Server Connectivity"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ControlBox      =   0   'False
   ForeColor       =   &H0094B5BA&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtUPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1185
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   420
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1845
      Left            =   120
      Picture         =   "frmDBConnInfo.frx":0000
      Stretch         =   -1  'True
      Top             =   210
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   15
      X2              =   6600
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   6600
      Y1              =   2175
      Y2              =   2175
   End
End
Attribute VB_Name = "frmDBConnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
On Error GoTo ErrHandle:
If Me.Caption = "Changing Database Connection" Then
    Unload Me
ElseIf gblnConnFromMnu = True Then
    Unload Me
Else
    End
End If
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub cmdOk_Click()
On Error GoTo ErrHand
If Check_all_textboxes = False Then
    MsgBox "Be sure to input valid data.", vbInformation, "Invalid data"
    txtServName.SetFocus
    Exit Sub
End If
'SaveSetting App.EXEName, "Section1", "Server_Name", txtServName.Text
'SaveSetting App.EXEName, "Section1", "DB_Name", txtDBName.Text
'SaveSetting App.EXEName, "Section1", "User_ID", decode(txtUID.Text)
SaveSetting App.EXEName, "Section1", "User_Password", decode(txtUPassword.Text)
Unload Me
Call Main
Exit Sub
ErrHand:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Public Function Check_all_textboxes() As Boolean
On Error GoTo ErrHandle:
Dim ctrControl As Control
Check_all_textboxes = True
For Each ctrControl In frmDBConnInfo
    If TypeOf ctrControl Is TextBox And ctrControl.Name <> "txtUPassword" Then
        If Trim(ctrControl.Text) = "" Then
            Check_all_textboxes = False
            Exit Function
        End If
    End If
Next
Set ctrControl = Nothing
Exit Function
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Function
Private Sub Show_Text_Data()
'txtServName = gstrServName
'txtDBName = gstrDBName
'txtUID = gstrDBUID
txtUPassword = gstrDBUpass
End Sub
Private Sub Form_Activate()
    txtUPassword.SetFocus
End Sub
Private Sub Form_Load()
    Call Show_Text_Data
End Sub
Public Function decode(cString) As String
On Error GoTo ErrHandle:
Dim lngAsc As Long
Dim strDecode As String
Dim j As Integer
strDecode = ""
For j = 1 To Len(cString)
    lngAsc = Asc(Mid(cString, j, 1)) + (j + 19)
    strDecode = strDecode + Chr$(lngAsc)
Next
decode = strDecode
Exit Function
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Function


Private Sub txtUPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
