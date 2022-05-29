VERSION 5.00
Begin VB.Form frmDefaultMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Object Information"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmDefaultMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6075
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Navigation"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   5820
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   135
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   375
         Left            =   4575
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3465
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   2355
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1245
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   405
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Object Information"
      ForeColor       =   &H00000000&
      Height          =   1620
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5820
      Begin VB.TextBox txtMenuCaption 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   720
         Width           =   4110
      End
      Begin VB.ComboBox cmbStatus 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmDefaultMenu.frx":08CA
         Left            =   1245
         List            =   "frmDefaultMenu.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdFind 
         Height          =   300
         Left            =   5370
         Picture         =   "frmDefaultMenu.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "PK"
         ToolTipText     =   "Search Screen"
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtMenuName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1245
         MaxLength       =   25
         TabIndex        =   1
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   360
         Width           =   4110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Object Caption:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Object Name:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Value:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   10
         Top             =   1080
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmDefaultMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo ErrHandle:
    Me.MousePointer = vbHourglass
    txtMenuName = ""
    txtMenuCaption = ""
    cmbStatus.ListIndex = 0
    prc_DisAllow
    Me.MousePointer = vbDefault
    Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdFind_Click()
On Error GoTo ErrHandle:
    frmSearchMenuInfo.Show vbModal
If strfindcode <> "" Then
    txtMenuName = strfindcode
    txtMenuName_Validate True
    strfindcode = ""
Else
    '--Nothing
End If
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrHandle:
    Me.MousePointer = vbHourglass
    cmbStatus.ListIndex = 0
    txtMenuName = ""
    txtMenuCaption = ""
    Call prc_Allow
    txtMenuName.SetFocus
    Me.MousePointer = vbDefault
    Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandle:
If Trim(txtMenuName) = "" Then
    MsgBox "Object Name can't be blank! Please Enter Object Name.", vbExclamation, "Validation Message"
    txtMenuName.SetFocus
    Exit Sub
ElseIf Trim(txtMenuCaption) = "" Then
    MsgBox "Object Caption can't be blank! Please Enter Object Caption.", vbExclamation, "Validation Message"
    txtMenuCaption.SetFocus
    Exit Sub
End If
'--sp_Add_tblLogin
    CnnMain.BeginTrans
    CnnMain.Execute "Exec Sp_Add_tblObject '" & txtMenuName & "','" & txtMenuCaption & "','" & cmbStatus.Text & "'"
    CnnMain.CommitTrans
Call prc_DisAllow
Me.MousePointer = vbDefault
MsgBox "Record successfully saved to the database.", vbInformation, "Save Message"
txtMenuName.SelLength = Len(txtMenuName)
txtMenuName.SetFocus
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrHandle:
If Trim(txtMenuName) = "" Then
    MsgBox "Object Name can't be blank! Please Enter Object Name.", vbExclamation, "Validation Message"
    txtMenuName.SetFocus
    Exit Sub
ElseIf Trim(txtMenuCaption) = "" Then
    MsgBox "Object Caption can't be blank! Please Enter Object Caption.", vbExclamation, "Validation Message"
    txtMenuCaption.SetFocus
    Exit Sub
End If
'--SP_Up_tblClient
Me.MousePointer = vbHourglass
CnnMain.BeginTrans
CnnMain.Execute "Exec Sp_Up_tblObject '" & txtMenuName & "','" & txtMenuCaption & "','" & cmbStatus.Text & "'"
CnnMain.CommitTrans
Me.MousePointer = vbDefault
MsgBox "Record successfully updated to the database.", vbInformation, "Update Message"
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "Error Number : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Error Message"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
    Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
    Call prc_DisAllow
    Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub txtMenuName_Validate(Cancel As Boolean)
On Error GoTo ErrHandle:
If txtMenuName = "" Then Exit Sub
Dim rsMenuInfo As New ADODB.Recordset
rsMenuInfo.Open "Select * From tblObject Where vObjectName='" & txtMenuName & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsMenuInfo.EOF Then
    With rsMenuInfo
        txtMenuCaption = !vObjectCaption & ""
        cmbStatus.Text = !vDefault & ""
    End With
    prc_DisAllow
    txtMenuCaption.SelStart = 0
    txtMenuCaption.SelLength = Len(txtMenuCaption)
    txtMenuCaption.SetFocus
    Set rsMenuInfo = Nothing
Else
    txtMenuCaption = ""
    cmbStatus.ListIndex = 0
    Dim strMenuName As String
    strMenuName = txtMenuName
    cmdNew_Click
    txtMenuName = strMenuName
    txtMenuCaption.SetFocus
End If
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Public Sub prc_DisAllow()
Dim obj As Object
  For Each obj In Me
        If TypeOf obj Is CommandButton Then
                If obj.Name = "cmdSave" Or obj.Name = "cmdCancel" Then
                obj.Enabled = False
        Else
                obj.Enabled = True
                End If
        End If
    Next
End Sub
Public Sub prc_Allow()
Dim obj As Object
    For Each obj In Me
        If TypeOf obj Is CommandButton Then
            If obj.Name = "cmdSave" Or obj.Name = "cmdCancel" Or obj.Name = "cmdBrowse" Or obj.Name = "Command1" Then
                obj.Enabled = True
            Else
                obj.Enabled = False
            End If
        End If
    Next
End Sub

Private Sub txtMenuCaption_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If cmdSave.Enabled = True Then
            cmdSave.SetFocus
        End If
    End If
End Sub

Private Sub txtMenuName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtMenuName_Validate True
End Sub
