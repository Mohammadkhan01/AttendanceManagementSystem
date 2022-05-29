VERSION 5.00
Begin VB.Form frmComInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organization Information"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmComInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5715
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3315
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4215
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3315
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Organization Information"
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   150
      TabIndex        =   8
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtComWeb 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1380
         MaxLength       =   25
         TabIndex        =   5
         ToolTipText     =   "Maximum 50 Characters"
         Top             =   2490
         Width           =   3855
      End
      Begin VB.TextBox txtComName 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Maximum 50 Characters"
         Top             =   330
         Width           =   3855
      End
      Begin VB.TextBox txtComAdd 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   1380
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Maximum 200 Characters"
         Top             =   690
         Width           =   3855
      End
      Begin VB.TextBox txtComPhone 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1380
         MaxLength       =   35
         TabIndex        =   2
         ToolTipText     =   "Maximum 50 Characters"
         Top             =   1410
         Width           =   3855
      End
      Begin VB.TextBox txtComFax 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Maximum 50 Characters"
         Top             =   1770
         Width           =   3855
      End
      Begin VB.TextBox txtComMail 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1380
         MaxLength       =   25
         TabIndex        =   4
         ToolTipText     =   "Maximum 50 Characters"
         Top             =   2130
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Site:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   630
         TabIndex        =   14
         Top             =   2490
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Organization:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   13
         Top             =   330
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   12
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   825
         TabIndex        =   11
         Top             =   1410
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   1035
         TabIndex        =   10
         Top             =   1770
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   870
         TabIndex        =   9
         Top             =   2130
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmComInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandle:
'---Validation Check
If Trim(txtComName) = "" Then
    MsgBox "Company Name can't be blank!" & "  Please Enter Company Name!", vbExclamation, "Remark Message"
    txtComName.SetFocus
    Exit Sub
ElseIf Trim(txtComAdd) = "" Then
        MsgBox "Company Address can't be blank!" & "  Please Enter Company Address!", vbExclamation, "Remark Message"
        txtComAdd.SetFocus
        Exit Sub
End If
Me.MousePointer = vbHourglass
'---Data Save or Update
Dim sql As String
Dim rsComInfo As New Recordset
rsComInfo.Open "Select vCompanyName from tblCompanyInfo", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsComInfo.EOF Then
    '--Update ComInfo
    sql = "Exec Sp_Up_tblCompanyInfo  '" & Trim(txtComName) & "','" & Trim(txtComAdd) & "'," & _
          " '" & Trim(txtComPhone) & "', '" & Trim(txtComFax) & "', '" & Trim(txtComMail) & "','" & Trim(txtComWeb) & "', '" & gUserId & "'"
    CnnMain.BeginTrans
    CnnMain.Execute sql
    CnnMain.CommitTrans
    MsgBox "Data successfully updated to the database.", vbInformation, "Successfull Message"
Else
    sql = "Exec Sp_Add_tblCompanyInfo '" & Trim(txtComName) & "','" & Trim(txtComAdd) & "'," & _
          " '" & Trim(txtComPhone) & "', '" & Trim(txtComFax) & "', '" & Trim(txtComMail) & "','" & Trim(txtComWeb) & "', '" & gUserId & "'"
    CnnMain.BeginTrans
    CnnMain.Execute sql
    CnnMain.CommitTrans
    MsgBox "Data successfully saved to the database.", vbInformation, "Successfull Message"
End If
Set rsComInfo = Nothing
txtComName.SelStart = 0
txtComName.SelLength = Len(txtComName)
txtComName.SetFocus
Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.Move (Screen.Width - Me.Width) / 2, (Me.ScaleHeight - Me.Height) / 2
Dim rsComInfo As New Recordset
rsComInfo.Open "Select * from tblCompanyInfo", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsComInfo.EOF Then
    txtComName = rsComInfo(0) & ""
    txtComAdd = rsComInfo(1) & ""
    txtComPhone = rsComInfo(2) & ""
    txtComFax = rsComInfo(3) & ""
    txtComMail = rsComInfo(4) & ""
    txtComWeb = rsComInfo(5) & ""
End If
Set rsComInfo = Nothing
txtComName.SelStart = 0
txtComName.SelLength = Len(txtComName)
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub txtComAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub txtComFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub txtComMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtComName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub txtComPhone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub txtComWeb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
