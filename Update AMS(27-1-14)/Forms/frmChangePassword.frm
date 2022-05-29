VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3555
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1845
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1950
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Password"
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   135
      TabIndex        =   4
      Top             =   120
      Width           =   3285
      Begin VB.TextBox txtConfirmPassword 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1545
         MaxLength       =   10
         PasswordChar    =   "l"
         TabIndex        =   2
         Top             =   1170
         Width           =   1575
      End
      Begin VB.TextBox txtNewPassword 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1545
         MaxLength       =   10
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   750
         Width           =   1575
      End
      Begin VB.TextBox txtOldPassword 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1545
         MaxLength       =   10
         PasswordChar    =   "l"
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   1170
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   6
         Top             =   750
         Width           =   1110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   5
         Top             =   360
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdate_Click()
On Error GoTo ErrHandle:
'--Validation Check
If Trim(txtOldPassword) = "" Then
    MsgBox "Old password can't be blank!" & "  Please enter old password", vbExclamation, "Remark Message"
    txtOldPassword.SetFocus
    Exit Sub
ElseIf Trim(txtNewPassword) = "" Then
    MsgBox "New password can't be blank! Please enter new password.", vbExclamation, "Remark Message"
    txtNewPassword.SetFocus
    Exit Sub
ElseIf Len(Trim(txtNewPassword)) < 6 Then
    MsgBox "New Password can't be less then 6 Digit!" & "  Please enter valid New Password", vbExclamation, "Remark Message"
    txtNewPassword.SelStart = 0
    txtNewPassword.SelLength = Len(txtNewPassword)
    txtNewPassword.SetFocus
    Exit Sub
ElseIf UCase(txtConfirmPassword) <> UCase(txtNewPassword) Then
    MsgBox "Confirm Password does't match with New Password! Please verify your Confirm Password!", vbCritical, "Remark Message"
    txtConfirmPassword.SelStart = 0
    txtConfirmPassword.SelLength = Len(txtOldPassword)
    txtConfirmPassword.SetFocus
    Exit Sub
End If
'--Update Password
CnnMain.BeginTrans
CnnMain.Execute "Update tblEmpDetails  Set Password='" & txtNewPassword & "' Where vEmployee='" & gUserId & "'"
CnnMain.CommitTrans
Me.MousePointer = vbDefault
MsgBox "Record successfully updated to the database.", vbInformation, "Update Message"
txtOldPassword = ""
txtNewPassword = ""
txtConfirmPassword = ""
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    CnnMain.RollbackTrans
    MsgBox "Error Number : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Error Message"
    Err.Clear
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
End Sub

Private Sub txtConfirmPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdUpdate.SetFocus
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtConfirmPassword.SetFocus
End Sub

Private Sub txtOldPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
If KeyAscii = 13 Then txtOldPassword_Validate True
End Sub

Private Sub txtOldPassword_Validate(Cancel As Boolean)
If txtOldPassword = "" Then Exit Sub
If UCase(txtOldPassword) <> UCase(gPassword) Then
    MsgBox "Old Password doesn't match! Please verify your Old Password!", vbCritical, "Remark Message"
    Cancel = True
    txtOldPassword.SelStart = 0
    txtOldPassword.SelLength = Len(txtOldPassword)
    txtOldPassword.SetFocus
    Exit Sub
End If
txtNewPassword.SetFocus
End Sub
