VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Management System"
   ClientHeight    =   4635
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   8880
   FontTransparent =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2738.506
   ScaleMode       =   0  'User
   ScaleWidth      =   8337.84
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00EF9D7E&
      Cancel          =   -1  'True
      Caption         =   "Cance&l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7545
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3907
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00EF9D7E&
      Caption         =   "&Connect"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3907
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Use SQL Server Authentication"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1695
      Left            =   4680
      TabIndex        =   2
      Top             =   1680
      Width           =   4050
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FDE3AE&
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
         Left            =   1200
         MaxLength       =   10
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   1080
         Width           =   2625
      End
      Begin VB.TextBox txtUserId 
         BackColor       =   &H00FDE3AE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   0
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Login ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   8
         Top             =   735
         Width           =   795
      End
      Begin VB.Image Image3 
         Height          =   510
         Left            =   135
         Picture         =   "frmLogin.frx":08CA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Login Authentication"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   180
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   3
         Top             =   1080
         Width           =   885
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3240
      Left            =   120
      TabIndex        =   5
      Top             =   135
      Width           =   4425
      _cx             =   4202109
      _cy             =   4200019
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENDANCE SOFTWARE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4710
      TabIndex        =   9
      Top             =   135
      Width           =   4050
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   45
      Picture         =   "frmLogin.frx":ABC0
      Top             =   3750
      Width           =   5760
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   -240
      Picture         =   "frmLogin.frx":BA90
      Stretch         =   -1  'True
      Top             =   3525
      Width           =   10170
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'Set CnnMain = Nothing
    'End
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
If txtUserId <> "" Then
    Dim tmpRec As New Recordset
    tmpRec.Open "Select vEmployee, Password, vActive, siDeptID From tblEmpDetails Where vEmployee='" & txtUserId & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not tmpRec.EOF Then
        Dim rsGroup As New Recordset
        If UCase(txtUserId) = UCase(tmpRec!vEmployee) And UCase(txtPassword) = UCase(tmpRec!Password) And tmpRec!vActive = "Yes" Then
            Module2.gUserId = StrConv(Trim(txtUserId), vbProperCase)
            Module2.gPassword = StrConv(Trim(txtPassword), vbProperCase)
'            Module2.gLevel = Trim(tmpRec!vRole)
            rsGroup.Open "Select siGroupID From tblDepartment Where siDeptID=" & tmpRec!siDeptID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
            Module2.giGroupID = Trim(rsGroup!siGroupID)
            Set rsGroup = Nothing
        ElseIf tmpRec!vActive = "No" Then
            Me.MousePointer = vbDefault
            MsgBox "Inactive Login ID! Please  contact with your Administrator.", vbCritical, "Contact Administrator"
            Set tmpRec = Nothing
            End
        Else
            Me.MousePointer = vbDefault
            MsgBox "Invalid Password! Please enter valid Password.", vbCritical, "Login Information"
            txtPassword.SelStart = 0
            txtPassword.SelLength = Len(txtPassword)
            txtPassword.SetFocus
            Exit Sub
        End If
    Else
        Me.MousePointer = vbDefault
        MsgBox "Invalid Login ID! Please enter valid Login ID.", vbCritical, "Login Information"
        txtUserId.SelStart = 0
        txtUserId.SelLength = Len(txtUserId)
        txtUserId.SetFocus
        Exit Sub
    End If
Else
    MsgBox "Please enter valid Login ID!", vbCritical, "Remark Message"
    txtUserId.SetFocus
    Exit Sub
End If
Set tmpRec = Nothing
'--Initialize Company Information
Dim rsCompany As New Recordset
rsCompany.Open "Select vCompanyName From tblCompanyInfo", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsCompany.EOF Then
    vgCompanyName = rsCompany!vCompanyName & ""
End If
Set rsCompany = Nothing
'--Initialize Option Table
'Dim rsOption As New Recordset
'rsOption.Open "Select * From tblOption", CnnMain, adOpenStatic, adLockReadOnly
'igMinPassDigit = rsOption!iMinPassDigit & ""
'igMonitorRefreshTime = rsOption!iMonitorRefreshTime & ""
'igSoftRunningCheck = rsOption!iSoftRunningCheck & ""
'vgBF = rsOption!vBF & ""
'Set rsOption = Nothing
Unload Me
frmMain.Show
Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
'    CnnMain.RollbackTrans
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
App.TaskVisible = False
ShockwaveFlash1.Movie = App.Path & "\Flash\IMART.swf"
ShockwaveFlash1.Play
Me.MousePointer = vbHourglass
Dim sBuffer As String
Dim lSize As Long
sBuffer = Space$(255)
lSize = Len(sBuffer)
Call GetComputerName(sBuffer, lSize)
If lSize > 0 Then
        gstrMachineName = Left$(sBuffer, lSize)
Else
        gstrMachineName = vbNullString
End If
Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtUserId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
