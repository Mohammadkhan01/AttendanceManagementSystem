VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdateAttendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Activity"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   Icon            =   "frmUpdateAttendance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Attendance Information"
      ForeColor       =   &H00000000&
      Height          =   5415
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3360
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "&Submit"
         Height          =   375
         Left            =   915
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4845
         Width           =   2310
      End
      Begin VB.ComboBox cmbEmployeeID 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   915
         TabIndex        =   30
         Top             =   360
         Width           =   2310
      End
      Begin VB.Frame Frame2 
         Caption         =   "Log In Date Time"
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   120
         TabIndex        =   24
         Top             =   1140
         Width           =   3105
         Begin VB.CommandButton cmdFindLoginDate 
            Height          =   300
            Left            =   2085
            Picture         =   "frmUpdateAttendance.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   27
            Tag             =   "PK"
            ToolTipText     =   "Search Screen"
            Top             =   367
            Width           =   315
         End
         Begin VB.TextBox txtLogInDate 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   795
            Locked          =   -1  'True
            TabIndex        =   26
            ToolTipText     =   "Maximum 25 Characters"
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtLoginTime 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   795
            TabIndex        =   25
            ToolTipText     =   "Maximum 25 Characters"
            Top             =   765
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   28
            Top             =   765
            Width           =   390
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Log Out Date Time"
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   3105
         Begin VB.TextBox txtLogoutTime 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   795
            TabIndex        =   21
            ToolTipText     =   "Maximum 25 Characters"
            Top             =   765
            Width           =   1275
         End
         Begin VB.TextBox txtLogoutDate 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   795
            Locked          =   -1  'True
            TabIndex        =   20
            ToolTipText     =   "Maximum 25 Characters"
            Top             =   360
            Width           =   1275
         End
         Begin VB.CommandButton cmdFindLogoutDate 
            Height          =   300
            Left            =   2085
            Picture         =   "frmUpdateAttendance.frx":09C4
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "PK"
            ToolTipText     =   "Search Screen"
            Top             =   367
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   23
            Top             =   765
            Width           =   390
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   22
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.TextBox txtOverTime 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   915
         TabIndex        =   17
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   765
         Width           =   1275
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   915
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   3900
         Width           =   2310
      End
      Begin VB.Label lblAreaName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Over Time:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   105
         TabIndex        =   34
         Top             =   765
         Width           =   780
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Remark:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   285
         TabIndex        =   33
         Top             =   3900
         Width           =   600
      End
      Begin VB.Label lblID 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   4800
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search Attendant Information"
      ForeColor       =   &H00000000&
      Height          =   7335
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   8010
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   0
         TabIndex        =   14
         Top             =   1335
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   16711680
         BackColor       =   8438015
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthBackColor  =   13629688
         StartOfWeek     =   22609921
         TitleBackColor  =   8438015
         TitleForeColor  =   0
         TrailingForeColor=   -2147483633
         CurrentDate     =   37113
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   6225
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   705
         Width           =   1665
      End
      Begin VB.TextBox txtToDate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   765
         Width           =   1275
      End
      Begin VB.CommandButton cmdToDate 
         Height          =   300
         Left            =   5670
         Picture         =   "frmUpdateAttendance.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "PK"
         ToolTipText     =   "Search Screen"
         Top             =   772
         Width           =   315
      End
      Begin VB.TextBox txtFromDate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton cmdFromDate 
         Height          =   300
         Left            =   5670
         Picture         =   "frmUpdateAttendance.frx":0BB8
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "PK"
         ToolTipText     =   "Search Screen"
         Top             =   367
         Width           =   315
      End
      Begin VB.ComboBox cmbGroupID 
         BackColor       =   &H00FDE3AE&
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbGroupName 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   765
         Width           =   2205
      End
      Begin VB.ComboBox cmbEmpID 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2205
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   7770
         _ExtentX        =   13705
         _ExtentY        =   10398
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3720
         TabIndex        =   12
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3570
         TabIndex        =   9
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Name:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   255
         TabIndex        =   5
         Top             =   765
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   255
         TabIndex        =   3
         Top             =   360
         Width           =   945
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateAttendance.frx":0CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateAttendance.frx":1104
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateAttendance.frx":1556
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateAttendance.frx":1E30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDesignation 
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   37
      Top             =   6000
      Width           =   1530
   End
   Begin VB.Label lblEmpName 
      Caption         =   "Label11"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   5685
      Width           =   1530
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1755
      Left            =   1725
      Stretch         =   -1  'True
      Top             =   5685
      Width           =   1755
   End
End
Attribute VB_Name = "frmUpdateAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmdMVClick As Integer '--For Month View Clicked

Private Sub cmbEmpId_Click()
    cmbEmpID_Validate True
End Sub

Private Sub cmbEmpID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmbEmpID_Validate True
End Sub

Private Sub cmbEmpID_Validate(Cancel As Boolean)
On Error GoTo ErrHandle:
If Trim(cmbEmpId.Text) = "" Then Exit Sub
Dim rsDept As New Recordset
Dim rsGroup As New Recordset
rsDept.Open "Select siDeptID From tblEmpDetails Where vEmployee='" & cmbEmpId.Text & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsDept.EOF Then
    rsGroup.Open "Select siGroupID From tblDepartment Where siDeptID=" & rsDept!siDeptID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsGroup.EOF Then
        cmbGroupID.Text = rsGroup!siGroupID & ""
        cmbGroupName.ListIndex = cmbGroupID.ListIndex
    Else
        MsgBox "Invalid Employee ID! Please verify Employee ID.", vbExclamation, "Remark Message"
        cmbEmpId.SelStart = 0
        cmbEmpId.SelLength = Len(cmbEmpId.Text)
        cmbEmpId.SetFocus
    End If
Else
    Set rsDept = Nothing
    Set rsGroup = Nothing
    Cancel = True
    MsgBox "Invalid Emloyee ID!", vbCritical, "Validation Message"
    cmbEmpId.SelStart = 0
    cmbEmpId.SelLength = Len(cmbEmpId)
    cmbEmpId.SetFocus
End If
Set rsDept = Nothing
Set rsGroup = Nothing
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmbEmployeeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cmbEmployeeID_Validate(Cancel As Boolean)
    lblID = ""
End Sub

Private Sub cmbGroupName_Click()
    cmbGroupID.ListIndex = cmbGroupName.ListIndex
End Sub

Private Sub cmdFindLoginDate_Click()
    cmdMVClick = 3
    MonthView1.Left = 0
    MonthView1.Top = 1335
    MonthView1.Visible = True
    MonthView1.Value = Date
End Sub

Private Sub cmdFindLogoutDate_Click()
    cmdMVClick = 4
    MonthView1.Left = 0
    MonthView1.Top = 1335
    MonthView1.Visible = True
    MonthView1.Value = Date
End Sub

Private Sub cmdFromDate_Click()
    cmdMVClick = 1
    MonthView1.Left = 5295
    MonthView1.Top = 1335
    MonthView1.Visible = True
    MonthView1.Value = Date
End Sub

Private Sub cmdSearch_Click()
'On Error GoTo ErrHandle:
If txtFromDate = "" Then
    MsgBox "From Date value can't be blank! Please Enter From Date value.", vbExclamation, "Remark Message"
    txtFromDate.SetFocus
    Exit Sub
ElseIf txtToDate = "" Then
    MsgBox "To Date value can't be blank! Please Enter To Date value.", vbExclamation, "Remark Message"
    txtToDate.SetFocus
    Exit Sub
End If
Me.MousePointer = vbHourglass
ListView1.ListItems.Clear
Dim rsEmp As New Recordset
Dim rsAttend As New Recordset
Dim itmX As ListItem
If Trim(cmbEmpId.Text) <> "" Then
    
    rsAttend.Open "Select * From tblAttend Where vEmployee='" & cmbEmpId.Text & "' And dtDate Between '" & Format(txtFromDate, "mm/dd/yyyy") & "' And '" & Format(txtToDate, "mm/dd/yyyy") & "' Order By dtDate Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsAttend.EOF Then
        Do While Not rsAttend.EOF
            Set itmX = ListView1.ListItems.Add(, , rsAttend!vEmployee, , 4)
            rsEmp.Open "Select vDesignation From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
            itmX.SubItems(1) = rsEmp!vDesignation & ""
            itmX.SubItems(2) = Format(rsAttend!dtDate, "dd/mm/yyyy") & ""
            itmX.SubItems(3) = Format(rsAttend!dtLogIn, "hh:mm AMPM")
            itmX.SubItems(4) = Format(rsAttend!vActualLogin, "HH:MM AMPM")
            itmX.SubItems(5) = rsAttend!vRemark & ""
            itmX.SubItems(6) = Format(rsAttend!vActualLogOut, "HH:MM AMPM")
            itmX.SubItems(7) = Format(rsAttend!dtLogOut, "HH:MM AMPM")
            itmX.SubItems(8) = rsAttend!iOvertime & ""
            itmX.SubItems(9) = rsAttend!iID & ""
            rsAttend.MoveNext
            rsEmp.Close
        Loop
    
    End If
ElseIf Trim(cmbEmpId.Text) = "" And cmbGroupName.Text <> "" Then
        rsAttend.Open "Select * From tblAttend Where dtDate Between '" & Format(txtFromDate, "dd/mm/yyyy") & "' And '" & Format(txtToDate, "dd/mm/yyyy") & "' Order By dtDate Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
        If Not rsAttend.EOF Then
            Dim rsFindDept As New Recordset
            Dim rsEmpGroup As New Recordset
            Do While Not rsAttend.EOF
                rsFindDept.Open "Select siDeptID From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                rsEmpGroup.Open "Select siGroupID From tblDepartment Where siDeptID=" & rsFindDept!siDeptID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
                If cmbGroupID.Text = rsEmpGroup!siGroupID Then
                    Set itmX = ListView1.ListItems.Add(, , rsAttend!vEmployee, , 4)
                    rsEmp.Open "Select vDesignation From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                    itmX.SubItems(1) = rsEmp!vDesignation & ""
                    itmX.SubItems(2) = Format(rsAttend!dtDate, "dd/mm/yyyy") & ""
                    itmX.SubItems(3) = Format(rsAttend!dtLogIn, "hh:mm AMPM")
                    itmX.SubItems(4) = Format(rsAttend!vActualLogin, "HH:MM AMPM")
                    itmX.SubItems(5) = rsAttend!vRemark & ""
                    itmX.SubItems(6) = Format(rsAttend!vActualLogOut, "HH:MM AMPM")
                    itmX.SubItems(7) = Format(rsAttend!dtLogOut, "HH:MM AMPM")
                    itmX.SubItems(8) = rsAttend!iOvertime & ""
                    itmX.SubItems(9) = rsAttend!iID & ""
                    rsEmp.Close
                End If
                rsFindDept.Close
                rsEmpGroup.Close
                rsAttend.MoveNext
            Loop
        End If
End If
Set rsAttend = Nothing
Set rsFindDept = Nothing
Set rsEmpGroup = Nothing
Set rsEmp = Nothing
ListView1.SetFocus
Me.MousePointer = vbDefault
Exit Sub
'ErrHandle:
'    Me.MousePointer = vbDefault
'    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
'    Err.Clear
End Sub

Private Sub cmdSubmit_Click()
On Error GoTo ErrHandle:
If Trim(cmbEmployeeID) = "" Then
    MsgBox "Employee ID can't be blank! Please Enter Employee ID.", vbExclamation, "Remark Message"
    cmbEmployeeID.SetFocus
    Exit Sub
ElseIf Trim(txtLogInDate) = "" Then
    MsgBox "Login Date can't be blank! Please Enter Login Date.", vbExclamation, "Remark Message"
    cmdFindLoginDate.SetFocus
    Exit Sub
End If
If lblID <> "" Then
    CnnMain.BeginTrans
    CnnMain.Execute "Exec Sp_Up_tblAttend " & lblID & ",'" & Format(txtLogInDate, "mm/dd/yyyy") & "','" & Format(txtLogInDate & " " & txtLoginTime, "mm/dd/yyyy HH:MM:SS AMPM") & "','" & Format(txtLogoutDate & " " & txtLogoutTime, "mm/dd/yyyy HH:MM:SS AMPM") & "'," & txtOverTime & ",'" & Trim(txtRemark) & "','Manual','" & gUserId & "'"
    CnnMain.CommitTrans
    MsgBox "Data successfull update to the database.", vbInformation, "Remark Message"
ElseIf lblID = "" Then
    Dim rsAttend As New Recordset
    Dim rsEmp As New Recordset
    rsAttend.Open "Select * From tblAttend Where vEmployee='" & cmbEmployeeID.Text & "' And  dtDate='" & Format(txtLogInDate, "mm/dd/yyyy") & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
    If rsAttend.EOF Then
        rsEmp.Open "Select * From tblEmpDetails Where vEmployee='" & cmbEmployeeID.Text & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
        If Not rsEmp.EOF Then
            CnnMain.BeginTrans
            CnnMain.Execute "Exec Sp_Manual_LogIn_tblAttend '" & StrConv(Trim(cmbEmployeeID.Text), vbProperCase) & "', '" & Format(txtLogInDate, "mm/dd/yyyy") & "', '" & Format(txtLogInDate & " " & txtLoginTime, "mm/dd/yyyy HH:MM:SS AMPM") & "', '" & rsEmp!vLogin & "', '" & rsEmp!vLogOut & "', " & Trim(txtOverTime) & ", '" & Trim(txtRemark) & "','Manual','" & gUserId & "'"
            CnnMain.Execute "Update tblAttend Set dtLogout='" & Format(txtLogoutDate & " " & txtLogoutTime, "mm/dd/yyyy HH:MM:SS AMPM") & "' Where vEmployee='" & cmbEmployeeID.Text & "' And dtDate='" & Format(txtLogInDate, "mm/dd/yyyy") & "'"
            CnnMain.CommitTrans
            MsgBox "Data successfull save to the database.", vbInformation, "Remark Message"
        End If
    Else
        MsgBox "You can't insert! Employee already login on that day.", vbExclamation, "Remark Message"
    End If
    Set rsEmp = Nothing
End If
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdToDate_Click()
    cmdMVClick = 2
    MonthView1.Left = 5295
    MonthView1.Top = 1335
    MonthView1.Visible = True
    MonthView1.Value = Date
End Sub

Private Sub Command1_Click()
CnnMain.BeginTrans
 CnnMain.Execute "Update tblAttend Set dtLogIn='" & Format(txtLogInDate & " " & txtLoginTime, "dd/mm/yyyy HH:MM:SS AMPM") & "' Where vEmployee='" & cmbEmployeeID.Text & "' And dtDate='" & Format(txtLogInDate, "dd/mm/yyyy") & "'"
         

'CnnMain.Execute Format(txtLogInDate & " " & txtLoginTime, "mm/dd/yyyy HH:MM:SS AMPM")
'CnnMain.Execute "Update tblAttend set dtlogin='" & Format(txtLogInDate & txtLoginTime, "mm/dd/yyyy HH:MM:SS AMPM") & "' where vEmpId='" & cmbEmployeeID.Text & "' and dtDate ='" & Format(txtLogInDate.Text, "dd/mm/yyyy") & "'"
' CnnMain.Execute "Update tblAttend Set dtLogIn='" & Format(txtLogInDate & " " & txtLoginTime, "dd/mm/yyyy HH:MM:SS AMPM") & "' Where vEmpID='" & cmbEmployeeID.Text & "' And dtDate='" & Format(txtLogInDate, "dd/mm/yyyy") & "'"
CnnMain.CommitTrans
MsgBox "Success"
cmdSearch_Click
End Sub

Private Sub Form_Click()
MonthView1.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
'==For Employee Image
Dim rsImage As New Recordset
rsImage.Open "Select vEmplName, vDesignation, iPhoto From tblEmpDetails Where vEmployee='" & gUserId & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsImage.EOF Then
    lblEmpName = rsImage!vEmplName & ""
    lblDesignation = rsImage!vDesignation & ""
    Set Image1.DataSource = rsImage
    Image1.DataField = "iPhoto"
End If
Set rsImage = Nothing
'==Data input for cmbEmpID Combo Box
Dim rsEmp As New Recordset
rsEmp.Open "Select vEmployee From tblEmpDetails Order By vEmployee Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp.EOF Then
    Do While Not rsEmp.EOF
        cmbEmpId.AddItem rsEmp!vEmployee & ""
        cmbEmployeeID.AddItem rsEmp!vEmployee & ""
        rsEmp.MoveNext
    Loop
End If
Set rsEmp = Nothing
'==
'==Data input for cmbGroupID and cmbGroupName Combo Box
Dim rsGroup As New Recordset
rsGroup.Open "Select siGroupID, vGroupName From tblGroup Order By vGroupName Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsGroup.EOF Then
    Do While Not rsGroup.EOF
        cmbGroupID.AddItem rsGroup!siGroupID & ""
        cmbGroupName.AddItem rsGroup!vGroupName & ""
        rsGroup.MoveNext
    Loop
End If
Set rsGroup = Nothing
'==For Listview1
ListView1.ColumnHeaders.Add , , "Employee", 1300
ListView1.ColumnHeaders.Add , , "Designation", 1500
ListView1.ColumnHeaders.Add , , "Date", 1200
ListView1.ColumnHeaders.Add , , "Log In", 1000
ListView1.ColumnHeaders.Add , , "Ac. Login", 1200
ListView1.ColumnHeaders.Add , , "Remark", 1500
ListView1.ColumnHeaders.Add , , "Ac. Logout", 1200
ListView1.ColumnHeaders.Add , , "Log Out", 1200
ListView1.ColumnHeaders.Add , , "Over Time", 1200
ListView1.ColumnHeaders.Add , , "ID", 500
ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
If cmbGroupName.ListCount > 0 Then cmbGroupName.ListIndex = 0
txtFromDate = Format(Now, "dd/mm/yyyy")
txtToDate = Format(Now, "dd/mm/yyyy")
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub Frame1_Click()
    MonthView1.Visible = False
End Sub

Private Sub Frame2_Click()
    MonthView1.Visible = False
End Sub

Private Sub Frame3_Click()
    MonthView1.Visible = False
End Sub

Private Sub Frame4_Click()
    MonthView1.Visible = False
End Sub

Private Sub ListView1_Click()
    MonthView1.Visible = False
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = 0
    Call Module3.lvSortByColumn(ListView1, ColumnHeader)
    ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
End Sub
Private Sub ListView1_DblClick()
On Error GoTo ErrHandle:
If ListView1.ListItems.Count > 0 Then
    cmbEmployeeID.Text = ListView1.SelectedItem.Text
    txtOverTime = 0
    txtLogInDate = Format(ListView1.SelectedItem.ListSubItems(2), "mm/dd/yyyy")
    txtLoginTime = Format(ListView1.SelectedItem.ListSubItems(3), "HH:MM:SS AMPM")
    txtLogoutDate = Format(ListView1.SelectedItem.ListSubItems(2), "mm/dd/yyyy")
    txtLogoutTime = Format(ListView1.SelectedItem.ListSubItems(7), "HH:MM:SS AMPM")
    txtRemark = ListView1.SelectedItem.ListSubItems(5) & ""
    txtOverTime = ListView1.SelectedItem.ListSubItems(8) & ""
    lblID = ListView1.SelectedItem.ListSubItems(9) & ""
End If
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle:
If KeyAscii = 13 Then
    If ListView1.ListItems.Count > 0 Then
        cmbEmployeeID.Text = ListView1.SelectedItem.Text
        txtOverTime = 0
        txtLogInDate = Format(ListView1.SelectedItem.ListSubItems(2), "dd/mm/yyyy")
        txtLoginTime = Format(ListView1.SelectedItem.ListSubItems(3), "HH:MM:SS AMPM")
        txtLogoutDate = Format(ListView1.SelectedItem.ListSubItems(2), "dd/mm/yyyy")
        txtLogoutTime = Format(ListView1.SelectedItem.ListSubItems(7), "HH:MM:SS AMPM")
        txtRemark = ListView1.SelectedItem.ListSubItems(5) & ""
        txtOverTime = ListView1.SelectedItem.ListSubItems(8) & ""
        lblID = ListView1.SelectedItem.ListSubItems(9) & ""
    End If
End If
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Select Case cmdMVClick
    Case 0
            'lblFrom = Format(DateClicked, "dd/mm/yyyy")
    Case 1
            txtFromDate = Format(DateClicked, "dd/mm/yyyy")
    Case 2
            txtToDate = Format(DateClicked, "dd/mm/yyyy")
    Case 3
            txtLogInDate = Format(DateClicked, "dd/mm/yyyy")
    Case 4
            txtLogoutDate = Format(DateClicked, "dd/mm/yyyy")
End Select
MonthView1.Visible = False
End Sub

Private Sub txtLoginTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub txtLogoutTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtOverTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
