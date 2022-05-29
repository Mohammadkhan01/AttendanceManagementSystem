VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAttendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Management System"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "frmAttendance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "List of Group Information"
      ForeColor       =   &H00000000&
      Height          =   4665
      Left            =   120
      TabIndex        =   14
      Top             =   2775
      Width           =   3360
      Begin VB.CommandButton cmdShowAll 
         Caption         =   "Show A&ll"
         Height          =   375
         Left            =   1665
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4155
         Width           =   1560
      End
      Begin VB.CommandButton cmdShowGroup 
         Caption         =   "Show &Group"
         Height          =   375
         Left            =   135
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4155
         Width           =   1515
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3825
         Left            =   135
         TabIndex        =   5
         Top             =   240
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   6747
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
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00EF9D7E&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attendance Information"
      ForeColor       =   &H00000000&
      Height          =   2565
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3360
      Begin VB.CommandButton cmdApplication 
         Caption         =   "&Application"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2040
         Width           =   1515
      End
      Begin VB.ComboBox cmbRemark 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   915
         TabIndex        =   2
         Top             =   1170
         Width           =   2310
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "&Submit"
         Height          =   375
         Left            =   1650
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton optLogOut 
         Caption         =   "Log Out"
         Height          =   195
         Left            =   2205
         TabIndex        =   8
         Top             =   1650
         Width           =   1020
      End
      Begin VB.OptionButton optLogIn 
         Caption         =   "Log In"
         Height          =   195
         Left            =   915
         TabIndex        =   3
         Top             =   1650
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.TextBox txtPassword 
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
         ForeColor       =   &H00000000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   915
         MaxLength       =   10
         PasswordChar    =   "l"
         TabIndex        =   1
         ToolTipText     =   "Maximum 10 Characters"
         Top             =   765
         Width           =   2310
      End
      Begin VB.TextBox txtEmpID 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   915
         MaxLength       =   25
         TabIndex        =   0
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   360
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
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Remark:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   285
         TabIndex        =   12
         Top             =   1170
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   765
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "List of Attendant Employee"
      ForeColor       =   &H00000000&
      Height          =   7335
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   7650
      Begin MSComctlLib.ListView ListView1 
         Height          =   6975
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   12303
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmAttendance.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttendance.frx":2BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttendance.frx":3046
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttendance.frx":3920
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApplication_Click()
'Dim rsEmp As New Recordset
'Dim rsCB As New Recordset
'rsEmp.Open "Select vEmpID From tblEmpDetails", CnnMain, adOpenForwardOnly, adLockReadOnly
'Do While Not rsEmp.EOF
'    rsCB.Open "Select vPassword From CBtblEmpDetails Where vEmpID='" & rsEmp!vEmpID & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
'    If Not rsCB.EOF Then
'        CnnMain.BeginTrans
'        CnnMain.Execute "Update tblEmpDetails set vPassword= '" & rsCB!vPassword & "' Where vEmpID='" & rsEmp!vEmpID & "'"
'        CnnMain.CommitTrans
'    End If
'    rsCB.Close
'    rsEmp.MoveNext
'Loop
'Set rsCB = Nothing
'Set rsEmp = Nothing
If gFlag = False Then
gFlag = True
frmLogin.Show
End If
End Sub

Private Sub cmdRefresh_Click()
'On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
ListView1.ListItems.Clear
Dim rsFindDept As New Recordset
Dim rsGroup As New Recordset
rsFindDept.Open "Select siDeptID From tblEmpDetails Where vEmployee='" & Trim(txtEmpID) & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
rsGroup.Open "Select siGroupID From tblDepartment Where siDeptID=" & rsFindDept!siDeptID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
rsFindDept.Close
If Not rsGroup.EOF Then
    Dim rsAttend As New Recordset
    Dim rsDate As New Recordset
    rsDate.Open "Select getdate()", CnnMain, adOpenForwardOnly, adLockReadOnly
    'rsAttend.Open "Select * From tblAttend Where siGroupID=" & rsFindGroup!siGroupID & " And dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Null", CnnMain, adOpenForwardOnly, adLockReadOnly
    'rsAttend.Open "Select * From tblAttend Where (siGroupID=" & rsFindGroup!siGroupID & " And dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Null) OR (siGroupID=" & rsFindGroup!siGroupID & " And dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Not Null And vRemark<>'')", CnnMain, adOpenForwardOnly, adLockReadOnly
    rsAttend.Open "Select * From tblAttend Where (dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Null) OR (dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Not Null And vRemark<>'')", CnnMain, adOpenForwardOnly, adLockReadOnly
    Set rsDate = Nothing
    If Not rsAttend.EOF Then
        Dim itmX As ListItem
        Dim rsEmpGroup As New Recordset
        Dim rsEmp As New Recordset
        Do While Not rsAttend.EOF
            rsFindDept.Open "Select siDeptID From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
            rsEmpGroup.Open "Select siGroupID From tblDepartment Where siDeptID=" & rsFindDept!siDeptID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
            If rsGroup!siGroupID = rsEmpGroup!siGroupID Then
                Set itmX = ListView1.ListItems.Add(, , rsAttend!vEmployee, , 4)
                rsEmp.Open "Select vDesignation From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                itmX.SubItems(1) = rsEmp!vDesignation & ""
                itmX.SubItems(2) = Format(rsAttend!dtLogIn, "hh:mm AMPM")
                itmX.SubItems(3) = Format(rsAttend!vActualLogin, "HH:MM AMPM")
                itmX.SubItems(4) = rsAttend!vRemark & ""
                itmX.SubItems(5) = Format(rsAttend!vActualLogOut, "HH:MM AMPM")
                itmX.SubItems(6) = Format(rsAttend!dtLogOut, "HH:MM AMPM")
                rsEmp.Close
            End If
            rsAttend.MoveNext
            rsFindDept.Close
            rsEmpGroup.Close
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

Private Sub cmdShowAll_Click()
'On Error GoTo ErrHandle:
If Trim(txtEmpID) = "" Then
    MsgBox "Login ID can't be blank! Please Enter Login ID.", vbExclamation, "Remark Message"
    txtEmpID.SetFocus
    Exit Sub
ElseIf Trim(txtPassword) = "" Then
    MsgBox "Password can't be blank! Please Enter valid password.", vbExclamation, "Remark Message"
    txtPassword.SetFocus
    Exit Sub
End If
ListView1.ListItems.Clear
Dim tmpRec As New Recordset
tmpRec.Open "Select vEmployee, Password From tblEmpDetails Where vEmployee = '" & txtEmpID & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not tmpRec.EOF Then
    If UCase(txtEmpID) = UCase(tmpRec!vEmployee) And UCase(txtPassword) = UCase(tmpRec!Password) Then
        '==Permission wise Group Attendance
        Me.MousePointer = vbHourglass
        Dim rsGroup As New Recordset
        Dim rsAttend As New Recordset
        Dim rsDate As New Recordset
        rsDate.Open "Select getdate()", CnnMain, adOpenForwardOnly, adLockReadOnly
        rsAttend.Open "Select * From tblAttend Where (dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Null) OR (dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Not Null And vRemark<>'')", CnnMain, adOpenForwardOnly, adLockReadOnly
        Set rsDate = Nothing
        If Not rsAttend.EOF Then
            Dim rsFindDept As New Recordset
            Dim rsFindGroup As New Recordset
            Dim rsPermission As New Recordset
            Do While Not rsAttend.EOF
                rsFindDept.Open "Select siDeptID From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                rsFindGroup.Open "Select siGroupID From tblDepartment Where siDeptID=" & rsFindDept!siDeptID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
                rsPermission.Open "Select vEmployee, vEnable From tblObjectPermission Where vEmployee='" & Trim(txtEmpID) & "' And vObjectName='" & rsFindGroup!siGroupID & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                If Not rsPermission.EOF Then
                    Dim itmX As ListItem
                    Dim rsEmp As New Recordset
                    Set itmX = ListView1.ListItems.Add(, , rsAttend!vEmployee, , 4)
                     rsEmp.Open "Select vDesignation From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                     itmX.SubItems(1) = rsEmp!vDesignation & ""
                     itmX.SubItems(2) = Format(rsAttend!dtLogIn, "hh:mm AMPM")
                     itmX.SubItems(3) = Format(rsAttend!vActualLogin, "HH:MM AMPM")
                     itmX.SubItems(4) = rsAttend!vRemark & ""
                     itmX.SubItems(5) = Format(rsAttend!vActualLogOut, "HH:MM AMPM")
                     itmX.SubItems(6) = Format(rsAttend!dtLogOut, "HH:MM AMPM")
                     rsEmp.Close
                End If
                rsFindDept.Close
                rsFindGroup.Close
                rsPermission.Close
                rsAttend.MoveNext
            Loop
        End If
    Else
        MsgBox "Invalid Password! Please enter valid Password.", vbCritical, "Validation Message"
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword)
        txtPassword.SetFocus
        Exit Sub
    End If
Else
    MsgBox "Invalid Employee Name! Please enter valid Employee Name.", vbCritical, "Validation Message"
    txtEmpID.SelStart = 0
    txtEmpID.SelLength = Len(txtEmpID)
    txtEmpID.SetFocus
    Exit Sub
End If
Set tmpRec = Nothing
Set rsAttend = Nothing
Set rsFindDept = Nothing
Set rsFindGroup = Nothing
Set rsPermission = Nothing
Set rsEmp = Nothing
Me.MousePointer = vbDefault
Exit Sub
'ErrHandle:
'    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
'    Err.Clear
End Sub

Private Sub cmdShowGroup_Click()
'On Error GoTo ErrHandle:
If Trim(txtEmpID) = "" Then
    MsgBox "Login ID can't be blank! Please Enter Login ID.", vbExclamation, "Remark Message"
    txtEmpID.SetFocus
    Exit Sub
ElseIf Trim(txtPassword) = "" Then
    MsgBox "Password can't be blank! Please Enter valid password.", vbExclamation, "Remark Message"
    txtPassword.SetFocus
    Exit Sub
End If
Dim tmpRec As New Recordset
tmpRec.Open "Select vEmployee, Password From tblEmpDetails Where vEmployee = '" & txtEmpID & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not tmpRec.EOF Then
    If UCase(txtEmpID) = UCase(tmpRec!vEmployee) And UCase(txtPassword) = UCase(tmpRec!Password) Then
        '===Check for Permission
        Dim rsPermission As New Recordset
        rsPermission.Open "Select vEnable From tblObjectPermission Where vEmployee='" & Trim(txtEmpID) & "' And vObjectName='" & ListView2.SelectedItem.Text & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
        If rsPermission.EOF Then
           MsgBox "Access denied for user : " & txtEmpID & "! You have no permission for this Group.", vbCritical, "Validation Message"
           ListView1.ListItems.Clear
           ListView2.SetFocus
           Exit Sub
        Else
        End If
        Set rsPermission = Nothing
        Me.MousePointer = vbHourglass
        ListView1.ListItems.Clear
        Dim rsAttend As New Recordset
        Dim rsDate As New Recordset
        rsDate.Open "Select getdate()", CnnMain, adOpenForwardOnly, adLockReadOnly
        rsAttend.Open "Select * From tblAttend Where (dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Null) OR (dtDate = '" & Format(rsDate(0), "mm/dd/yyyy") & "' And dtLogOut is Not Null And vRemark<>'')", CnnMain, adOpenForwardOnly, adLockReadOnly
        Set rsDate = Nothing
        If Not rsAttend.EOF Then
            Dim itmX As ListItem
            Dim rsFindDept As New Recordset
            Dim rsEmpGroup As New Recordset
            Dim rsEmp As New Recordset
            Do While Not rsAttend.EOF
                rsFindDept.Open "Select siDeptID From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                rsEmpGroup.Open "Select siGroupID From tblDepartment Where siDeptID=" & rsFindDept!siDeptID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
                If ListView2.SelectedItem.Text = rsEmpGroup!siGroupID Then
                    Set itmX = ListView1.ListItems.Add(, , rsAttend!vEmployee, , 4)
                    rsEmp.Open "Select vDesignation From tblEmpDetails Where vEmployee='" & rsAttend!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                    itmX.SubItems(1) = rsEmp!vDesignation & ""
                    itmX.SubItems(2) = Format(rsAttend!dtLogIn, "hh:mm AMPM")
                    itmX.SubItems(3) = Format(rsAttend!vActualLogin, "HH:MM AMPM")
                    itmX.SubItems(4) = rsAttend!vRemark & ""
                    itmX.SubItems(5) = Format(rsAttend!vActualLogOut, "HH:MM AMPM")
                    itmX.SubItems(6) = Format(rsAttend!dtLogOut, "HH:MM AMPM")
                    rsEmp.Close
                End If
                rsAttend.MoveNext
                rsFindDept.Close
                rsEmpGroup.Close
            Loop
        End If
        Set rsAttend = Nothing
        Set rsFindDept = Nothing
        Set rsEmpGroup = Nothing
        Set rsEmp = Nothing
        ListView1.SetFocus
        Me.MousePointer = vbDefault
    Else
        MsgBox "Invalid Password! Please enter valid Password.", vbCritical, "Validation Message"
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword)
        txtPassword.SetFocus
        Exit Sub
    End If
Else
    MsgBox "Invalid Employee Name! Please enter valid Employee Name.", vbCritical, "Validation Message"
    txtEmpID.SelStart = 0
    txtEmpID.SelLength = Len(txtEmpID)
    txtEmpID.SetFocus
    Exit Sub
End If
Set tmpRec = Nothing
Exit Sub
'ErrHandle:
'    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
'    Err.Clear
End Sub

Private Sub cmdSubmit_Click()
'On Error GoTo ErrHandle:
If Trim(txtEmpID) = "" Then
    MsgBox "Login ID can't be blank! Please Enter Login ID.", vbExclamation, "Remark Message"
    txtEmpID.SetFocus
    Exit Sub
ElseIf Trim(txtPassword) = "" Then
    MsgBox "Password can't be blank! Please Enter valid password.", vbExclamation, "Remark Message"
    txtPassword.SetFocus
    Exit Sub
ElseIf optLogIn.Value = False And optLogOut.Value = False Then
    MsgBox "Please verify your Log In or Log Out option!", vbExclamation, "Remark Message"
    Exit Sub
End If
Dim sql As String
Dim rsCheck As New Recordset
Dim rsCheckAttend As New Recordset
Dim tmpRec As New Recordset
tmpRec.Open "Select vEmployee, vLogin, vLogOut, Password From tblEmpDetails Where vEmployee = '" & txtEmpID & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not tmpRec.EOF Then
    If UCase(txtEmpID) = UCase(tmpRec!vEmployee) And UCase(txtPassword) = UCase(tmpRec!Password) Then
        If optLogIn.Value = True Then
            rsCheck.Open "Select Getdate()", CnnMain, adOpenForwardOnly, adLockReadOnly
            rsCheckAttend.Open "Select * From tblAttend Where dtDate='" & Mid(Format(rsCheck(0), "mm/dd/yyyy"), 1, 11) & "' And vEmployee = '" & txtEmpID & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
            If Not rsCheckAttend.EOF Then
                'MsgBox "You can't Log In twice a day!", vbCritical, "Validation Message"
                CnnMain.BeginTrans
                CnnMain.Execute "Update tblAttend Set vRemark='" & Trim(cmbRemark.Text) & "' Where dtDate='" & Mid(Format(rsCheck(0), "mm/dd/yyyy"), 1, 11) & "' And vEmployee = '" & Trim(txtEmpID) & "'  And dtLogOut is Null"
                CnnMain.CommitTrans
                Set rsCheck = Nothing
                Set rsCheckAttend = Nothing
                Call cmdRefresh_Click
                MsgBox "Remark update successfull", vbInformation, "Remark Message"
                Exit Sub
            Else
                CnnMain.BeginTrans
                CnnMain.Execute "Exec Sp_LogIn_tblAttend '" & StrConv(Trim(txtEmpID), vbProperCase) & "', '" & tmpRec!vLogin & "', '" & tmpRec!vLogOut & "', '" & Trim(cmbRemark.Text) & "'"
                CnnMain.CommitTrans
                cmdRefresh_Click
                MsgBox "Log In Successfull.", vbInformation, "Remark Message"
            End If
        ElseIf optLogOut.Value = True Then
            rsCheck.Open "Select Getdate()", CnnMain, adOpenForwardOnly, adLockReadOnly
            rsCheckAttend.Open "Select * From tblAttend Where dtDate='" & Mid(Format(rsCheck(0), "mm/dd/yyyy"), 1, 11) & "' And vEmployee='" & Trim(txtEmpID) & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
            If Not rsCheckAttend.EOF Then
               If rsCheckAttend!dtLogOut <> "" Then
                    MsgBox "You can't Log Out twice a day!", vbCritical, "Validation Message"
                    Set rsCheck = Nothing
                    Set rsCheckAttend = Nothing
                    Exit Sub
                End If
            Else
                MsgBox "You have to Log In for Log Out!", vbCritical, "Validation Message"
                Set rsCheck = Nothing
                Set rsCheckAttend = Nothing
                Exit Sub
            End If
            CnnMain.BeginTrans
            CnnMain.Execute "Exec Sp_LogOut_tblAttend '" & StrConv(Trim(txtEmpID), vbProperCase) & "'"
            CnnMain.CommitTrans
            MsgBox "Log Out Successfull.", vbInformation, "Remark Message"
            '==
            Dim intCount As Integer
            intCount = 1
            Do While intCount <= ListView1.ListItems.Count
                Set ListView1.SelectedItem = ListView1.ListItems(intCount)
                If ListView1.SelectedItem.Text = Trim(txtEmpID) And ListView1.SelectedItem.ListSubItems(4) <> "" Then
                    ListView1.SelectedItem.ListSubItems(6) = Format(rsCheck(0), "HH:MM AMPM") & ""
                ElseIf ListView1.SelectedItem.Text = Trim(txtEmpID) And ListView1.SelectedItem.ListSubItems(4) = "" Then
                    ListView1.ListItems.Remove intCount
                End If
                intCount = intCount + 1
            Loop
            '==
        End If
        txtEmpID = ""
        txtPassword = ""
        cmbRemark.Text = ""
    Else
        MsgBox "Invalid Password! Please enter valid Password.", vbCritical, "Validation Message"
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword)
        txtPassword.SetFocus
        Exit Sub
    End If
Else
    MsgBox "Invalid Employee Name! Please enter valid Employee Name.", vbCritical, "Validation Message"
    txtEmpID.SelStart = 0
    txtEmpID.SelLength = Len(txtEmpID)
    txtEmpID.SetFocus
    Exit Sub
End If
Set rsCheck = Nothing
Set rsCheckAttend = Nothing
Set tmpRec = Nothing
'Call cmdRefresh_Click
Exit Sub
'ErrHandle:
'    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
'    Err.Clear
End Sub

Private Sub Command1_Click()
    frmLogin.Show
End Sub





Private Sub Form_Load()
'On Error GoTo ErrHandle:
Me.Move (Screen.Width - Me.Width) / 2, (frmGroup.ScaleHeight - Me.Height) / 2
'==Input Data for cmbRemark Combo Box
'--Initialize Option Table

Dim rsRemark As New Recordset
rsRemark.Open "Select vRemark From tblRemark Order By vRemark Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsRemark.EOF Then
    Do While Not rsRemark.EOF
        cmbRemark.AddItem rsRemark!vRemark & ""
        rsRemark.MoveNext
    Loop
End If
Set rsRemark = Nothing
'==Input Data for Group (ListView2)
ListView2.ColumnHeaders.Add , , "ID", 800
ListView2.ColumnHeaders.Add , , "Group Name", 2200
ListView2.ColumnHeaders.Item(ListView2.SortKey + 1).Icon = ListView2.SortOrder + 1
Dim rsAdd As New Recordset
rsAdd.Open "Select * From tblGroup Order By vGroupName Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
Dim itmX As ListItem
If Not rsAdd.EOF Then
    Do While Not rsAdd.EOF
        Set itmX = ListView2.ListItems.Add(, , rsAdd!siGroupID, , 3)
        itmX.SubItems(1) = rsAdd!vGroupName & ""
        rsAdd.MoveNext
    Loop
End If
Set rsAdd = Nothing
'==For Listview1
ListView1.ColumnHeaders.Add , , "Employee", 1500
ListView1.ColumnHeaders.Add , , "Designation", 1500
ListView1.ColumnHeaders.Add , , "Login Time", 1600
ListView1.ColumnHeaders.Add , , "Attend Time", 1600
ListView1.ColumnHeaders.Add , , "Remark", 1600
ListView1.ColumnHeaders.Add , , "Ac. Logout Time", 1600
ListView1.ColumnHeaders.Add , , "Logout Time", 1600
'Call SetHeaderFontStyle(ListView1) '---For Bold Column Header
ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
Exit Sub
'ErrHandle:
'    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
'    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
If Cancel = 0 Then
    Dim intResponse As Integer
    intResponse = MsgBox("Are you sure you want to exit the Application?", vbYesNo + vbQuestion, "Confirmation")
    If intResponse = vbYes Then
        End
    Else
        Cancel = 1
        Me.MousePointer = vbDefault
    End If
End If
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = 0
    Call Module3.lvSortByColumn(ListView1, ColumnHeader)
    ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
End Sub

Private Sub ListView1_DblClick()
'On Error GoTo ErrHandle:
If ListView1.ListItems.Count > 0 Then
    txtEmpID = ListView1.SelectedItem.Text
    cmbRemark.Text = ListView1.SelectedItem.ListSubItems(4)
    txtPassword = ""
    txtPassword.SetFocus
End If
Exit Sub
'ErrHandle:
'    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
'    Err.Clear
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
'On Error GoTo ErrHandle:
If KeyAscii = 13 Then
    If ListView1.ListItems.Count > 0 Then
        txtEmpID = ListView1.SelectedItem.Text
        cmbRemark.Text = ListView1.SelectedItem.ListSubItems(4)
        txtPassword = ""
        txtPassword.SetFocus
    End If
End If
Exit Sub
'ErrHandle:
'    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
'    Err.Clear
End Sub
Private Sub txtEmpID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtPassword.SetFocus
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSubmit.SetFocus
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.ColumnHeaders.Item(ListView2.SortKey + 1).Icon = 0
    Call Module3.lvSortByColumn(ListView2, ColumnHeader)
    ListView2.ColumnHeaders.Item(ListView2.SortKey + 1).Icon = ListView2.SortOrder + 1
End Sub
Private Sub cmbRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSubmit.SetFocus
End Sub
