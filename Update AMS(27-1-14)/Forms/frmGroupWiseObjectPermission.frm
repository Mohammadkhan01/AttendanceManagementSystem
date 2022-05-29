VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGroupWiseObjectPermission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Wise Object Permission"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   Icon            =   "frmGroupWiseObjectPermission.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11460
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEnableAllMenu 
      Appearance      =   0  'Flat
      Caption         =   "Enable All Object Permission"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3795
      TabIndex        =   10
      Top             =   7155
      Width           =   3255
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit Permission"
      Height          =   375
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7065
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "Object Information"
      ForeColor       =   &H00000000&
      Height          =   6735
      Left            =   3795
      TabIndex        =   7
      Top             =   120
      Width           =   7530
      Begin MSComctlLib.ListView ListView1 
         Height          =   6375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   11245
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ColHdrIcons     =   "ImageList2"
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
   Begin VB.Frame Frame2 
      Caption         =   "Group Criteria"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3540
      Begin VB.ComboBox cmbRole 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmGroupWiseObjectPermission.frx":08CA
         Left            =   1695
         List            =   "frmGroupWiseObjectPermission.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   330
         Width           =   1725
      End
      Begin VB.OptionButton optIndvGroup 
         Appearance      =   0  'Flat
         Caption         =   "Inidividual Group"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton optAllGroup 
         Appearance      =   0  'Flat
         Caption         =   "All Group"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Object Criteria"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3540
      Begin VB.OptionButton optRemoveObject 
         Appearance      =   0  'Flat
         Caption         =   "Remove Object"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   630
         Width           =   1575
      End
      Begin VB.OptionButton optReplaceObject 
         Appearance      =   0  'Flat
         Caption         =   "Replace Object"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   930
         Width           =   1695
      End
      Begin VB.OptionButton optAddObject 
         Appearance      =   0  'Flat
         Caption         =   "Add Object"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupWiseObjectPermission.frx":08CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupWiseObjectPermission.frx":0D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupWiseObjectPermission.frx":1172
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4440
      Left            =   120
      Picture         =   "frmGroupWiseObjectPermission.frx":1A4C
      Stretch         =   -1  'True
      Top             =   2895
      Width           =   3540
   End
End
Attribute VB_Name = "frmGroupWiseObjectPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkEnableAllMenu_Click()
On Error GoTo ErrHandle:
Dim iCount As Integer
iCount = 1
If ListView1.ListItems.Count > 0 Then
    Do While iCount <= ListView1.ListItems.Count
        If chkEnableAllMenu.Value = 1 Then
            ListView1.ListItems(iCount).Checked = True
            ListView1.ListItems.Item(iCount).SubItems(2) = "Yes"
        Else
            ListView1.ListItems(iCount).Checked = False
            ListView1.ListItems.Item(iCount).SubItems(2) = "No"
        End If
        iCount = iCount + 1
    Loop
End If
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmbRole_Click()

Dim rsChange As New ADODB.Recordset
rsChange.Open "Select siDeptId from tblDepartment where vDeptName='" & cmbRole.Text & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsChange.EOF Then
        Label1.Caption = rsChange!siDeptId
    Else
    End If
Set rsChange = Nothing

End Sub

Private Sub cmdSubmit_Click()
On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
Dim intCount As Integer
Dim rsEmp As New Recordset
If optIndvGroup.Value = True And optAddObject.Value = True Then
    intCount = 1
    rsEmp.Open "Select vEmployee From tblEmpDetails Where siDeptId =" & Label1.Caption & " Order By vEmployee Asc ", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsEmp.EOF Then
        Do While Not rsEmp.EOF
            Do While intCount <= ListView1.ListItems.Count
                If ListView1.ListItems.Item(intCount).Checked = True Then
                    Set ListView1.SelectedItem = ListView1.ListItems(intCount)
                    CnnMain.BeginTrans
                    CnnMain.Execute "Delete From tblObjectPermission Where vEmployee='" & rsEmp!vEmployee & "' And vObjectName='" & ListView1.ListItems.Item(intCount).SubItems(1) & "'"
                    CnnMain.Execute "Exec Sp_Add_tblObjectPermission '" & rsEmp!vEmployee & "','" & ListView1.ListItems.Item(intCount).SubItems(1) & "','" & ListView1.ListItems.Item(intCount).SubItems(2) & "'"
                    CnnMain.CommitTrans
                End If
                intCount = intCount + 1
            Loop
            intCount = 1
            rsEmp.MoveNext
        Loop
    End If
ElseIf optAllGroup.Value = True And optAddObject.Value = True Then
    intCount = 1
    rsEmp.Open "Select vEmployee From tblEmpDetails Order By vEmployee Asc ", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsEmp.EOF Then
        Do While Not rsEmp.EOF
            Do While intCount <= ListView1.ListItems.Count
                If ListView1.ListItems.Item(intCount).Checked = True Then
                    Set ListView1.SelectedItem = ListView1.ListItems(intCount)
                    CnnMain.BeginTrans
                    CnnMain.Execute "Delete From tblObjectPermission Where vEmployee='" & rsEmp!vEmployee & "' And vObjectName='" & ListView1.ListItems.Item(intCount).SubItems(1) & "'"
                    CnnMain.Execute "Exec Sp_Add_tblObjectPermission '" & rsEmp!vEmployee & "','" & ListView1.ListItems.Item(intCount).SubItems(1) & "','" & ListView1.ListItems.Item(intCount).SubItems(2) & "'"
                    CnnMain.CommitTrans
                End If
                intCount = intCount + 1
            Loop
            intCount = 1
            rsEmp.MoveNext
        Loop
    End If
ElseIf optIndvGroup.Value = True And optRemoveObject.Value = True Then
    intCount = 1
    rsEmp.Open "Select vEmployee From tblEmpDetails Where siDeptId=" & Trim(Label1.Caption) & " Order By vEmployee Asc ", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsEmp.EOF Then
        Do While Not rsEmp.EOF
            Do While intCount <= ListView1.ListItems.Count
                If ListView1.ListItems.Item(intCount).Checked = True Then
                    Set ListView1.SelectedItem = ListView1.ListItems(intCount)
                    CnnMain.BeginTrans
                    CnnMain.Execute "Delete From tblObjectPermission Where vEmployee='" & rsEmp!vEmployee & "' And vObjectName='" & ListView1.ListItems.Item(intCount).SubItems(1) & "'"
                    CnnMain.CommitTrans
                End If
                intCount = intCount + 1
            Loop
            intCount = 1
            rsEmp.MoveNext
        Loop
    End If
ElseIf optAllGroup.Value = True And optRemoveObject.Value = True Then
    intCount = 1
    Do While intCount <= ListView1.ListItems.Count
        If ListView1.ListItems.Item(intCount).Checked = True Then
            Set ListView1.SelectedItem = ListView1.ListItems(intCount)
            CnnMain.BeginTrans
            CnnMain.Execute "Delete From tblObjectPermission Where vObjectName='" & ListView1.ListItems.Item(intCount).SubItems(1) & "'"
            CnnMain.CommitTrans
        End If
        intCount = intCount + 1
    Loop
ElseIf optIndvGroup.Value = True And optReplaceObject.Value = True Then
    intCount = 1
    rsEmp.Open "Select vEmployee From tblEmpDetails Where siDeptId=" & Trim(Label1.Caption) & " Order By vEmployee Asc ", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsEmp.EOF Then
        Do While Not rsEmp.EOF
            CnnMain.BeginTrans
            CnnMain.Execute "Delete From tblObjectPermission Where vEmployee='" & rsEmp!vEmployee & "'"
            CnnMain.CommitTrans
            Do While intCount <= ListView1.ListItems.Count
                If ListView1.ListItems.Item(intCount).Checked = True Then
                    Set ListView1.SelectedItem = ListView1.ListItems(intCount)
                    CnnMain.BeginTrans
                    CnnMain.Execute "Exec Sp_Add_tblObjectPermission '" & rsEmp!vEmployee & "','" & ListView1.ListItems.Item(intCount).SubItems(1) & "','" & ListView1.ListItems.Item(intCount).SubItems(2) & "'"
                    CnnMain.CommitTrans
                End If
                intCount = intCount + 1
            Loop
            intCount = 1
            rsEmp.MoveNext
        Loop
    End If
ElseIf optAllGroup.Value = True And optReplaceObject.Value = True Then
    intCount = 1
    rsEmp.Open "Select vEmployee From tblempDetails Order By vEmployee Asc ", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsEmp.EOF Then
        Do While Not rsEmp.EOF
            CnnMain.BeginTrans
            CnnMain.Execute "Delete From tblObjectPermission Where vEmployee='" & rsEmp!vEmployee & "'"
            CnnMain.CommitTrans
            Do While intCount <= ListView1.ListItems.Count
                If ListView1.ListItems.Item(intCount).Checked = True Then
                    Set ListView1.SelectedItem = ListView1.ListItems(intCount)
                    CnnMain.BeginTrans
                    CnnMain.Execute "Exec Sp_Add_tblObjectPermission '" & rsEmp!vEmployee & "','" & ListView1.ListItems.Item(intCount).SubItems(1) & "','" & ListView1.ListItems.Item(intCount).SubItems(2) & "'"
                    CnnMain.CommitTrans
                End If
                intCount = intCount + 1
            Loop
            intCount = 1
            rsEmp.MoveNext
        Loop
    End If
End If
Me.MousePointer = vbDefault
Set rsEmp = Nothing
MsgBox "Group Wise Object Permission successfully updated to the database.", vbInformation, "Remark Message"
Exit Sub
ErrHandle:
   Me.MousePointer = vbDefault
   MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
   Err.Clear
End Sub

Private Sub Form_Load()
'On Error GoTo ErrHandle:

Me.MousePointer = vbHourglass
Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
ListView1.ColumnHeaders.Add , , "Object Caption", 3500, 0
ListView1.ColumnHeaders.Add , , "Object Name", 2300, 0
ListView1.ColumnHeaders.Add , , "Enable", 1200, 0
'Call Module3.SetHeaderFontStyle(ListView1) '---For Bold Column Header
ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
'--Data Initialize For Listview1
Dim itmX As ListItem
ListView1.ListItems.Clear
Dim rsRole As New ADODB.Recordset
rsRole.Open "Select vDeptName from tblDepartment order by vDeptName asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsRole.EOF Then
    Do While Not rsRole.EOF
        cmbRole.AddItem rsRole!vDeptName
        rsRole.MoveNext
    Loop
End If
cmbRole.ListIndex = 0
Dim rsObject As New Recordset
rsObject.Open "Select * From tblObject Order By vObjectCaption Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsObject.EOF Then
    Do While Not rsObject.EOF
        Set itmX = ListView1.ListItems.Add(, , rsObject!vObjectCaption, , 3)
        itmX.SubItems(1) = rsObject!vObjectName & ""
        itmX.SubItems(2) = rsObject!vDefault & ""
        rsObject.MoveNext
    Loop
Else
    Set rsObject = Nothing
    MsgBox "Object Information Not Found! You have to input one or more Object Information!", vbExclamation, "Validation Message"
    Exit Sub
End If
Set rsObject = Nothing
Me.MousePointer = vbDefault
Exit Sub
'ErrHandle:
 '   Me.MousePointer = vbDefault
  '  MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
   ' Err.Clear
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.Sorted = True
    ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = 0
    Call Module3.lvSortByColumn(ListView1, ColumnHeader)
    ListView1.Sorted = False
    ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
Dim intCount As Integer
For intCount = 1 To ListView1.ListItems.Count
 If ListView1.ListItems(intCount).Checked = True Then
    ListView1.ListItems.Item(intCount).SubItems(2) = "Yes"
   Else
      ListView1.ListItems.Item(intCount).SubItems(2) = "No"
 End If
Next
Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

