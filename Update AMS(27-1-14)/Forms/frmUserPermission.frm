VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmUserPermission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Object Permission"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "frmUserPermission.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEnableAllMenu 
      Caption         =   "Enable All Object Permission"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2880
      TabIndex        =   5
      Top             =   6810
      Width           =   3255
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit Permission"
      Height          =   375
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   6600
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
            Picture         =   "frmUserPermission.frx":08CA
            Key             =   "Group"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPermission.frx":11A4
            Key             =   "User"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPermission.frx":1A7E
            Key             =   "Expand"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      ForeColor       =   &H00000000&
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2625
      Begin MSComctlLib.TreeView TV 
         Height          =   6105
         Left            =   120
         TabIndex        =   3
         Top             =   255
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   10769
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Object Information"
      ForeColor       =   &H00000000&
      Height          =   6495
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   6810
      Begin MSComctlLib.ListView ListView1 
         Height          =   6135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   10821
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   6600
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
            Picture         =   "frmUserPermission.frx":2958
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPermission.frx":2DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPermission.frx":31FC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUserPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LoginName As String

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
    Screen.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdSubmit_Click()
On Error GoTo ErrHandle:
'--Update Menu Permission
Me.MousePointer = vbHourglass
CnnMain.BeginTrans
CnnMain.Execute "Delete From tblObjectPermission Where vEmployee='" & LoginName & "'"
CnnMain.CommitTrans
Dim intCount As Integer
intCount = 1
    Do While intCount <= ListView1.ListItems.Count
        If ListView1.ListItems.Item(intCount).Checked = True Then
            Set ListView1.SelectedItem = ListView1.ListItems(intCount)
            CnnMain.BeginTrans
            CnnMain.Execute "Exec Sp_Up_tblObjectPermission '" & LoginName & "','" & ListView1.ListItems.Item(intCount).SubItems(1) & "','" & ListView1.ListItems.Item(intCount).SubItems(2) & "'"
            CnnMain.CommitTrans
        Else
            CnnMain.BeginTrans
            CnnMain.Execute "Delete From tblObjectPermission Where vEmployee='" & LoginName & "' And vObjectName='" & ListView1.ListItems.Item(intCount).SubItems(1) & "'"
            CnnMain.CommitTrans
        End If
        intCount = intCount + 1
    Loop
Me.MousePointer = vbDefault
MsgBox "Object Permission successfully updated to the database.", vbInformation, "Remark Message"
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
ListView1.ColumnHeaders.Add , , "Object Caption", 3000, 0
ListView1.ColumnHeaders.Add , , "Object Name", 2100, 0
ListView1.ColumnHeaders.Add , , "Enable", 1370, 0
'Call Module3.SetHeaderFontStyle(ListView1) '---For Bold Column Header
ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1

Dim nodTV As Node
Dim rsGroup As New Recordset
Dim rsEmp As New Recordset
rsGroup.Open "Select siGroupID,vGroupName From tblGroup", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsGroup.EOF Then
    Do While Not rsGroup.EOF
        Set nodTV = TV.Nodes.Add(, , "'" & rsGroup!siGroupID & "'", "" & rsGroup!vGroupName & "", "Group")
        nodTV.ExpandedImage = "Expand"
        rsEmp.Open "Select vEmployee From tblEmpDetails Order By vEmployee Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
        If Not rsEmp.EOF Then
            Dim rsFindDept As New Recordset
            Dim rsEmpGroup As New Recordset
            Do While Not rsEmp.EOF
                rsFindDept.Open "Select siDeptID From tblEmpDetails Where vEmployee='" & rsEmp!vEmployee & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
                rsEmpGroup.Open "Select siGroupID From tblDepartment Where siDeptID=" & rsFindDept!siDeptID & "", CnnMain, adOpenForwardOnly
                If rsGroup!siGroupID = rsEmpGroup!siGroupID Then
                    Set nodTV = TV.Nodes.Add("'" & rsGroup!siGroupID & "'", tvwChild, , "" & rsEmp!vEmployee & "", "User")
                End If
                rsEmp.MoveNext
                rsFindDept.Close
                rsEmpGroup.Close
            Loop
        End If
        rsEmp.Close
        rsGroup.MoveNext
    Loop
End If
Set rsGroup = Nothing
Set rsEmp = Nothing
Set rsFindDept = Nothing
Set rsEmpGroup = Nothing
Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
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
Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ErrHandle:
    Me.MousePointer = vbHourglass
    chkEnableAllMenu.Value = 0
    LoginName = TV.SelectedItem.Text
    ListView1.ListItems.Clear
    Dim rsEmp As New Recordset
    rsEmp.Open "Select vEmployee From tblEmpDetails Where vEmployee='" & TV.SelectedItem.Text & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsEmp.EOF Then
        Dim rsMenu As New Recordset
        'rsMenu.Open "Select * From tblObjectPermission Where vEmpID='" & TV.SelectedItem.Text & "' Order By vObjectName Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
        rsMenu.Open "Select * From tblObject Order By vObjectName Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
        If Not rsMenu.EOF Then
            Frame4.Caption = "Object Permission For " & LoginName
            Call prc_DataShowFormat(rsMenu)
        End If
    Else
        ListView1.ListItems.Clear
        Frame4.Caption = "Object Permission"
    End If
    Set rsMenu = Nothing
   Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Public Sub prc_DataShowFormat(rsMenu As Recordset)
On Error GoTo ErrHandle:
Dim itmX As ListItem
ListView1.ListItems.Clear
Dim intCount As Integer
intCount = 1
If Not rsMenu.EOF Then
    Do While Not rsMenu.EOF
        Set itmX = ListView1.ListItems.Add(, , rsMenu!vObjectCaption, , 3)
        itmX.SubItems(1) = rsMenu!vObjectName & ""
        'itmX.SubItems(2) = rsMenu!vDefault & ""
        If rsMenu!vDefault = "Yes" Then
            itmX.SubItems(2) = rsMenu!vDefault & ""
            ListView1.ListItems.Item(intCount).Checked = True
        Else
            itmX.SubItems(2) = rsMenu!vDefault & ""
            ListView1.ListItems.Item(intCount).Checked = False
        End If
        intCount = intCount + 1
        rsMenu.MoveNext
    Loop
    '==
    Dim rsCheck As New Recordset
    Dim intCheck As Integer
    intCheck = 1
    Do While intCheck <= ListView1.ListItems.Count
        Set ListView1.SelectedItem = ListView1.ListItems(intCheck)
        rsCheck.Open "Select * From tblObjectPermission Where vObjectName='" & ListView1.SelectedItem.ListSubItems(1) & "' And   vEmployee='" & TV.SelectedItem.Text & "' Order By vObjectName Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
        If Not rsCheck.EOF Then
            If rsCheck!vEnable = "Yes" Then
                ListView1.ListItems.Item(intCheck).Checked = True
                ListView1.SelectedItem.ListSubItems(2) = "Yes"
            Else
                ListView1.ListItems.Item(intCheck).Checked = False
                ListView1.SelectedItem.ListSubItems(2) = "No"
            End If
        End If
        rsCheck.Close
        intCheck = intCheck + 1
    Loop
End If
Set rsCheck = Nothing
Set rsMenu = Nothing
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

