VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Information"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmGroup.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroup.frx":116E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "List of Group Information"
      ForeColor       =   &H00000000&
      Height          =   4320
      Left            =   127
      TabIndex        =   13
      Top             =   2565
      Width           =   5820
      Begin MSComctlLib.ListView ListView1 
         Height          =   3930
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   6932
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
   Begin VB.Frame Frame2 
      Caption         =   "Navigation"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   127
      TabIndex        =   12
      Top             =   1485
      Width           =   5820
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   135
         MaskColor       =   &H8000000F&
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
         MaskColor       =   &H8000000F&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3465
         MaskColor       =   &H8000000F&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   2355
         MaskColor       =   &H8000000F&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1245
         MaskColor       =   &H8000000F&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   405
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group Information"
      ForeColor       =   &H00000000&
      Height          =   1290
      Left            =   127
      TabIndex        =   9
      Top             =   120
      Width           =   5820
      Begin VB.CommandButton cmdFindGroup 
         Height          =   300
         Left            =   2205
         Picture         =   "frmGroup.frx":1A48
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "PK"
         ToolTipText     =   "Search Screen"
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtGroupId 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   1
         ToolTipText     =   "Maximum 3 Characters"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtGroupName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1215
         MaxLength       =   25
         TabIndex        =   2
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   750
         Width           =   4410
      End
      Begin VB.Label lblAreaId 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Group ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblAreaName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Name:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   750
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
On Error GoTo ErrHandle:
Dim Control As Object
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
Call prc_DisAllow
txtGroupID.Enabled = True
txtGroupID.BackColor = &HFFFFFF
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdClose_Click()
On Error GoTo ErrHandle:
    Unload Me
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdFindGroup_Click()
On Error GoTo ErrHandle:
    frmSearchGroup.Show vbModal
If strfindcode <> "" Then
    txtGroupID = strfindcode
    txtGroupID_Validate True
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
Dim Control As Object
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
txtGroupID.Enabled = False
txtGroupID.BackColor = &H80000018
Call prc_Allow
txtGroupName.SetFocus
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandle:
If Trim(txtGroupName) = "" Then
    MsgBox "Group Name can't be blank! Please Enter Group Name.", vbExclamation, "Remark Message"
    txtGroupName.SetFocus
    Exit Sub
End If
'--Name Already Exist
Dim rsName As New Recordset
rsName.Open "Select vGroupName from tblGroup Where vGroupName='" & txtGroupName & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsName.EOF Then
    MsgBox "Group name already exist! Group Name should not be duplicate!", vbExclamation, "Remark Message"
    txtGroupName.SetFocus
 Exit Sub
End If
Set rsName = Nothing
'------sp_Add_CBtblZone
Me.MousePointer = vbHourglass
CnnMain.BeginTrans
CnnMain.Execute "Exec Sp_Add_tblGroup '" & Trim(txtGroupName) & "'"
CnnMain.CommitTrans
Dim rsGroupID As New Recordset
rsGroupID.Open "Select siGroupID from tblGroup Order By siGroupID Desc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsGroupID.EOF Then
    txtGroupID = rsGroupID!siGroupID & ""
    CnnMain.BeginTrans
    CnnMain.Execute "Exec Sp_Add_tblObject '" & Trim(txtGroupID) & "','" & "Group_" & Trim(txtGroupName) & "','No'"
    CnnMain.CommitTrans
End If
Set rsGroupID = Nothing
Call prc_DisAllow
Me.MousePointer = vbDefault
MsgBox "Record successfully saved to the database.", vbInformation, "Remark Message"
Call cmdDataSearch_Click
txtGroupID.Enabled = True
txtGroupID.BackColor = &HFFFFFF
txtGroupName.SelStart = 0
txtGroupName.SelLength = Len(txtGroupName)
txtGroupName.SetFocus
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
On Error GoTo ErrHandle:
    Call prc_DisAllow
    txtGroupID = ListView1.SelectedItem.Text
    txtGroupID_Validate True
    Exit Sub
ErrHandle:
        MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
        Err.Clear
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle:
If KeyAscii = 13 Then
    Call prc_DisAllow
    txtGroupID = ListView1.SelectedItem.Text
    txtGroupID_Validate True
End If
Exit Sub
ErrHandle:
        MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
        Err.Clear
End Sub

Private Sub txtGroupID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 12 Or KeyAscii > 14 And KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtGroupID_Validate False
End Sub

Private Sub txtGroupID_Validate(Cancel As Boolean)
On Error GoTo ErrHandle:
If txtGroupID = "" Then
    cmdNew_Click
    Exit Sub
End If
Dim rsGroup As New ADODB.Recordset
rsGroup.Open "Select vGroupName From tblGroup Where siGroupId='" & Trim(txtGroupID) & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsGroup.EOF Then
    With rsGroup
        txtGroupName = !vGroupName & ""
    End With
    txtGroupName.SelStart = 0
    txtGroupName.SelLength = Len(txtGroupName)
    txtGroupName.SetFocus
    prc_DisAllow
    cmdUpdate.Enabled = True
Else
    MsgBox "Invalid Group ID!", vbExclamation, "Validation Message"
    Cancel = True
    txtGroupID.SelStart = 0
    txtGroupID.SelLength = Len(txtGroupID)
    txtGroupName = ""
End If
Set rsGroup = Nothing
If txtGroupID.Enabled = False Then
    txtGroupID.Enabled = True
    txtGroupID.BackColor = &HFFFFFF
End If
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrHandle:
If Trim(txtGroupID.Text) = "" Then
    MsgBox "Group ID can't be blank! Please Enter Group Id.", vbExclamation, "Remark Message"
    txtGroupID.SetFocus
    Exit Sub
End If
If Trim(txtGroupName.Text) = "" Then
    MsgBox "Group Name can't be blank! Please Enter Group Name.", vbExclamation, "Remark Message"
    txtGroupName.SetFocus
    Exit Sub
End If
'--SP_Up_CBtblZone-------
CnnMain.BeginTrans
CnnMain.Execute "Exec Sp_Up_tblGroup " & Trim(txtGroupID) & ",'" & Trim(txtGroupName) & "'"
CnnMain.Execute "Update tblObject Set vObjectCaption='" & "Group_" & Trim(txtGroupName) & "' Where vObjectName='" & txtGroupID & "'"
CnnMain.CommitTrans
Me.MousePointer = vbDefault
MsgBox "Record successfully update to the database.", vbInformation, "Update Message"
Call cmdDataSearch_Click
txtGroupName.SelStart = 0
txtGroupName.SelLength = Len(txtGroupName)
txtGroupName.SetFocus
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub Form_Load()
On Error GoTo ErrHandle:
'Me.Move (Screen.Width - Me.Width) / 2, (mdiSwitchBoard.ScaleHeight - Me.Height) / 2
Call prc_DisAllow
ListView1.ColumnHeaders.Add , , "Group ID", 1500
ListView1.ColumnHeaders.Add , , "Group Name", 4000
'Call SetHeaderFontStyle(ListView1) '---For Bold Column Header
ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
Call cmdDataSearch_Click
Me.MousePointer = vbDefault
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
            If obj.Name = "cmdSave" Or obj.Name = "cmdCancel" Or obj.Name = "cmdFindGroup" Then
                obj.Enabled = True
            Else
                obj.Enabled = False
            End If
        End If
    Next
End Sub
Public Sub clear_all_textboxes()
    txtGroupID = ""
    txtGroupName = ""
End Sub
Private Sub txtGroupName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cmdDataSearch_Click()
On Error GoTo ErrHandle:
Dim rsAdd As New Recordset
rsAdd.Open "Select * From tblGroup Order By vGroupName Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
Dim itmX As ListItem
ListView1.ListItems.Clear
If Not rsAdd.EOF Then
    Do While Not rsAdd.EOF
        Set itmX = ListView1.ListItems.Add(, , rsAdd!siGroupID, , 3)
        itmX.SubItems(1) = rsAdd!vGroupName & ""
        rsAdd.MoveNext
    Loop
End If
Set rsAdd = Nothing
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

