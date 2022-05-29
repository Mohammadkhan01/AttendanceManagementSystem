VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmEntryRemark 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remark Information"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmEntryRemark.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Navigation"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   135
      TabIndex        =   8
      Top             =   1440
      Width           =   5820
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1245
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   405
         Width           =   1095
      End
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remark Information"
      ForeColor       =   &H00000000&
      Height          =   1245
      Left            =   135
      TabIndex        =   10
      Top             =   90
      Width           =   5820
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1275
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "Maximum 100 Characters"
         Top             =   735
         Width           =   4365
      End
      Begin VB.TextBox txtRemarkID 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1275
         TabIndex        =   1
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblAreaId 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Remark ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         TabIndex        =   12
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblAreaName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Remark:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   615
         TabIndex        =   11
         Top             =   720
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List of Remark Information"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   135
      TabIndex        =   9
      Top             =   2520
      Width           =   5820
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   4895
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntryRemark.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntryRemark.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntryRemark.frx":116E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEntryRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
On Error GoTo ErrHandle:
txtRemarkID = ""
txtRemark = ""
Call prc_DisAllow
txtRemarkID.Enabled = True
txtRemarkID.BackColor = &HFFFFFF
Exit Sub
ErrHandle:
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
Private Sub cmdNew_Click()
On Error GoTo ErrHandle:
    txtRemarkID = ""
    txtRemark = ""
    txtRemarkID.Enabled = False
    txtRemarkID.BackColor = &HC0FFFF
    txtRemark.SetFocus
    Call prc_Allow
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandle:
If Trim(txtRemark) = "" Then
    MsgBox "Earn type can't be blank! Please enter Earn type!", vbExclamation, "Remark Message"
    txtRemark.SetFocus
    Exit Sub
End If
'---Data Save
Me.MousePointer = vbHourglass
CnnMain.BeginTrans
CnnMain.Execute "Sp_Add_tblRemark '" & txtRemark & "'"
CnnMain.CommitTrans
Dim rsEarnID As New Recordset
rsEarnID.Open "Select iAutoID from tblRemark Order by iAutoID Desc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEarnID.EOF Then
    txtRemarkID = rsEarnID!iAutoID & ""
End If
Set rsEarnID = Nothing
Call prc_DisAllow
txtRemarkID.Enabled = True
txtRemarkID.SelStart = Len(txtRemarkID)
txtRemarkID.SetFocus
txtRemarkID.BackColor = &HFFFFFF
Me.MousePointer = vbDefault
MsgBox "Remark successfully saved to the database", vbInformation, "Remark Message"
Call cmdDataSearch_Click
ListView1.SetFocus
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    CnnMain.RollbackTrans
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrHandle:
If Trim(txtRemarkID) = "" Then
    MsgBox "Remark ID can't be blank!  Please enter Remark ID!", vbExclamation, "Remark Message"
    txtRemarkID.SetFocus
    Exit Sub
ElseIf Trim(txtRemark) = "" Then
    MsgBox "Remark can't be blank!  Please enter Remark!", vbExclamation, "Remark Message"
    txtRemark.SetFocus
    Exit Sub
End If
'--Data Update
Me.MousePointer = vbHourglass
CnnMain.BeginTrans
CnnMain.Execute "Update tblRemark Set vRemark='" & txtRemark & "' Where iAutoID='" & txtRemarkID & "'"
CnnMain.CommitTrans
MsgBox "Record successfully updated to the database.", vbInformation, "Update Message"
Call cmdDataSearch_Click
Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    CnnMain.RollbackTrans
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
Me.MousePointer = vbHourglass
Call prc_DisAllow
ListView1.ColumnHeaders.Add , , "ID", 1000
ListView1.ColumnHeaders.Add , , "Remark", 4500
'Call SetHeaderFontStyle(ListView1) '---For Bold Column Header
ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
Dim rsAdd As New Recordset
rsAdd.Open "Select iAutoID, vRemark from tblRemark", CnnMain, adOpenForwardOnly, adLockReadOnly
Dim itmX As ListItem
If Not rsAdd.EOF Then
    Do While Not rsAdd.EOF
        Set itmX = ListView1.ListItems.Add(, , rsAdd!iAutoID, , 3)
        itmX.SubItems(1) = rsAdd!vRemark & ""
        rsAdd.MoveNext
    Loop
End If
Set rsAdd = Nothing
Me.MousePointer = vbDefault
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
    txtRemarkID = ListView1.SelectedItem.Text
    txtRemarkID_Validate True
    Exit Sub
ErrHandle:
        MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
        Err.Clear
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle:
If KeyAscii = 13 Then
    Call prc_DisAllow
    txtRemarkID = ListView1.SelectedItem.Text
    txtRemarkID_Validate True
End If
Exit Sub
ErrHandle:
        MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
        Err.Clear
End Sub

Private Sub txtRemarkID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 12 Or KeyAscii > 14 And KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtRemarkID_Validate True
End Sub

Private Sub txtRemarkID_Validate(Cancel As Boolean)
On Error GoTo ErrHandle:
If txtRemarkID = "" Then
    cmdNew_Click
    Exit Sub
End If
Me.MousePointer = vbHourglass
Dim rsAdd As New Recordset
rsAdd.Open "Select iAutoID, vRemark from tblRemark Where iAutoID='" & txtRemarkID & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsAdd.EOF Then
    txtRemark = rsAdd!vRemark & ""
    txtRemarkID.BackColor = &HFFFFFF
    txtRemarkID.Enabled = True
    cmdUpdate.Enabled = True
    txtRemark.SelStart = 0
    txtRemark.SelLength = Len(txtRemark)
    txtRemark.SetFocus
Else
    Me.MousePointer = vbDefault
    MsgBox "Invalid Earn ID!", vbExclamation, "Validation Message"
    Cancel = True
    txtRemarkID.SelStart = 0
    txtRemarkID.SelLength = Len(txtRemarkID)
    txtRemark = ""
End If
Set rsAdd = Nothing
Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdDataSearch_Click()
On Error GoTo ErrHandle:
Dim rsAdd As New Recordset
rsAdd.Open "Select iAutoID, vRemark from tblRemark", CnnMain, adOpenForwardOnly, adLockReadOnly
Dim itmX As ListItem
ListView1.ListItems.Clear
If Not rsAdd.EOF Then
    Do While Not rsAdd.EOF
        Set itmX = ListView1.ListItems.Add(, , rsAdd!iAutoID, , 3)
        itmX.SubItems(1) = rsAdd!vRemark & ""
        rsAdd.MoveNext
    Loop
End If
Set rsAdd = Nothing
ListView1.SetFocus
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
            If obj.Name = "cmdSave" Or obj.Name = "cmdCancel" Or obj.Name = "cmdBrowse" Or obj.Name = "cmdEarnDate" Then
                obj.Enabled = True
            Else
                obj.Enabled = False
            End If
        End If
    Next
End Sub
