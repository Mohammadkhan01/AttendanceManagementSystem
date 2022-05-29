VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSearchMenuInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Screen"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
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
      Height          =   6375
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   6585
      Begin MSComctlLib.ListView ListView1 
         Height          =   6015
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   10610
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   7815
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
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
      Left            =   6480
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1170
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmSearchMenuInfo.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchMenuInfo.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchMenuInfo.frx":08A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "IMART SOFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   180
      TabIndex        =   4
      Top             =   6000
      Width           =   2235
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6315
      Left            =   120
      Picture         =   "frmSearchMenuInfo.frx":117E
      Stretch         =   -1  'True
      Top             =   180
      Width           =   2325
   End
End
Attribute VB_Name = "frmSearchMenuInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
On Error GoTo ErrHandle:
Unload Me
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle:
strfindcode = ListView1.SelectedItem.SubItems(1)
Unload Me
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2

ListView1.ColumnHeaders.Add , , "Object Caption", 3000, 0
ListView1.ColumnHeaders.Add , , "Object Name", 2050, 0
ListView1.ColumnHeaders.Add , , "Enable", 1200, 0
'Call Module3.SetHeaderFontStyle(ListView1) '---For Bold Column Header
ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1

Dim rsMenu As New Recordset
rsMenu.Open "Select * From tblObject  Order By vObjectCaption Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsMenu.EOF Then
    Call prc_DataShowFormat(rsMenu)
Else
    ListView1.ListItems.Clear
    MsgBox "You have to input one or more Menu Information", vbCritical, "Remark Message"
End If
Set rsMenu = Nothing
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
strfindcode = ListView1.SelectedItem.SubItems(1)
Unload Me
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle:
    If KeyAscii = 13 Then
        strfindcode = ListView1.SelectedItem.Text
        Unload Me
    End If
    Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Public Sub prc_DataShowFormat(rsComputer As Recordset)
On Error GoTo ErrHandle:
Dim itmX As ListItem
ListView1.ListItems.Clear
If Not rsComputer.EOF Then
    Do While Not rsComputer.EOF
        Set itmX = ListView1.ListItems.Add(, , rsComputer(1), , 3)
        itmX.SubItems(1) = rsComputer(0) & ""
        itmX.SubItems(2) = rsComputer(2) & ""
        rsComputer.MoveNext
    Loop
        Frame1.Caption = "Total Computer: " & ListView1.ListItems.Count
Else
    MsgBox "Computer Information Not Found! You have to input one or more Computer Information!", vbExclamation, "Validation Message"
    Exit Sub
End If
Set rsComputer = Nothing
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
