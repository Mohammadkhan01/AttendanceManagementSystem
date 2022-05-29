VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSearchGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Screen"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
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
   ScaleHeight     =   7260
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearch 
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
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   2325
   End
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
      TabIndex        =   4
      Top             =   120
      Width           =   5835
      Begin MSComctlLib.ListView ListView1 
         Height          =   6000
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   10583
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
      Left            =   7170
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   5760
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1290
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
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
            Picture         =   "frmSearchGroup.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchGroup.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchGroup.frx":08A4
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
      Left            =   150
      TabIndex        =   6
      Top             =   6000
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Search By Group Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1680
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5685
      Left            =   120
      Picture         =   "frmSearchGroup.frx":117E
      Stretch         =   -1  'True
      Top             =   810
      Width           =   2325
   End
End
Attribute VB_Name = "frmSearchGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
If ListView1.ListItems.Count > 0 Then
    strfindcode = ListView1.SelectedItem.Text
End If
Unload Me
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.Move (Screen.Width - Me.Width) / 2, (frmGroup.ScaleHeight - Me.Height) / 2
ListView1.ColumnHeaders.Add , , "Group ID", 1500
ListView1.ColumnHeaders.Add , , "Group Name", 4000
'Call Module3.SetHeaderFontStyle(ListView1) '---For Bold Column Header
ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
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

Private Sub SetHeaderFontStyle() '---For Bold Column Header
   Dim LF As LOGFONT
   Dim r As Long
   Dim hCurrFont As Long
   Dim hOldFont As Long
   Dim hHeader As Long
   
  'get the windows handle to the header
  'portion of the listview
   hHeader = SendMessageLong(ListView1.hwnd, LVM_GETHEADER, 0, 0)
   
  'get the handle to the font used in the header
   hCurrFont = SendMessageLong(hHeader, WM_GETFONT, 0, 0)
   
  'get the LOGFONT details of the
  'font currently used in the header
   r = GetObject(hCurrFont, Len(LF), LF)
   
  'if GetObject was sucessful...
   If r > 0 Then
     
     'set the font attributes according to the selected check boxes
      'If chkHeaderFont(optBold).Value = 1 Then
            LF.lfWeight = FW_BOLD
      'Else: LF.lfWeight = FW_NORMAL
      'End If

      'LF.lfItalic = chkHeaderFont(optItalic).Value = 1
      'LF.lfUnderline = chkHeaderFont(optUnderlined).Value = 1
      'LF.lfStrikeOut = chkHeaderFont(optStrikeout).Value = 1
     
     'clean up by deleting any previous font
      r = DeleteObject(hHeaderFont)
      
     'create a new font for the header control to use.
     'This font must NOT be deleted until it is no
     'longer required by the control, typically when
     'the application ends (see the Unload sub), or
     'above as a new font is to be created.
      hHeaderFont = CreateFontIndirect(LF)
      
     'select the new font as the header font
      hOldFont = SelectObject(hHeader, hHeaderFont)
      
     'and inform the listview header of the change
      r = SendMessageLong(hHeader, WM_SETFONT, hHeaderFont, True)
   
   End If

End Sub
Private Sub FullRowSelect() '----For Full Row Select

   Dim rStyle As Long
   Dim r As Long
   
  'get the current ListView style
   rStyle = SendMessageLong(ListView1.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)

    
   'set the extended style bit
      rStyle = rStyle Or LVS_EX_FULLROWSELECT
   
  'set the new ListView style
   r = SendMessageLong(ListView1.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)

End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = 0
    Call Module3.lvSortByColumn(ListView1, ColumnHeader)
    ListView1.ColumnHeaders.Item(ListView1.SortKey + 1).Icon = ListView1.SortOrder + 1
End Sub
Private Sub ListView1_DblClick()
On Error GoTo ErrHandle:
strfindcode = ListView1.SelectedItem.Text
Unload Me
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtSearch_Validate True
End Sub

Private Sub txtSearch_Validate(Cancel As Boolean)
On Error GoTo ErrHandle:
Dim rsGroup As New Recordset
rsGroup.Open "select siGroupID, vGroupName From tblGroup Where vGroupName like '" & txtSearch & "%' Order By vGroupName Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsGroup.EOF Then
    Call prc_DataShowFormat(rsGroup)
    ListView1.SetFocus
Else
    ListView1.ListItems.Clear
    Frame1.Caption = "Total Group: " & ListView1.ListItems.Count
    MsgBox "Group information not found due to your search.", vbExclamation, "Remark Message"
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch)
    txtSearch.SetFocus
End If
Set rsGroup = Nothing
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Public Sub prc_DataShowFormat(rsGroup As Recordset)
On Error GoTo ErrHandle:
Dim itmX As ListItem
ListView1.ListItems.Clear
If Not rsGroup.EOF Then
    Do While Not rsGroup.EOF
        Set itmX = ListView1.ListItems.Add(, , rsGroup!siGroupID, , 3)
        itmX.SubItems(1) = rsGroup!vGroupName & ""
        rsGroup.MoveNext
    Loop
        Frame1.Caption = "Total Group: " & ListView1.ListItems.Count
Else
    MsgBox "Not Found Group! You have to input one or more Group!", vbExclamation, "Remark Message"
    Exit Sub
End If
Set rsGroup = Nothing
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
