VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CBIS - Section "
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmSection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9340
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   4815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   3360
      TabIndex        =   8
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox cmbSection 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   0
      Picture         =   "frmSection.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Section Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intPosition As Integer


Private Sub cmbSection_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
If KeyAscii = 13 Then cmbSection_Validate True
End Sub

Private Sub cmbSection_Validate(Cancel As Boolean)
If cmbSection = "" Then Exit Sub
ListView1.ListItems.Clear
Dim itmX As ListItem
Dim rstcmb As New Recordset
rstcmb.Open "Select vEmployee,vDesignation,vSection from tblEmpDetails where vSection='" & cmbSection.Text & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rstcmb.EOF Then
        Do While Not rstcmb.EOF
            Set itmX = ListView1.ListItems.Add(, , rstcmb!vSection & "")
            itmX.SubItems(1) = rstcmb!vEmployee & ""
            itmX.SubItems(2) = rstcmb!vDesignation & ""
            rstcmb.MoveNext
        Loop
    Else
    '-----Nothing
    End If
Set rstcmb = Nothing
Frame1.Caption = ListView1.ListItems.Count

End Sub

Private Sub Command1_Click()
cmbSection.Text = ""
cmbSection.SetFocus
End Sub

Private Sub Command2_Click()
If cmbSection = "" Then
MsgBox "Section should not be blank ! Please Enter Section Name", vbExclamation, "Remark Message"
cmbSection.SetFocus
Exit Sub
End If
Dim rsTemp As New ADODB.Recordset
rsTemp.Open "Select vSection from tblSection where vSection='" & cmbSection.Text & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsTemp.EOF Then
    MsgBox "This is existing Section", vbInformation, "Remark Message"
    cmbSection.SelStart = 0
    cmbSection.SelLength = Len(cmbSection)
    cmbSection.SetFocus
    Exit Sub
    Else
End If

CnnMain.BeginTrans
CnnMain.Execute "Exec Sp_Add_tblSection '" & cmbSection & "','" & gUserId & "'"
CnnMain.CommitTrans
MsgBox "Data has been inserted to Database successfully", vbInformation, "Remark Message"
cmbSection_Validate True
Call Search
Exit Sub
End Sub

Private Sub Command3_Click()
If cmbSection = "" Then Exit Sub
CnnMain.BeginTrans
CnnMain.Execute "Delete from tblSection where vSection='" & cmbSection.Text & "'"
CnnMain.CommitTrans
Call Search
End Sub

Private Sub Command4_Click()
cmbSection = ""
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Search()
List1.Clear
Dim rstSection As New Recordset
rstSection.Open "Select vSection from tblSection", CnnMain, adOpenForwardOnly, adLockReadOnly
While Not rstSection.EOF
    cmbSection.AddItem rstSection(0) & ""
    List1.AddItem rstSection(0) & ""
    rstSection.MoveNext
Wend
rstSection.Close
Set rstSection = Nothing
End Sub
Private Sub Form_Load()
frmSection.Move (Screen.Width - frmSection.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
ListView1.ColumnHeaders.Add , , "Section", 1500, 0
ListView1.ColumnHeaders.Add , , "Employee", 1500, 0
ListView1.ColumnHeaders.Add , , "Disignation", 1600, 0
Call Search
End Sub

Private Sub List1_Click()
'Dim rstPosition As New Recordset
'rstPosition.Open "Select vsection from tblSection", CnnMain, adOpenForwardOnly, adLockReadOnly

intPosition = List1.ListIndex
cmbSection.Text = List1.Text
cmbSection_Validate True
End Sub
