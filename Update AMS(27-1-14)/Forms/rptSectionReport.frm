VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rptSectionReport 
   Caption         =   "Section Wise Report"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5475
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview Report"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Wise Attendant Report"
      ForeColor       =   &H00000000&
      Height          =   1980
      Left            =   135
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdToDate 
         Height          =   300
         Left            =   4725
         Picture         =   "rptSectionReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Calender"
         Top             =   1575
         Width           =   315
      End
      Begin VB.CommandButton cmdFromDate 
         Height          =   300
         Left            =   4725
         Picture         =   "rptSectionReport.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Calender"
         Top             =   1200
         Width           =   315
      End
      Begin VB.OptionButton OptIndv 
         Caption         =   "&Individual Employee"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbEmpId 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "rptSectionReport.frx":0624
         Left            =   2880
         List            =   "rptSectionReport.frx":0626
         TabIndex        =   4
         Text            =   "cmbEmpId"
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton optall 
         Caption         =   "Section Wise All Employees :"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblTo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2880
         TabIndex        =   12
         Top             =   1560
         Width           =   1845
      End
      Begin VB.Label lblFrom 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2865
         TabIndex        =   11
         Top             =   1200
         Width           =   1845
      End
      Begin VB.Label lbls 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2220
         TabIndex        =   10
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbls 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2070
         TabIndex        =   9
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblDID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Emp ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   570
      End
   End
   Begin Crystal.CrystalReport CryRpt 
      Left            =   0
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "rptSectionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmdMVClick As Integer

Private Sub cmdFromDate_Click()
    cmdMVClick = 1
    MonthView1.Visible = True
    MonthView1.Value = Date
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdToDate_Click()
    cmdMVClick = 2
    MonthView1.Visible = True
    MonthView1.Value = Date
End Sub

Private Sub Combo1_Click()
Call Employee_Search
End Sub


Private Sub Form_Click()
    MonthView1.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
lblFrom = Format(Now, "mm/dd/yyyy")
lblTo = Format(Now, "mm/dd/yyyy")
'--Option Searching for permission
Dim rsPermission As New ADODB.Recordset
rsPermission.Open "Select * From tblObjectPermission Where vEmployee='" & gUserId & "' And vObjectName='mnuSectionReport'", CnnMain, adOpenForwardOnly, adLockReadOnly
    If rsPermission.EOF Then
    Combo1.Enabled = False
    Else
    Combo1.Enabled = True
    End If
Set rsPermission = Nothing
       
'--Data Initialize for Section
Dim rsSection As New ADODB.Recordset
rsSection.Open "Select vSection from tblSection", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsSection.EOF Then
        While Not rsSection.EOF
            Combo1.AddItem rsSection!vSection & ""
            rsSection.MoveNext
        Wend
    End If
rsSection.Close
Set rsSection = Nothing
Combo1.ListIndex = 0
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub Employee_Search()
'--Data Initialize for cmbEmpId Combo Box
cmbEmpId.Clear
cmbEmpId.Refresh
Dim rsEmp As New Recordset
rsEmp.Open "Select vEmployee from tblEmpDetails where vSection='" & Trim(Combo1.Text) & "' Order by vEmployee Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp.EOF Then
    Do While Not rsEmp.EOF
        cmbEmpId.AddItem rsEmp!vEmployee & ""
        rsEmp.MoveNext
    Loop
    cmbEmpId.Text = gUserId
End If
rsEmp.Close
Set rsEmp = Nothing
cmbEmpId.Refresh
'cmbEmpId.ListIndex = 0
End Sub
Private Sub Frame1_Click()
    MonthView1.Visible = False
End Sub
Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Select Case cmdMVClick
    Case 1
            lblFrom = Format(DateClicked, "mm/dd/yyyy")
            MonthView1.Visible = False
            cmdToDate.SetFocus
    Case 2
            lblTo = Format(DateClicked, "mm/dd/yyyy")
            MonthView1.Visible = False
            cmdPreview.SetFocus
End Select
End Sub
Private Sub cmdPreview_Click()
On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
Dim strSql As String
If OptIndv.Value = True And cmbEmpId.Text <> "" And lblFrom <> "" And lblTo <> "" Then
    strSql = "Select * From vrSectionWiseAttendnt where vrSectionWiseAttendnt.vEmployee='" & cmbEmpId.Text & "' And vrSectionWiseAttendnt.dtDate between '" & lblFrom & "' and '" & lblTo & "' Order by vrSectionWiseAttendnt.vDesignation asc,vrSectionWiseAttendnt.vEmployee Asc, vrSectionWiseAttendnt.dtDate Asc"
ElseIf optall.Value = True And lblFrom <> "" And lblTo <> "" Then
    strSql = "Select * From vrSectionWiseAttendnt Where vrSectionWiseAttendnt.dtDate Between '" & lblFrom & "' and '" & lblTo & "' and vrSectionWiseAttendnt.vSection='" & Combo1.Text & "' Order by vrSectionWiseAttendnt.vDesignation asc,vrSectionWiseAttendnt.vEmployee Asc, vrSectionWiseAttendnt.dtDate Asc"
Else
    MsgBox "Please enter valid data!", vbExclamation, "Validation Message"
    Me.MousePointer = vbDefault
    Exit Sub
End If
With CryRpt
    .Connect = "DSN =" & gstrServName & "; UID = " & gstrDBUID & "; pwd=" & gstrDBUpass & "; DATABASE = " & gstrDBName
    .SQLQuery = strSql
    If OptIndv.Value = True Then
        .ReportFileName = App.Path & "\Report\Report2.rpt"
    Else
        .ReportFileName = App.Path & "\Report\Report2.rpt"
    End If
    .Formulas(0) = "User= '" & gUserId & "'"
    .Formulas(1) = "FromDate= '" & Format(lblFrom, "dd/mm/yyyy") & "'"
    .Formulas(2) = "ToDate= '" & Format(lblTo, "dd/mm/yyyy") & "'"
    .WindowTitle = "Date Wise Attendant Report"
    .WindowState = crptMaximized
    .Action = 1
End With
strSql = ""
Me.MousePointer = vbDefault
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Me.MousePointer = vbDefault
End Sub

