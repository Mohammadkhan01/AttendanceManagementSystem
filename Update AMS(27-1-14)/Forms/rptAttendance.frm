VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rptAttendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Wise Attendant Report"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "rptAttendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   120
      TabIndex        =   12
      Top             =   240
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
   Begin Crystal.CrystalReport CryRpt 
      Left            =   0
      Top             =   2040
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
   Begin VB.Frame Frame1 
      Caption         =   "Date Wise Attendant Report"
      ForeColor       =   &H00000000&
      Height          =   2220
      Left            =   135
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2880
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Designation Wise"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbEmpId 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "rptAttendance.frx":0ECA
         Left            =   2865
         List            =   "rptAttendance.frx":0ECC
         TabIndex        =   13
         Text            =   "cmbEmpId"
         Top             =   435
         Width           =   2175
      End
      Begin VB.OptionButton OptIndv 
         Caption         =   "&Individual Employee"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   225
         TabIndex        =   6
         Top             =   435
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "&All Employees"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdFromDate 
         Height          =   300
         Left            =   4725
         Picture         =   "rptAttendance.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Calender"
         Top             =   1200
         Width           =   315
      End
      Begin VB.CommandButton cmdToDate 
         Height          =   300
         Left            =   4725
         Picture         =   "rptAttendance.frx":11E0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Calender"
         Top             =   1575
         Width           =   315
      End
      Begin VB.Label lblDID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Emp ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2250
         TabIndex        =   11
         Top             =   435
         Width           =   570
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
         TabIndex        =   10
         Top             =   1200
         Width           =   780
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
         TabIndex        =   9
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lblFrom 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2865
         TabIndex        =   8
         Top             =   1200
         Width           =   1845
      End
      Begin VB.Label lblTo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   1560
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview Report"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H000000FF&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "rptAttendance"
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

Private Sub Form_Click()
    MonthView1.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.Move (Screen.Width - Me.Width) / 2, (frmMain.ScaleHeight - Me.Height) / 2
lblFrom = Format(Now, "mm/dd/yyyy")
lblTo = Format(Now, "mm/dd/yyyy")
'--Data Initialize for cmbEmpId Combo Box
Dim rsEmp As New Recordset
rsEmp.Open "Select vEmployee from tblEmpDetails Order by vEmployee Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp.EOF Then
    Do While Not rsEmp.EOF
        cmbEmpId.AddItem rsEmp!vEmployee & ""
        rsEmp.MoveNext
    Loop
    cmbEmpId.Text = gUserId
End If
Set rsEmp = Nothing
Dim rsDesignation As New ADODB.Recordset
rsDesignation.Open "Select distinct(vDesignation) from tblEmpDetails", CnnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsDesignation.EOF Then
        Do While Not rsDesignation.EOF
            Combo1.AddItem rsDesignation!vDesignation & ""
            rsDesignation.MoveNext
        Loop
    End If
Set rsDesignation = Nothing
Combo1.ListIndex = 0
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
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
'On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass

        
Dim strSql As String
If OptIndv.Value = True And cmbEmpId.Text <> "" And lblFrom <> "" And lblTo <> "" Then
    strSql = "Select * From vrDateWiseAttendant_1 where vrDateWiseAttendant_1.vEmployee='" & cmbEmpId.Text & "' And vrDateWiseAttendant_1.dtDate between '" & lblFrom & "' and '" & lblTo & "' Order by vrDateWiseAttendant_1.vDesignation asc,vrDateWiseAttendant_1.vEmployee Asc, vrDateWiseAttendant_1.dtDate Asc"
ElseIf optall.Value = True And lblFrom <> "" And lblTo <> "" Then
    strSql = "Select * From vrDateWiseAttendant_1 Where vrDateWiseAttendant_1.dtDate Between '" & lblFrom & "' and '" & lblTo & "' Order by vrDateWiseAttendant_1.vDesignation asc,vrDateWiseAttendant_1.vEmployee Asc, vrDateWiseAttendant_1.dtDate Asc"
ElseIf Option1.Value = True And lblFrom <> "" And lblTo <> "" Then
    strSql = "Select * From vrDateWiseAttendant_1 Where vrDateWiseAttendant_1.dtDate Between '" & lblFrom & "' and '" & lblTo & "' and vrDateWiseAttendant_1.vDesignation ='" & Trim(Combo1.Text) & "' order by vrDateWiseAttendant_1.vEmployee asc,vrDateWiseAttendant_1.dtDate asc"
Else
    MsgBox "Please enter valid data!", vbExclamation, "Validation Message"
    Me.MousePointer = vbDefault
    Exit Sub
End If
With CryRpt
    .Connect = "DSN =" & gstrServName & "; UID = " & gstrDBUID & "; pwd=" & gstrDBUpass & "; DATABASE = " & gstrDBName
    .SQLQuery = strSql
    If OptIndv.Value = True Then
        .ReportFileName = App.Path & "\Report\DateWiseAttendt.rpt"
    ElseIf Option1.Value = True Then
        .ReportFileName = App.Path & "\Report\DateWiseAttendt.rpt"
    Else
        .ReportFileName = App.Path & "\Report\DateWiseAttendt.rpt"
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
'ErrHandle:
'    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
' r = vbDefault
'    Err.Clear
    Me.MousePointer = vbDefault
End Sub
