VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Details Information"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ForeColor       =   &H8000000F&
   Icon            =   "frmEmpDetails.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Navigation"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   157
      TabIndex        =   39
      Top             =   6000
      Width           =   7050
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1935
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   3045
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4155
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   375
         Left            =   5265
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   405
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   825
         MaskColor       =   &H000000FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   405
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employe Details Information"
      ForeColor       =   &H00000000&
      Height          =   5895
      Left            =   150
      TabIndex        =   19
      Top             =   120
      Width           =   7050
      Begin VB.CheckBox chkPhotoUpdate 
         Caption         =   "Update Picture"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   56
         Top             =   3240
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker txtLogout 
         Height          =   255
         Left            =   5160
         TabIndex        =   55
         Top             =   3960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20709378
         CurrentDate     =   38046
      End
      Begin MSComCtl2.DTPicker txtLogin 
         Height          =   255
         Left            =   5160
         TabIndex        =   54
         Top             =   3600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20709378
         CurrentDate     =   38046
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Only Active Employee"
         Height          =   195
         Left            =   1440
         TabIndex        =   53
         Top             =   240
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5160
         TabIndex        =   52
         Top             =   5160
         Width           =   1695
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   1440
         TabIndex        =   40
         Top             =   3510
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
         StartOfWeek     =   20709377
         TitleBackColor  =   8438015
         TitleForeColor  =   0
         TrailingForeColor=   -2147483633
         CurrentDate     =   37113
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00BDD2D2&
         Height          =   300
         Left            =   6000
         Picture         =   "frmEmpDetails.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   49
         Tag             =   "PK"
         ToolTipText     =   "Browse Picture"
         Top             =   315
         Width           =   315
      End
      Begin VB.Frame Frame3 
         Caption         =   "Employee's Picture      "
         ForeColor       =   &H00000000&
         Height          =   2790
         Left            =   4440
         TabIndex        =   50
         Top             =   360
         Width           =   2535
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   2070
            Left            =   180
            Top             =   300
            Width           =   2160
         End
      End
      Begin VB.ComboBox cmbActive 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmEmpDetails.frx":09C4
         Left            =   5190
         List            =   "frmEmpDetails.frx":09CE
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4725
         Width           =   1665
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
         Left            =   5190
         MaxLength       =   10
         PasswordChar    =   "l"
         TabIndex        =   13
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   4320
         Width           =   1665
      End
      Begin VB.ComboBox txtEmpID 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   570
         Width           =   2475
      End
      Begin VB.TextBox txtGroupName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   46
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   2040
         Width           =   2160
      End
      Begin VB.TextBox txtDeptName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   45
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   1680
         Width           =   1830
      End
      Begin VB.TextBox txtGroupID 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   43
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   2040
         Width           =   600
      End
      Begin VB.CheckBox chkMarried 
         Caption         =   "Married"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3300
         TabIndex        =   38
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4845
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   5400
         Visible         =   0   'False
         Width           =   1965
      End
      Begin MSComDlg.CommonDialog CommDlg 
         Left            =   4440
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdHD 
         BackColor       =   &H00BDD2D2&
         Height          =   300
         Left            =   2460
         Picture         =   "frmEmpDetails.frx":09DB
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "PK"
         ToolTipText     =   "Search Screen"
         Top             =   5520
         Width           =   315
      End
      Begin VB.CommandButton cmdBD 
         BackColor       =   &H00BDD2D2&
         Height          =   300
         Left            =   2460
         Picture         =   "frmEmpDetails.frx":0CED
         Style           =   1  'Graphical
         TabIndex        =   34
         Tag             =   "PK"
         ToolTipText     =   "Search Screen"
         Top             =   5160
         Width           =   315
      End
      Begin VB.TextBox txtSalary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3300
         MaxLength       =   25
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   5520
         Width           =   945
      End
      Begin VB.CommandButton cmdFindEmp 
         BackColor       =   &H00BDD2D2&
         Height          =   300
         Left            =   3930
         Picture         =   "frmEmpDetails.frx":0FFF
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "PK"
         ToolTipText     =   "Search Screen"
         Top             =   570
         Width           =   315
      End
      Begin VB.TextBox txtHD 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   11
         ToolTipText     =   "Maximum 10 Characters"
         Top             =   5520
         Width           =   1005
      End
      Begin VB.TextBox txtEmpName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   2
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   960
         Width           =   2805
      End
      Begin VB.TextBox txtDesignation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   3
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   1320
         Width           =   2805
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   9
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   4800
         Width           =   2805
      End
      Begin VB.TextBox txtDeptId 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   4
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox txtFName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   5
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   2400
         Width           =   2805
      End
      Begin VB.TextBox txtBD 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         ToolTipText     =   "Maximum 10 Characters"
         Top             =   5160
         Width           =   1005
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   8
         ToolTipText     =   "Maximum 25 Characters"
         Top             =   4440
         Width           =   2805
      End
      Begin VB.CommandButton cmdFindDept 
         BackColor       =   &H00BDD2D2&
         Height          =   300
         Left            =   3930
         Picture         =   "frmEmpDetails.frx":10F9
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "PK"
         ToolTipText     =   "Search Screen"
         Top             =   1680
         Width           =   315
      End
      Begin VB.TextBox txtPresentAdd 
         BackColor       =   &H00FFFFFF&
         Height          =   750
         Left            =   1440
         MaxLength       =   256
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         ToolTipText     =   "Maximum 100 Characters"
         Top             =   2760
         Width           =   2805
      End
      Begin VB.TextBox txtPmtAdd 
         BackColor       =   &H00FFFFFF&
         Height          =   750
         Left            =   1440
         MaxLength       =   256
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         ToolTipText     =   "Maximum 100 Characters"
         Top             =   3600
         Width           =   2805
      End
      Begin VB.Label Label5 
         Caption         =   "Section :"
         Height          =   255
         Left            =   4440
         TabIndex        =   51
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Active:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4650
         TabIndex        =   48
         Top             =   4725
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4365
         TabIndex        =   47
         Top             =   4320
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Group:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   930
         TabIndex        =   44
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log Out:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4440
         TabIndex        =   42
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log In:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4500
         TabIndex        =   41
         Top             =   3600
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         Height          =   195
         Left            =   4440
         TabIndex        =   37
         Top             =   5400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblBasiSalary 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2805
         TabIndex        =   33
         Top             =   5520
         Width           =   480
      End
      Begin VB.Label lblHD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Hire Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   705
         TabIndex        =   32
         Top             =   5520
         Width           =   705
      End
      Begin VB.Label lblDID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   465
         TabIndex        =   31
         Top             =   585
         Width           =   945
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   660
         TabIndex        =   30
         Top             =   5175
         Width           =   750
      End
      Begin VB.Label lblPropitors 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   225
         TabIndex        =   29
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label lblDesignation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   525
         TabIndex        =   28
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label lblFName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   345
         TabIndex        =   27
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label lblPresentAdd 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Present Address:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   26
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Label lblDeptId 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         TabIndex        =   25
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   990
         TabIndex        =   24
         Top             =   4800
         Width           =   420
      End
      Begin VB.Label lblPmtAddress 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Pmt. Address:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   435
         TabIndex        =   23
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lblPhone 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   900
         TabIndex        =   22
         Top             =   4440
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmEmpDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmdMVClick As Integer '--For Month View Clicked
'--For Image
Dim strPRInfo As String
Dim lngOffset As Long
Dim lngLogoSize As Long
Dim varLogo As Variant
Dim varChunk As Variant
'----------for monthviue--------
Dim blnBD As Boolean
Dim blnHD As Boolean

Public StrFileName As String
Const conChunkSize = 100

Private Sub Check1_Click()
 txtEmpID.Clear
txtEmpID = ""
txtEmpID.Refresh
Dim Control As Object
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
If Check1.Value = 1 Then
Dim rsEmp1 As New Recordset
rsEmp1.Open "Select vEmployee From tblEmpDetails where vActive='Yes' Order By vEmployee Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp1.EOF Then
    Do While Not rsEmp1.EOF
        txtEmpID.AddItem rsEmp1!vEmployee & ""
        rsEmp1.MoveNext
    Loop
End If
Set rsEmp1 = Nothing
End If
If Check1.Value = 0 Then
Dim rsEmp As New Recordset
rsEmp.Open "Select vEmployee From tblEmpDetails Order By vEmployee Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp.EOF Then
    Do While Not rsEmp.EOF
        txtEmpID.AddItem rsEmp!vEmployee & ""
        rsEmp.MoveNext
    Loop
End If
Set rsEmp = Nothing
End If
End Sub

Private Sub Frame1_Click()
    MonthView1.Visible = False
End Sub

Public Sub Disable_Monthview()
    MonthView1.Enabled = False
    MonthView1.Visible = False
End Sub

Private Sub cmdBD_Click()
    cmdMVClick = 1
    MonthView1.Visible = True
    MonthView1.Value = Date
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo ErrHandle:
chkPhotoUpdate.Value = 1
CommDlg.Filter = "Bmp Files (*.bmp)|*.bmp|Jpg Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif|All Files (*.*)|*.*"
CommDlg.ShowOpen
Image1.Picture = LoadPicture(CommDlg.FileName)
txtPath = CommDlg.FileName
StrFileName = CommDlg.FileName
If txtPath = "" Then
    Image1.Picture = LoadPicture(App.Path & "\Image\DefaultImage.jpg")
    txtPath = App.Path & "\Image\DefaultImage.jpg"
    StrFileName = txtPath
    chkPhotoUpdate.Value = 0
Else
    '-- Nothing
End If
Exit Sub
ErrHandle:
    Select Case Err.Number
        Case 481
            MsgBox "Format does not support!", vbExclamation, "Remark Message"
            Err.Clear
        Case Else
            MsgBox "Error ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
            Err.Clear
    End Select
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrHandle:
Dim Control As Object
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
prc_DisAllow
txtEmpID.Enabled = True
txtEmpID.BackColor = &HFFFFFF
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

Private Sub cmdFindDept_Click()
On Error GoTo ErrHandle:
frmSearchDept.Show vbModal
If strfindcode <> "" Then
    txtDeptId = strfindcode
    txtDeptID_Validate True
    strfindcode = ""
Else
    '--Nothing
End If
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdFindEmp_Click()
On Error GoTo ErrHandle:
frmSearchEmp.Show vbModal
If strfindcode <> "" Then
    txtEmpID = strfindcode
    txtEmpId_Validate True
    strfindcode = ""
Else
    '--Nothing
End If
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub cmdHD_Click()
    cmdMVClick = 2
    MonthView1.Visible = True
    MonthView1.Value = Date
End Sub
Private Sub cmdNew_Click()
On Error GoTo ErrHandle:
Dim Control As Object
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
Call prc_Allow
cmbActive.ListIndex = 0
'txtEmpId.SetFocus
Image1.Picture = LoadPicture(App.Path & "\Image\DefaultImage.jpg")
txtPath = App.Path & "\Image\DefaultImage.jpg"
StrFileName = txtPath
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub cmdSave_Click()
On Error GoTo ErrHandle:
If Trim(txtEmpID.Text) = "" Then
    MsgBox "Employee ID can't be blank!" & "  Please enter Employee ID", vbExclamation, "Remark Message"
    txtEmpID.SetFocus
    Exit Sub
End If
If Trim(txtEmpName) = "" Then
    MsgBox "Employee Name can't be blank!" & "  Please Enter Employee Name", vbExclamation, "Remark Message"
    txtEmpName.SetFocus
    Exit Sub
End If
If Trim(txtDeptId) = "" Then
    MsgBox "Department Name can't be blank!" & "  Please select Department Name!", vbExclamation, "Remark Message"
    Exit Sub
End If
If Len(txtPassword) < 6 Then
    MsgBox "Password should be minimum 6 digit", vbInformation, "Remark Message"
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
    txtPassword.SetFocus
    Exit Sub
End If
If Trim(txtSalary) = "" Then
    MsgBox "Employee's Salary can't be blank!" & "  Please enter Salary!", vbExclamation, "Remark Message"
    txtSalary.SetFocus
    Exit Sub
End If
'--Name Already Exist
Dim rsName As New Recordset
rsName.Open "Select vEmployee from tblEmpDetails Where vEmployee='" & txtEmpID.Text & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsName.EOF Then
MsgBox "Employee's name already exist! Employee's Name should not be duplicate!", vbExclamation, "Remark Message"
txtEmpID.SetFocus
Exit Sub
End If
Set rsName = Nothing
Me.MousePointer = vbHourglass
'--Save to Employee Details Table
CnnMain.BeginTrans
CnnMain.Execute "Exec Sp_Add_tblEmpDetails '" & Trim(txtEmpID.Text) & "','" & Trim(txtEmpName) & "','" & Trim(txtDesignation) & "'," & Trim(txtDeptId) & ", '" & Trim(txtFName) & "','" & Trim(txtPresentAdd) & "','" & Trim(txtPmtAdd) & "','" & Trim(txtPhone) & "','" & Trim(txtEmail) & "','" & Trim(txtBD) & "','" & Trim(txtHD) & "'," & Trim(chkMarried.Value) & ", " & Trim(txtSalary) & ", '" & Trim(txtLogin) & "', '" & Trim(txtLogout) & "', '" & Trim(txtPassword) & "', '" & cmbActive.Text & "','" & Combo1.Text & "'"
CnnMain.CommitTrans
'--Convert Image to Binary
Dim strFile As String
Dim strGetImage As String
strFile = GetImageData(StrFileName)
strGetImage = StrFileName
Dim rsImage As New Recordset
rsImage.Open "Select iPhoto From tblEmpDetails where vEmployee='" & txtEmpID.Text & "'", CnnMain, adOpenDynamic, adLockOptimistic
rsImage!iPhoto = GetImageData(StrFileName)
rsImage.Update
rsImage.Requery
Set rsImage = Nothing
Call prc_DisAllow
'--For Menu Permission
Dim rsMenu As New Recordset
rsMenu.Open "Select * From tblObject Where vDefault='Yes'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsMenu.EOF Then
    Do While Not rsMenu.EOF
        CnnMain.BeginTrans
        CnnMain.Execute "Exec Sp_Add_tblObjectPermission '" & txtEmpID & "','" & rsMenu!vObjectName & "','" & rsMenu!vDefault & "'"
        CnnMain.CommitTrans
        rsMenu.MoveNext
    Loop
End If
Set rsMenu = Nothing
'txtEmpId.Enabled = True
txtEmpID.SelStart = Len(txtEmpID.Text)
txtEmpID.SetFocus
txtEmpID.BackColor = &HFFFFFF
Me.MousePointer = vbDefault
MsgBox "Record successfully saved to the database.", vbInformation, "Save Message"
txtEmpName.SetFocus
Exit Sub
ErrHandle:
    CnnMain.RollbackTrans
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Select Case cmdMVClick
    Case 0
            'lblFrom = Format(DateClicked, "dd/mm/yyyy")
    Case 1
            txtBD = Format(DateClicked, "mm/dd/yyyy")
    Case 2
            txtHD = Format(DateClicked, "mm/dd/yyyy")
End Select
MonthView1.Visible = False
End Sub

Private Sub txtDeptID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 12 Or KeyAscii > 14 And KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtDeptID_Validate False
End Sub

Private Sub txtDesignation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtEmpID_Click()
    txtEmpId_Validate True
End Sub

Private Sub txtEmpID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtEmpId_Validate True
End Sub

Private Sub txtEmpId_Validate(Cancel As Boolean)
On Error GoTo ErrHandle:
If Trim(txtEmpID.Text) = "" Then
   cmdNew_Click
   Exit Sub
End If

Dim rsEmp As New Recordset
rsEmp.Open "Select * From tblEmpDetails Where vEmployee='" & Trim(txtEmpID.Text) & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp.EOF Then
    Me.MousePointer = vbHourglass
    With rsEmp
        txtEmpName = !vEmplName & ""
        txtDesignation = !vDesignation & ""
        txtDeptId = !siDeptID & ""
        txtDeptID_Validate True
        txtFName = !vFName & ""
        txtPresentAdd = !vPresentAdd & ""
        txtPmtAdd = !vPmtAdd & ""
        txtPhone = !vPhone & ""
        txtEmail = !vEmail & ""
        txtBD = Format(!dtBD, "dd/mm/yyyy") & ""
        txtHD = Format(!dtHD, "dd/mm/yyyy") & ""
        txtSalary = !mSalary & ""
        chkMarried.Value = !siStatus & ""
        txtPassword = !Password & ""
        cmbActive.Text = !vActive & ""
        Combo1.Text = !vSection & ""
        txtLogin = !vLogin & ""
        txtLogout = !vLogOut & ""
        
    End With
    Set Image1.DataSource = rsEmp
    Image1.DataField = "iPhoto"
    txtPath = ""
    prc_DisAllow
    'cmdUpdate.Enabled = True
    Me.MousePointer = vbDefault
    txtEmpName.SelStart = 0
    txtEmpName.SelLength = Len(txtEmpName)
    txtEmpName.SetFocus
Else
    Me.MousePointer = vbHourglass
    Dim s As String
    s = Trim(txtEmpID.Text)
    cmdNew_Click
    txtEmpID = s
    txtEmpName.SetFocus
    Me.MousePointer = vbDefault
End If
Set rsEmp = Nothing
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrHandle:
If Trim(txtEmpID.Text) = "" Then
    MsgBox "Employee ID can't be blank!" & "  Please enter Employee ID", vbExclamation, "Remark Message"
    txtEmpName.SetFocus
    Exit Sub
End If
If Trim(txtEmpName) = "" Then
    MsgBox "Employee Name can't be blank!" & "  Please Enter Employee Name", vbExclamation, "Remark Message"
    txtEmpName.SetFocus
    Exit Sub
End If
If Trim(txtDeptId) = "" Then
    MsgBox "Department Name can't be blank!" & "  Please select Department Name!", vbExclamation, "Remark Message"
    Exit Sub
End If
If Len(txtPassword) < 6 Then
    MsgBox "Password should be minimum 6 digit", vbInformation, "Remark Message"
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
    txtPassword.SetFocus
    Exit Sub
End If
If Trim(txtSalary) = "" Then
    MsgBox "Employee's Salary can't be blank!" & "  Please enter Salary!", vbExclamation, "Remark Message"
    txtSalary.SetFocus
    Exit Sub
End If

If Trim(txtPath) = "" And chkPhotoUpdate.Value = 1 Then
    MsgBox "If you checked the Photo Update Check Box then you have to select a Photo!", vbExclamation, "Remark Message"
    chkPhotoUpdate.SetFocus
    Exit Sub
End If
Me.MousePointer = vbHourglass
'--Sp_Up_tblEmpDetails-------
If chkPhotoUpdate.Value = 1 And txtPath <> "" Then
    CnnMain.BeginTrans
    CnnMain.Execute "Sp_Up_tblEmpDetails '" & Trim(txtEmpID.Text) & "','" & Trim(txtEmpName) & "','" & Trim(txtDesignation) & "'," & Trim(txtDeptId) & ", '" & Trim(txtFName) & "','" & Trim(txtPresentAdd) & "','" & Trim(txtPmtAdd) & "','" & Trim(txtPhone) & "','" & Trim(txtEmail) & "','" & Format(txtBD, "mm/dd/yyyy") & "','" & Format(txtHD, "mm/dd/yyyy") & "'," & Trim(chkMarried.Value) & ", " & Trim(txtSalary) & ", '" & Trim(txtLogin) & "', '" & Trim(txtLogout) & "', '" & Trim(txtPassword) & "', '" & cmbActive.Text & "','" & Trim(Combo1.Text) & "'"
    CnnMain.CommitTrans
    '-------Image Update
    Dim strFile As String
    Dim strGetImage As String
    strFile = GetImageData(StrFileName)
    strGetImage = StrFileName
    Dim rsImage As New Recordset
    rsImage.Open "Select iPhoto From tblEmpDetails where vEmployee='" & Trim(txtEmpID.Text) & "'", CnnMain, adOpenDynamic, adLockOptimistic
    rsImage!iPhoto = GetImageData(StrFileName)
    rsImage.Update
    rsImage.Requery
    Set rsImage = Nothing
Else
    CnnMain.BeginTrans
    CnnMain.Execute "Sp_Up_tblEmpDetails '" & Trim(txtEmpID.Text) & "','" & Trim(txtEmpName) & "','" & Trim(txtDesignation) & "'," & Trim(txtDeptId) & ", '" & Trim(txtFName) & "','" & Trim(txtPresentAdd) & "','" & Trim(txtPmtAdd) & "','" & Trim(txtPhone) & "','" & Trim(txtEmail) & "','" & Format(txtBD, "mm/dd/yyyy") & "','" & Format(txtHD, "mm/dd/yyyy") & "'," & Trim(chkMarried.Value) & ", " & Trim(txtSalary) & ", '" & Trim(txtLogin) & "', '" & Trim(txtLogout) & "', '" & Trim(txtPassword) & "','" & cmbActive.Text & "','" & Combo1.Text & "'"
    CnnMain.CommitTrans
End If
Me.MousePointer = vbDefault
MsgBox "Record successfully updated to the database.", vbInformation, "Successfull Update"
chkPhotoUpdate.Value = 0
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub Form_Load()
On Error GoTo ErrHandle:
Me.Move (Screen.Width - Me.Width) / 2, (frmGroup.ScaleHeight - Me.Height) / 2
'==Data input for txtEmpID Combo Box
If Check1.Value = 0 Then
Dim rsEmp As New Recordset
rsEmp.Open "Select vEmployee From tblEmpDetails Order By vEmployee Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp.EOF Then
    Do While Not rsEmp.EOF
        txtEmpID.AddItem rsEmp!vEmployee & ""
        rsEmp.MoveNext
    Loop
End If
Set rsEmp = Nothing
ElseIf Check1.Value = 1 Then
Dim rsEmp1 As New Recordset
rsEmp1.Open "Select vEmployee From tblEmpDetails where vActive='Yes' Order By vEmployee Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp1.EOF Then
    Do While Not rsEmp1.EOF
        txtEmpID.AddItem rsEmp1!vEmployee & ""
        rsEmp1.MoveNext
    Loop
End If
Set rsEmp1 = Nothing
End If



Dim rsSection As New Recordset
rsSection.Open "Select vSection from tblSection", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsSection.EOF Then
    Do While Not rsSection.EOF
        Combo1.AddItem rsSection!vSection & ""
        rsSection.MoveNext
    Loop
End If
'==
Call prc_DisAllow
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
            If obj.Name = "cmdSave" Or obj.Name = "cmdCancel" Or obj.Name = "cmdBrowse" Or obj.Name = "cmdFindEmp" Or obj.Name = "cmdFindDept" Or obj.Name = "cmdBD" Or obj.Name = "cmdHD" Then
                obj.Enabled = True
            Else
                obj.Enabled = False
            End If
        End If
    Next
End Sub
Public Sub clear_all_textboxes()
        txtEmpID.Text = ""
        txtEmpName = ""
        txtFName = ""
End Sub
Private Sub txtEmpName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    'If KeyAscii = 13 Then txtDescription.SetFocus
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Public Sub Data_Show()
Dim rsEmp As New ADODB.Recordset
rsEmp.Open "Select * from CBtblEmpDetails where vEmployee='" & Trim(txtEmpID.Text) & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsEmp.EOF Then
    With rsEmp
        txtEmpName = !vEmpName & ""
        txtFName = !vFName & ""
        txtDeptId = !iDeptId & ""
        txtPresentAdd = !vPresentAdd & ""
        txtPmtAdd = !vPmtAdd & ""
        chkMarried.Value = !iStatus & ""
        txtPhone = !vPhone & ""
        txtEmail = !vEmail & ""
        txtBD = Format(!dtBD, "mm/dd/yyyy") & ""
        txtHD = Format(!dtHD, "mm/dd/yyyy") & ""
        txtSalary = !mSalary & ""
        txtDesignation = !vDesignation & ""
        txtPassword = !Password & ""
        txtLogin = !vLogin & ""
        txtLogout = !vLogOut & ""
        '--xx
        Set Image1.DataSource = rsEmp
        Image1.DataField = "iPhoto"
        txtPath = ""
    End With
Else
    MsgBox "Invalid EmpID!", vbExclamation, "Validation Message"
    txtEmpID.SelStart = 0
    txtEmpID.SelLength = Len(txtEmpID.Text)
    clear_all_textboxes
End If
'Set Image1.DataSource = Nothing
'Set rsEmp = Nothing
End Sub
Private Sub txtDeptID_Validate(Cancel As Boolean)
On Error GoTo ErrHandle:
If txtDeptId = "" Then
    Exit Sub
End If
Dim rsDept As New ADODB.Recordset
rsDept.Open "Select * from tblDepartment Where siDeptId='" & Trim(txtDeptId) & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsDept.EOF Then
    Dim rsGroup As New Recordset
    txtDeptName = rsDept!vDeptName & ""
    txtGroupID = rsDept!siGroupID & ""
    rsGroup.Open "Select vGroupName From tblGroup Where siGroupID=" & rsDept!siGroupID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
    txtGroupName = rsGroup!vGroupName & ""
    Set rsGroup = Nothing
Else
    MsgBox "Invalid Department ID! Please enter valid Department ID!", vbExclamation, "Validation Message"
    Cancel = True
    txtDeptId.SelStart = 0
    txtDeptId.SelLength = Len(txtDeptId)
    txtDeptName = ""
    txtGroupID = ""
    txtGroupName = ""
End If
Set rsDept = Nothing
txtFName.SetFocus
Exit Sub
ErrHandle:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
Private Sub txtFName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtLogIn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtLogOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtPmtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtPresentAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 12 Or KeyAscii > 14 And KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetImageData(ByVal FileName As String) As Byte() 'Byte Array to presurve Byte data of the Photofile
On Error Resume Next
    Dim bytData() As Byte
    Dim intFreeFile As Integer
    If Not FileExists(FileName) Then ' To veryfy the File Path
    On Error Resume Next
        Err.Raise 53, "GetImageData", Err.Description
    End If
    On Error GoTo ErrorHandler
    intFreeFile = FreeFile()
    Open FileName For Binary As #intFreeFile ' To convert the File as binary
    ReDim bytData(FileLen(FileName)) As Byte ' To alocate space for the binary file
    Get #intFreeFile, , bytData
    Close #intFreeFile
    GetImageData = bytData
    Exit Function
ErrorHandler:
    On Error Resume Next
    Err.Raise Err.Number, GetImageData, Err.Description
End Function
Private Function FileExists(ByVal FileName As String) As Boolean
Dim intAttr As Integer
On Error GoTo ErrorFileExist
intAttr = GetAttr(FileName)
FileExists = ((intAttr And vbDirectory) = 0)
GoTo ExitFileExist
ErrorFileExist:
FileExists = False
Resume ExitFileExist
ExitFileExist:
On Error GoTo 0
End Function
