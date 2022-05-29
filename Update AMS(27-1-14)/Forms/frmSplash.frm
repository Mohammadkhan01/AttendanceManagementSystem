VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5760
   ClientLeft      =   1020
   ClientTop       =   1500
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   6600
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      Begin VB.Image Image1 
         Height          =   6345
         Left            =   -120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8130
      End
      Begin VB.Line Line2 
         BorderColor     =   &H008080FF&
         X1              =   0
         X2              =   7440
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Phone to Phone Information System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   3570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H008080FF&
         X1              =   0
         X2              =   7440
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblLicenseTo 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome To"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   945
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Customized For"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2595
         TabIndex        =   3
         Top             =   5280
         Width           =   1560
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "GATEWAY IMART"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   330
         Left            =   2160
         TabIndex        =   4
         Top             =   5640
         Width           =   2565
      End
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   -480
      Picture         =   "frmSplash.frx":D0D8
      Stretch         =   -1  'True
      Top             =   10680
      Width           =   1095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.MousePointer = vbHourglass
End Sub

Private Sub Timer1_Timer()
    Me.MousePointer = vbDefault
    Unload Me
    frmAttendance.Show
End Sub
