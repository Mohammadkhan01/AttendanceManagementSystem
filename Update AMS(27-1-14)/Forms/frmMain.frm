VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Management System"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8940
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   600
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1181
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1471
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":175A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":200C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":383F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4119
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6481
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7635
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":87E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":90C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A877
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CompanyInfo"
            Object.ToolTipText     =   "Company Information"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ChangePassword"
            Object.ToolTipText     =   "Change Password"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ChangeActivity"
            Object.ToolTipText     =   "Change Activity"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "GroupInfo"
            Object.ToolTipText     =   "Group Information"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DepartmentInfo"
            Object.ToolTipText     =   "Department Information"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EmployeeInfo"
            Object.ToolTipText     =   "Employee Information"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RemarkInfo"
            Object.ToolTipText     =   "Remark Information"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ObjectInfo"
            Object.ToolTipText     =   "Object Information"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ObjectPermission"
            Object.ToolTipText     =   "Object Permission"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Options"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TechSupport"
            Object.ToolTipText     =   "Techinical Support"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About Application"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Application"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7035
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3254
            MinWidth        =   3245
            Picture         =   "frmMain.frx":ACC9
            Key             =   "stbUser"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4400
            MinWidth        =   2998
            Picture         =   "frmMain.frx":B5A3
            Key             =   "sbrRole"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2990
            MinWidth        =   2999
            Picture         =   "frmMain.frx":BE7D
            Key             =   "sbrDate"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2470
            MinWidth        =   2470
            Picture         =   "frmMain.frx":C757
            TextSave        =   "11:51 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Picture         =   "frmMain.frx":D031
      Stretch         =   -1  'True
      Top             =   360
      Width           =   9075
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCompanyInfo 
         Caption         =   "Organization Information"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeActivity 
         Caption         =   "Change Activity"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSectionInformation 
         Caption         =   "Section Information"
      End
      Begin VB.Menu mnuFileStep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuMaintenace 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuGroupInfo 
         Caption         =   "Group Information"
      End
      Begin VB.Menu mnuMonitorSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepartmentInfo 
         Caption         =   "Department Information"
      End
      Begin VB.Menu mnuMonitorSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployeeInfo 
         Caption         =   "Employee Information"
      End
      Begin VB.Menu mnuMonitorSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemarkEntry 
         Caption         =   "Remark Entry"
      End
      Begin VB.Menu mnuMonitorSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObjectInfo 
         Caption         =   "Object Information"
      End
      Begin VB.Menu mnuMonitorSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObjectPermission 
         Caption         =   "Object Permission"
      End
      Begin VB.Menu mnuMonitorStep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGroupWisePermission 
         Caption         =   "Group Wise Permission"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuOptiions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuGroupWiseReport 
         Caption         =   "Section Wise Attendant Report"
      End
      Begin VB.Menu mnuReportSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDateWiseAttendant 
         Caption         =   "Employee Wise Attendant Report"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTechSupport 
         Caption         =   "Technical Support"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo ErrHandle:
App.TaskVisible = False
'--Initialize StatusBar
    StatusBar1.Panels.Item(3) = "User: " & gUserId & "    "
    Dim rsGroup As New Recordset
    rsGroup.Open "Select vGroupName From tblGroup Where siGroupID=" & giGroupID & "", CnnMain, adOpenForwardOnly, adLockReadOnly
    StatusBar1.Panels.Item(4) = "Group: " & rsGroup!vGroupName & ""
    StatusBar1.Panels.Item(5) = Format(Now, "dd-mmm-yyyy")
    Set rsGroup = Nothing
'--Initialize Menu Icon
    Dim i As Integer
    Dim hMenu, hSubMenu, menuID, x As Variant
    hMenu = GetMenu(hwnd)
'--File Menu
    hSubMenu = GetSubMenu(hMenu, 0)
    menuID = GetMenuItemID(hSubMenu, 0)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(1).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 2)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(2).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 4)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(3).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 6)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(3).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 8)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(4).Picture, 0&)
'--Maintainance Information
    hSubMenu = GetSubMenu(hMenu, 1)
    menuID = GetMenuItemID(hSubMenu, 0)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(5).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 2)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(6).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 4)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(7).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 6)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(8).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 8)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(9).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 10)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(10).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 12)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(10).Picture, 0&)

'--Tools
'    hSubMenu = GetSubMenu(hMenu, 2)
'    menuID = GetMenuItemID(hSubMenu, 0)
'        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(11).Picture, 0&)
'--Report
    hSubMenu = GetSubMenu(hMenu, 2)
    menuID = GetMenuItemID(hSubMenu, 0)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(12).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 2)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(12).Picture, 0&)
'    menuID = GetMenuItemID(hSubMenu, 4)
'        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(12).Picture, 0&)
'    menuID = GetMenuItemID(hSubMenu, 6)
'        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(12).Picture, 0&)
'    menuID = GetMenuItemID(hSubMenu, 8)
'        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(12).Picture, 0&)
''--About
    hSubMenu = GetSubMenu(hMenu, 3)
    menuID = GetMenuItemID(hSubMenu, 0)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(13).Picture, 0&)
    menuID = GetMenuItemID(hSubMenu, 2)
        x = SetMenuItemBitmaps(hMenu, menuID, 0, ImageList2.ListImages(14).Picture, 0&)
'--Object Permission
Dim rsObject As New Recordset
Dim rsSearchObject As New Recordset
rsObject.Open "Select * From tblObject Order By vObjectName Asc", CnnMain, adOpenForwardOnly, adLockReadOnly
If Not rsObject.EOF Then
    Do While Not rsObject.EOF
        rsSearchObject.Open "Select * From tblObjectPermission Where vEmployee='" & gUserId & "' And vObjectName='" & rsObject!vObjectName & "'", CnnMain, adOpenForwardOnly, adLockReadOnly
        If Not rsSearchObject.EOF Then
            Select Case rsSearchObject!vObjectName
                Case "mnuCompanyInfo"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuCompanyInfo.Enabled = True
                        Toolbar1.Buttons.Item(1).Enabled = True
                    Else
                        mnuCompanyInfo.Enabled = False
                        Toolbar1.Buttons.Item(1).Enabled = False
                    End If
                Case "mnuChangePassword"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuChangePassword.Enabled = True
                        Toolbar1.Buttons.Item(2).Enabled = True
                    Else
                        mnuChangePassword.Enabled = False
                        Toolbar1.Buttons.Item(2).Enabled = False
                    End If
                Case "mnuChangeActivity"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuChangeActivity.Enabled = True
                        Toolbar1.Buttons.Item(3).Enabled = True
                    Else
                        mnuChangeActivity.Enabled = False
                        Toolbar1.Buttons.Item(3).Enabled = False
                    End If
                Case "mnuExit"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuExit.Enabled = True
                        Toolbar1.Buttons.Item(13).Enabled = True
                    Else
                        mnuExit.Enabled = False
                        Toolbar1.Buttons.Item(13).Enabled = False
                    End If
                Case "mnuGroupInfo"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuGroupInfo.Enabled = True
                        Toolbar1.Buttons.Item(4).Enabled = True
                    Else
                        mnuGroupInfo.Enabled = False
                        Toolbar1.Buttons.Item(4).Enabled = False
                    End If
                Case "mnuDepartmentInfo"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuDepartmentInfo.Enabled = True
                        Toolbar1.Buttons.Item(5).Enabled = True
                    Else
                        mnuDepartmentInfo.Enabled = False
                        Toolbar1.Buttons.Item(5).Enabled = False
                    End If
                Case "mnuEmployeeInfo"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuEmployeeInfo.Enabled = True
                        Toolbar1.Buttons.Item(6).Enabled = True
                    Else
                        mnuEmployeeInfo.Enabled = False
                        Toolbar1.Buttons.Item(6).Enabled = False
                    End If
                Case "mnuRemarkEntry"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuRemarkEntry.Enabled = True
                        Toolbar1.Buttons.Item(7).Enabled = True
                    Else
                        mnuRemarkEntry.Enabled = False
                        Toolbar1.Buttons.Item(7).Enabled = False
                    End If
                Case "mnuObjectInfo"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuObjectInfo.Enabled = True
                        Toolbar1.Buttons.Item(8).Enabled = True
                        
                    Else
                        mnuObjectInfo.Enabled = False
                        Toolbar1.Buttons.Item(8).Enabled = False
                    End If
                Case "mnuObjectPermission"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuObjectPermission.Enabled = True
                        Toolbar1.Buttons.Item(9).Enabled = True
                    Else
                        mnuObjectPermission.Enabled = False
                        Toolbar1.Buttons.Item(9).Enabled = False
                    End If
                Case "mnuOptions"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuOptiions.Enabled = True
                        Toolbar1.Buttons.Item(10).Enabled = True
                    Else
                        mnuOptiions.Enabled = False
                        Toolbar1.Buttons.Item(10).Enabled = False
                    End If
                
                Case "mnuDateWiseAttendant"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuDateWiseAttendant.Enabled = True
                    Else
                        mnuDateWiseAttendant.Enabled = False
                    End If
                Case "mnuTechSupport"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuTechSupport.Enabled = True
                        Toolbar1.Buttons.Item(11).Enabled = True
                    Else
                        mnuTechSupport.Enabled = False
                        Toolbar1.Buttons.Item(11).Enabled = False
                    End If
                Case "mnuAbout"
                    If rsSearchObject!vEnable = "Yes" Then
                        mnuAbout.Enabled = True
                        Toolbar1.Buttons.Item(12).Enabled = True
                    Else
                        mnuAbout.Enabled = False
                        Toolbar1.Buttons.Item(12).Enabled = False
                    End If
            End Select
        Else
            Select Case rsObject!vObjectName
                Case "mnuCompanyInfo"
                    If rsObject!vDefault = "Yes" Then
                        mnuCompanyInfo.Enabled = True
                        Toolbar1.Buttons.Item(1).Enabled = True
                    Else
                        mnuCompanyInfo.Enabled = False
                        Toolbar1.Buttons.Item(1).Enabled = False
                    End If
                Case "mnuChangePassword"
                    If rsObject!vDefault = "Yes" Then
                        mnuChangePassword.Enabled = True
                        Toolbar1.Buttons.Item(2).Enabled = True
                    Else
                        mnuChangePassword.Enabled = False
                        Toolbar1.Buttons.Item(2).Enabled = False
                    End If
                Case "mnuChangeActivity"
                    If rsObject!vDefault = "Yes" Then
                        mnuChangeActivity.Enabled = True
                        Toolbar1.Buttons.Item(3).Enabled = True
                    Else
                        mnuChangeActivity.Enabled = False
                        Toolbar1.Buttons.Item(3).Enabled = False
                    End If
                Case "mnuExit"
                    If rsObject!vDefault = "Yes" Then
                        mnuExit.Enabled = True
                        Toolbar1.Buttons.Item(13).Enabled = True
                    Else
                        mnuExit.Enabled = False
                        Toolbar1.Buttons.Item(13).Enabled = False
                    End If
                Case "mnuGroupInfo"
                    If rsObject!vDefault = "Yes" Then
                        mnuGroupInfo.Enabled = True
                        Toolbar1.Buttons.Item(4).Enabled = True
                    Else
                        mnuGroupInfo.Enabled = False
                        Toolbar1.Buttons.Item(4).Enabled = False
                    End If
                Case "mnuDepartmentInfo"
                    If rsObject!vDefault = "Yes" Then
                        mnuDepartmentInfo.Enabled = True
                        Toolbar1.Buttons.Item(5).Enabled = True
                    Else
                        mnuDepartmentInfo.Enabled = False
                        Toolbar1.Buttons.Item(5).Enabled = False
                    End If
                Case "mnuEmployeeInfo"
                    If rsObject!vDefault = "Yes" Then
                        mnuEmployeeInfo.Enabled = True
                        Toolbar1.Buttons.Item(6).Enabled = True
                    Else
                        mnuEmployeeInfo.Enabled = False
                        Toolbar1.Buttons.Item(6).Enabled = False
                    End If
                Case "mnuRemarkEntry"
                    If rsObject!vDefault = "Yes" Then
                        mnuRemarkEntry.Enabled = True
                        Toolbar1.Buttons.Item(7).Enabled = True
                    Else
                        mnuRemarkEntry.Enabled = False
                        Toolbar1.Buttons.Item(7).Enabled = False
                    End If
                Case "mnuObjectInfo"
                    If rsObject!vDefault = "Yes" Then
                        mnuObjectInfo.Enabled = True
                        Toolbar1.Buttons.Item(8).Enabled = True
                        
                    Else
                        mnuObjectInfo.Enabled = False
                        Toolbar1.Buttons.Item(8).Enabled = False
                    End If
                Case "mnuObjectPermission"
                    If rsObject!vDefault = "Yes" Then
                        mnuObjectPermission.Enabled = True
                        Toolbar1.Buttons.Item(9).Enabled = True
                    Else
                        mnuObjectPermission.Enabled = False
                        Toolbar1.Buttons.Item(9).Enabled = False
                    End If
                Case "mnuOptions"
                    If rsObject!vDefault = "Yes" Then
                        mnuOptiions.Enabled = True
                        Toolbar1.Buttons.Item(10).Enabled = True
                    Else
                        mnuOptiions.Enabled = False
                        Toolbar1.Buttons.Item(10).Enabled = False
                    End If
                
                Case "mnuDateWiseAttendant"
                    If rsObject!vDefault = "Yes" Then
                        mnuDateWiseAttendant.Enabled = True
                    Else
                        mnuDateWiseAttendant.Enabled = False
                    End If
                
                Case "mnuTechSupport"
                    If rsObject!vDefault = "Yes" Then
                        mnuTechSupport.Enabled = True
                        Toolbar1.Buttons.Item(11).Enabled = True
                    Else
                        mnuTechSupport.Enabled = False
                        Toolbar1.Buttons.Item(11).Enabled = False
                    End If
                Case "mnuAbout"
                    If rsObject!vDefault = "Yes" Then
                        mnuAbout.Enabled = True
                        Toolbar1.Buttons.Item(12).Enabled = True
                    Else
                        mnuAbout.Enabled = False
                        Toolbar1.Buttons.Item(12).Enabled = False
                    End If
            End Select

        End If
        rsSearchObject.Close
        rsObject.MoveNext
    Loop
End If
Set rsSearchObject = Nothing
Set rsObject = Nothing
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrHandle:
Me.MousePointer = vbHourglass
If Cancel = 0 Then
    Dim intResponse As Integer
    intResponse = MsgBox("Are you sure you want to exit the Application?", vbYesNo + vbQuestion, "Confirmation")
    If intResponse = vbYes Then
        'End
        gFlag = False
        CnnMain.Close
        Unload Me
        
   
    Else
        Cancel = 1
        Me.MousePointer = vbDefault
    End If
End If
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuChangeActivity_Click()
    frmUpdateAttendance.Show
End Sub

Private Sub mnuChangePassword_Click()
    frmChangePassword.Show
End Sub

Private Sub mnuCompanyInfo_Click()
    frmComInfo.Show
End Sub

Private Sub mnuClientWiseTransInfo_Click()
    'rptfrmClientTrans.Show
End Sub

Private Sub mnuDateWiseAttendant_Click()
    rptAttendance.Show
End Sub

Private Sub mnuDepartmentInfo_Click()
    frmDepartMent.Show
End Sub

Private Sub mnuEmployeeInfo_Click()
    frmEmpDetails.Show
End Sub

Private Sub mnuExit_Click()
Dim frm As Form
For Each frm In Forms
    Unload frm
    Set frm = Nothing
    gFlag = False
    Unload Me
Next
End Sub

Private Sub mnuObject_Click()
    frmDefaultMenu.Show vbModal
End Sub

Private Sub mnuGroupInfo_Click()
    frmGroup.Show
End Sub

Private Sub mnuGroupWisePermission_Click()
    frmGroupWiseObjectPermission.Show
End Sub

Private Sub mnuGroupWiseReport_Click()
    rptSectionReport.Show
End Sub

Private Sub mnuObjectInfo_Click()
    frmDefaultMenu.Show
End Sub

Private Sub mnuObjectPermission_Click()
    frmUserPermission.Show
End Sub

Private Sub mnuOptiions_Click()
    MsgBox "Under Construction! Please try it later.", vbExclamation, "Remark Message"
End Sub

Private Sub mnuRemarkEntry_Click()
    frmEntryRemark.Show
End Sub

Private Sub mnuSectionInformation_Click()
    frmSection.Show
End Sub

Private Sub mnuTechSupport_Click()
    Dim strHelp As String
    strHelp = App.Path & "\Help\Main.htm"
    Module2.ShellExecute hwnd, "open", strHelp, vbNullString, vbNullString, 2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandle:
Select Case Button.Key
Case "CompanyInfo"
    frmComInfo.Show
    
Case "ChangePassword"
    frmChangePassword.Show
    
Case "ChangeActivity"
    frmUpdateAttendance.Show

Case "GroupInfo"
    frmGroup.Show

Case "DepartmentInfo"
    frmDepartMent.Show

Case "EmployeeInfo"
    frmEmpDetails.Show
    
Case "RemarkInfo"
    frmEntryRemark.Show
    
Case "ObjectInfo"
    frmDefaultMenu.Show

Case "ObjectPermission"
    frmUserPermission.Show

Case "Options"
    MsgBox "Under Construction! Please try it later.", vbExclamation, "Remark Message"

Case "TechSupport"
    MsgBox "Under Construction! Please try it later.", vbExclamation, "Remark Message"

Case "About"
    frmAbout.Show

Case "Exit"
    Unload Me

End Select
Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
End Sub
