Attribute VB_Name = "Module2"
Option Explicit
Public gFlag As Boolean
Public gLevel As Integer
Public giGroupID As Integer
Public gUserId As String
Public gPassword As String
Public gstrMachineName As String
Public vgCompanyName As String
'--For New Zone/Area/SubArea Button Click
Public gZoneClick As Integer
Public gAreaClick As Integer
Public gSubAreaClick As Integer
Public gDeptClick As Integer
'----
Public gStrAccNo As String
Public gGroup As String
Public gPinCheck As String
Public gCrime As Integer
Public gNewClientId As String
'----For Client Registration
Public vClientId As String
Public Password As String
Public vSrNo As String
Public mBalance As String
Public mTotalDebit As String
Public mTotalCredit As String
'-----Search for Default Web Browser------------------
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
'---------Api For Computer Name
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'-------Auto Drop Down Combo Box
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'--Add Jpg Menu
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Sub prc_DisAllow(frm As Form)
Dim obj As Object
  For Each obj In frm
        If TypeOf obj Is CommandButton Then
                If obj.Name = "cmdSave" Or obj.Name = "cmdCancel" Then
                obj.Enabled = False
        Else
                obj.Enabled = True
                End If
        End If
    Next
End Sub
