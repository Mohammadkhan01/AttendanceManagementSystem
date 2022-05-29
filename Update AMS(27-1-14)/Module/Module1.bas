Attribute VB_Name = "Module1"
Option Explicit

Public CnnMain As ADODB.Connection

Public gstrServName As String
'Public gstrDBName As String
'Public gstrDBUID As String
Public gstrDBUpass As String

Public strfindcode As String
Public strfindname As String

Public gblnConnErr As Boolean 'This is for Database connection in SQL Server
Public gblnConnFromMnu As Boolean 'This is for Database connection in SQL Server

Public Sub opendata()
'On Error GoTo Jail
gblnConnErr = False
Set CnnMain = New ADODB.Connection
CnnMain.CursorLocation = adUseClient
'CnnMain.Provider = "SQLOLEDB"
'CnnMain.ConnectionString = "driver={SQL Server};" & _
       "server=" & gstrServName & ";uid=" & gstrDBUID & ";pwd=" & gstrDBUpass & ";database=" & gstrDBName
CnnMain.ConnectionString = "Data Source=" & App.Path & "\Murphy.mdb; jet oledb:Database Password=" & gstrDBUpass & ""
CnnMain.ConnectionTimeout = 30
CnnMain.Open
Exit Sub

'Jail:
'If Err.Number = -2147467259 Then
'    MsgBox Err.Description & vbCrLf & "Please contact your Network Administrator or Database Administrator.", vbExclamation + vbOKOnly, "Connection Failed"
'    gblnConnErr = True
'ElseIf Err.Number = -2147217843 Then
'    MsgBox Err.Description & vbCrLf & "Please contact your Database Administrator.", vbExclamation + vbOKOnly, "Connection Failed"
'    gblnConnErr = True
'Else
'        MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Connection Failed"
'        Err.Clear
'        End
'End If

End Sub

Public Sub Main()
Get_Data_From_Registry
If Connection_Info = False Then
    frmDBConnInfo.Show
    Exit Sub
End If
opendata
Dim msgResult As VbMsgBoxResult
If gblnConnErr = True Then
    msgResult = MsgBox("Would you connect to the Database again?", vbYesNo, "Message")
    If msgResult = vbYes Then
        frmDBConnInfo.Caption = "Try to connect again"
        frmDBConnInfo.Show
        frmDBConnInfo.txtUPassword.SetFocus
    Else
        If gblnConnFromMnu = False Then
            End
        End If
    End If
    Exit Sub
End If
frmSplash.Show
End Sub

Public Function VerifyDate(strtmp As String) As Boolean
VerifyDate = True
If Not IsDate(strtmp) Then
    MsgBox "Invalid Date", vbInformation, "Sorry!"
    VerifyDate = False
    Exit Function
Else
    If Not Len(strtmp) = 10 Then
        MsgBox "Please Enter Date this format" & vbCrLf _
                & "dd/mm/yyyy", vbInformation, "Sorry!"
                VerifyDate = False
                Exit Function
    ElseIf Val(Mid(strtmp, 4, 2)) > 12 Then
        MsgBox "Please Enter Date this format" & vbCrLf _
                & "dd/mm/yyyy", vbInformation, "Sorry!"
        VerifyDate = False
        Exit Function
    End If
End If
End Function
Private Function Connection_Info() As Boolean
Connection_Info = True

End Function

Private Sub Get_Data_From_Registry()
'gstrServName = GetSetting(App.EXEName, "Section1", "Server_Name")
'gstrDBUID = Encode(GetSetting(App.EXEName, "Section1", "User_ID"))
gstrDBUpass = Encode(GetSetting(App.EXEName, "Section1", "User_Password"))
'gstrDBName = GetSetting(App.EXEName, "Section1", "DB_Name")
End Sub

Public Function Encode(cString) As String
Dim lngAsc As Long
Dim strEncode As String
Dim i As Integer
strEncode = ""
For i = 1 To Len(cString)
    lngAsc = Asc(Mid(cString, i, 1)) - (i + 19)
    strEncode = strEncode + Chr$(lngAsc)
Next
Encode = strEncode
End Function
