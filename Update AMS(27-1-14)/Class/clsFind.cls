VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public HeaderCaption1 As String
Public HeaderCaption2 As String
Public ColumnValue1 As String
Public ColumnValue2 As String

Public Sub grdinitialize(CtlFlex As MSFlexGrid, RsTemp1 As ADODB.Recordset, Col1 As String, Col2 As String, row1 As Integer, row2 As Integer)
On Error GoTo A:

Dim i As Integer
i = 1

frmQuery.flex.ColWidth(0) = 1000
frmQuery.flex.ColWidth(1) = 3200

frmQuery.flex.Rows = RsTemp1.RecordCount + 1

frmQuery.flex.TextMatrix(0, 0) = Col1
frmQuery.flex.TextMatrix(0, 1) = Col2


With frmQuery.flex
    Do While Not RsTemp1.EOF
            .TextMatrix(i, 0) = RsTemp1(row1) & ""
            .TextMatrix(i, 1) = RsTemp1(row2) & ""
    RsTemp1.MoveNext
    i = i + 1
    Loop
End With
'---
Exit Sub
A:
    MsgBox "ID : " & Err.Number & Chr(10) & "Description : " & Err.Description, vbExclamation, "Call Database Administrator"
    Err.Clear
    Screen.MousePointer = vbDefault
End Sub
