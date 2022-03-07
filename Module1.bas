Attribute VB_Name = "Module1"
Option Explicit
Public db As Connection
Public consulta As Recordset
Public pri As String
Public sParametros(1 To 10) As String

'Public Sub abrirbase()
'
'       If db.State = adStateOpen Then
'        db.Close
'    End If
'    db.CursorLocation = adUseClient
'    db.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + pri + ""
'
'End Sub
