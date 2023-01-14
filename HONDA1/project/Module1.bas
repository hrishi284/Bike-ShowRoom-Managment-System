Attribute VB_Name = "Module1"
Option Explicit
Public rs2 As New Recordset
Public rs1 As New Recordset
Public rs As New Recordset
Public con As ADODB.Connection
Public nm As Long
Public str As String
Public DataPath As String
Public DPath As String
Public DataP As String
Public Sub getconnected()
DataPath = "D:\BIke\"
DPath = "D:\Custmer\"
DataP = "D:\Employee\"
Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\bike.mdb;"
con.Open

End Sub

Public Sub retdata(ByVal q As String)
If rs1.State = 1 Then
rs1.Close
End If

Call getconnected
rs1.CursorLocation = adUseClient
rs1.Open q, con, adOpenDynamic, adLockOptimistic

End Sub

Public Sub inupdel(ByVal q As String)
If rs.State = 1 Then
rs.Close
End If
Call getconnected
rs.Open q, con, adOpenDynamic, adLockOptimistic


End Sub

Public Sub numval(ByVal q As String)
If rs.State = 1 Then
rs.Close
End If
rs.Open q, con, adOpenDynamic, adLockOptimistic
If Not rs.EOF Or Not rs.BOF Then
nm = rs.Fields(0).Value
End If
End Sub

Public Sub dataret(ByVal q As String)
If rs2.State = 1 Then
rs2.Close
End If

Call getconnected
rs2.CursorLocation = adUseClient
rs2.Open q, con, adOpenDynamic, adLockOptimistic

End Sub


