Attribute VB_Name = "Module1"
Public u_paw As String
Public grade As String
Public Function chaxun(sqlstr As String)
  Dim cn As New ADODB.Connection
  Dim cn_str As String
  cn_str = "driver=sql server;server=(local);database=da1"
  cn.Open cn_str
  Dim rs As New ADODB.Recordset
  rs.Open sqlstr, cn, adOpenDynamic, adLockOptimistic
  Set chaxun = rs
End Function

 
