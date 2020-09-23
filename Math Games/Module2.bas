Attribute VB_Name = "Module2"
Option Explicit
Public KidsName As String
Public Z As Integer
Public varPoints As Double
Public varChk As Boolean
Public rs As DAO.Recordset
Public dbs As DAO.Database
Public usrId As Integer
Public Sub ConnectAccessDb()
    
    Set dbs = OpenDatabase(App.Path & "\Animation\Score.mdb", False, False, ";Pwd=zohaib?308")
    
End Sub
Public Sub CloseAccessDb()
  
    rs.Close
    dbs.Close
    Set rs = Nothing
    Set dbs = Nothing

End Sub

