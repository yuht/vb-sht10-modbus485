Attribute VB_Name = "DB_database"
Option Explicit

'需要引用 Microsoft ADO Ext x.x for DDL and Security 这个是ADOX
'需要引用 Microsoft ActiveX Data Objects XXX Library 这个是ADODB类库 用来操作ACCESS数据库或者SQLSERVER数据库

Dim Cat As New ADOX.Catalog '不用cat用另外一个名字也可以
Dim Conn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Public DBpath_name As String  '数据库参数
'

Public Function DB_CreateDataBase(DataBase_Path_And_Name As String) As Boolean
    On Error Resume Next
    Debug.Print "create DB"
    DBpath_name = "provider=Microsoft.Jet.OLEDB.4.0;data source=" & DataBase_Path_And_Name & ";" '创建数据库
    'Debug.Print DBpath_name
    
    Cat.Create DBpath_name '    创建数据库
    
    If Err And Err <> &H80040E17 Then  '数据库 DataBase_Path_And_Name 已经存在。
        
        Debug.Print Err.Number, Hex(Err.Number)
        Debug.Print Err.Description
        Err.Clear
        DB_CreateDataBase = False
        Exit Function
    End If
    
    DB_CreateDataBase = True
End Function
'

Public Function DB_Create_Table(DB_TableName As Table) As Boolean
    On Error Resume Next
    Debug.Print "create table"
    Cat.ActiveConnection = DBpath_name
    Cat.Tables.Append DB_TableName '建立数据表
    
    If Err And Err <> &H80040E3F Then   '表 DB_TableName 已存在。
        Debug.Print Err.Number, Hex(Err.Number)
        Debug.Print Err.Description
        Err.Clear
        DB_Create_Table = False
        Exit Function
    End If
    DB_Create_Table = True
End Function
'

Public Function DB_InsertRecord(Table_Name As String, RecordArray() As String) As Boolean
    On Error Resume Next
    Dim i As Integer
    Conn.Open DBpath_name
    Rs.CursorLocation = adUseClient
    Rs.Open Table_Name, Conn, adOpenKeyset, adLockPessimistic
    Rs.AddNew '往表中添加新记录
    For i = 0 To UBound(RecordArray)
        Rs.Fields(i).Value = RecordArray(i)
    Next
    Debug.Print "rs.RecordCount ", Rs.RecordCount
    Rs.Update
    Rs.Close
    Conn.Close
    
    If Err Then
        Debug.Print Err.Number
        Debug.Print Err.Description
        Err.Clear
    End If
    
End Function
 
