'NOTE: You can use the "Rename" command on the context menu to change the class name "Service1" in both code and config file together.
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.IO.Compression
 
 
Public Class Service1
  Implements IService1

  Private MyPassword As String = "Your Password here"
  Private FilePath As String = "Your File Path in the server to write log of transactions"
    
  Private Dim _DBname As String = "Your DataBase"
  Private Dim _server As String = "localhost"
  Private Dim pswd As String = "Your Password"

  Private Dim _port As String = "3306"
  Private Dim user As String = "Your user"
  Private Dim _timeout As UInt16 = 30000
  
  
    ReadOnly Property ConnectionString() As String
        Get
            Return String.Format("server={0}; port={1};user id={2}; password={3}; database={4};Connection timeout={5};Charset=utf8; pooling=false; convert zero datetime=True", 
                                 _server, _port, user, pswd, _DBname, _timeout)
        End Get
   '******************************************************************************************************************************
   '******************************************************************************************************************************
    
    Public Function SelectSql(ByVal sqlstr As String, ByVal FieldNo As String) As CompositeType Implements IService1.SelectSql

        Dim Command As MySqlCommand = Nothing
        Dim sqlReader As MySqlDataReader = Nothing
        Dim ArrayValues As CompositeType = New CompositeType
        Dim MyList As New List(Of String)()
        Dim conn As MySqlConnection = Nothing
        Dim index As Integer
        Try
            Dim wrapper As New MyEncryption(MyPassword)
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + sqlstr + vbCrLf, True)

            sqlstr = wrapper.DecryptData(sqlstr)
            ' My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + sqlstr + vbCrLf, True)
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + ConnectionString + vbCrLf, True)


            conn = New MySqlConnection
            conn.ConnectionString = ConnectionString
            conn.Open()

            Command = conn.CreateCommand()
            Command.CommandText = sqlstr
            sqlReader = Command.ExecuteReader()
            ' My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", sqlstr + vbCrLf, True)
            'If FieldNo = "1" Then
            '    My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "FieldNo:1" + vbCrLf, True)
            'End If
            index = Convert.ToInt32(FieldNo)
            If sqlReader.HasRows Then
                While sqlReader.Read() 'And counter < 100
                    'My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "FieldNo:" + FieldNo + " type(FieldNo:)" + FieldNo.GetType().ToString + vbCrLf, True)
                    'My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "1:Mylist.count" + MyList.Count.ToString + vbCrLf, True)

                    For i As Integer = 0 To index - 1


                        If Not IsDBNull(sqlReader(i)) Then
                            MyList.Add(sqlReader(i))
                        Else
                            MyList.Add("")
                        End If


                        '   My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "sqlreader:" + sqlReader(i) + vbCrLf, True)
                    Next
                End While
            End If


            '   My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "2¨:Mylist.count" + MyList.Count.ToString + vbCrLf, True)

            If MyList.Count > 0 Then

                ArrayValues.StrResults = MyList.ToArray()
                '  My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "ArrayValues.StrResults.Length:" + ArrayValues.StrResults.Length.ToString + vbCrLf, True)

            End If

        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Exception:" + ex.Message + vbCrLf, True)

            ' No stacktrace to client, just message...
            Throw New FaultException(ex.Message)
        Finally
            If Not sqlReader Is Nothing Then
                sqlReader.Close()
                sqlReader.Dispose()
            End If
            If Not Command Is Nothing Then Command.Dispose()
            If Not conn Is Nothing Then
                conn.Close()
                conn.Dispose()
            End If

        End Try

        Return ArrayValues

    End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
   
    Public Function SelectScalar(ByVal sqlstr As String) As Integer Implements IService1.SelectScalar
        Dim result As Integer = -1
        Dim conn As MySqlConnection = Nothing
        Dim Command As MySqlCommand = Nothing
        Try
            Dim wrapper As New MyEncryption(MyPassword)
            sqlstr = wrapper.DecryptData(sqlstr)
            conn = New MySqlConnection
            conn.ConnectionString = ConnectionString
            conn.Open()
            Command = conn.CreateCommand()
            Command.CommandText = sqlstr
            result = Convert.ToInt32(Command.ExecuteScalar())

        Catch ex As Exception

            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Exception:" + ex.Message + vbCrLf, True)
            Throw New FaultException(ex.Message)
        Finally
            If Not conn Is Nothing Then
                conn.Close()
                conn.Dispose()
            End If

            If Not Command Is Nothing Then Command.Dispose()

        End Try



        Return result
    End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
   
    Public Function UpdateSql(ByVal sqlstr As String) As Integer Implements IService1.UpdateSql
        Dim result As Integer = ExecuteNonQueryCommand(sqlstr)

        Return result
    End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
   
    Public Function DeleteSql(ByVal sqlstr As String) As Integer Implements IService1.DeleteSql
        Dim result As Integer = ExecuteNonQueryCommand(sqlstr)

        Return result
    End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
   
    Public Function InsertSql(ByVal sqlstr As String) As Integer Implements IService1.InsertSql
        Dim result As Integer = ExecuteNonQueryCommand(sqlstr)

        Return result
    End Function
      
   '******************************************************************************************************************************
   '******************************************************************************************************************************
   
    Public Function TableExists(ByVal TableName As String) As Boolean Implements IService1.TableExists
        Dim bresult As Boolean = False
        Dim conn As MySqlConnection = Nothing
        Dim Command As MySqlCommand = Nothing

        Dim sql_st As String = "show tables like '" + TableName + "'"

        Try
            Dim wrapper As New MyEncryption(MyPassword)
            TableName = wrapper.DecryptData(TableName)

            conn = New MySqlConnection
            conn.ConnectionString = ConnectionString
            conn.Open()

            Command = conn.CreateCommand()
            Command.CommandText = sql_st
            If Convert.ToInt32(Command.ExecuteScalar()) > 0 Then
                bresult = True
            End If


        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Exception:" + ex.Message + vbCrLf, True)
            Throw New FaultException(ex.Message)
        Finally
            If Command IsNot Nothing Then
                Command.Dispose()
            End If
            If Not conn Is Nothing Then
                conn.Close()
                conn.Dispose()
            End If
        End Try


        Return bresult
    End Function
      
   '******************************************************************************************************************************
   '******************************************************************************************************************************
    
    Private Function FindNextID_internal(ByVal TableName As String, ByVal par As String, ByRef conn As MySqlConnection) As Integer
        Dim iresult As Integer = -1
        Dim sql_st As String
        Dim obj As Object = Nothing


        Dim Command As MySqlCommand = Nothing
        Try


            sql_st = "select max(" + par + ") from " + TableName
            Command = conn.CreateCommand()
            Command.CommandText = sql_st
            obj = Command.ExecuteScalar


            If IsDBNull(obj) Then
                iresult = 0
            Else
                iresult = Convert.ToInt32(obj) + 1
            End If

        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Exception:" + ex.Message + vbCrLf, True)
            Throw New FaultException(ex.Message)
        Finally
            If Command IsNot Nothing Then
                Command.Dispose()
            End If

        End Try

        Return iresult
    End Function
      
   '******************************************************************************************************************************
   '******************************************************************************************************************************
      
    Public Function FindNextID(ByVal TableName As String, ByVal par As String) As Integer Implements IService1.FindNextID
        Dim iresult As Integer = -1
        Dim conn As MySqlConnection = Nothing
        Dim Command As MySqlCommand = Nothing
        Try

            Dim wrapper As New MyEncryption(MyPassword)
            TableName = wrapper.DecryptData(TableName)
            conn = New MySqlConnection
            conn.ConnectionString = ConnectionString
            conn.Open()

            iresult = FindNextID_internal(TableName, par, conn)


        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Exception:" + ex.Message + vbCrLf, True)
            Throw New FaultException(ex.Message)
        Finally
            If Command IsNot Nothing Then
                Command.Dispose()
            End If

            If Not conn Is Nothing Then
                conn.Close()
                conn.Dispose()
            End If
        End Try


        Return iresult
    End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
   
    Public Function CloneTableRow(ByVal TableName As String, ByVal ID As Integer, ByRef NewID As Integer, ByVal ArrayParameter() As String, ByVal InitialValue() As String) As Boolean Implements IService1.CloneTableRow
        Dim bresult As Boolean = False
        Dim sql_st As String = Nothing
        Dim temptable As String = "temp"
        Dim counter As Integer = 1
        NewID = 0

        Dim conn As MySqlConnection = Nothing
        Dim Command As MySqlCommand = Nothing

        ' My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Entry in CloneTableRow function" + vbCrLf, True)

        Try
            Dim wrapper As New MyEncryption(MyPassword)
            TableName = wrapper.DecryptData(TableName)

            conn = New MySqlConnection
            conn.ConnectionString = ConnectionString
            conn.Open()
            'While TableExists(temptable, MyPassword) = vbTrue
            '    temptable += counter.ToString
            '    counter += 1
            'End While


            ' My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Passwd OK! Connection OK!" + vbCrLf, True)


            sql_st = "CREATE TEMPORARY TABLE " + temptable + "  Select * FROM " + TableName + " WHERE ID='" + ID.ToString + "'"
            Command = conn.CreateCommand()
            Command.CommandText = sql_st
            Command.ExecuteNonQuery()
            '  My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " " + sql_st + vbCrLf, True)

            NewID = FindNextID_internal(TableName, "ID", conn)
            ' My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " New ID: " + NewID.ToString + vbCrLf, True)

            sql_st = "update " + temptable + " set ID='" + NewID.ToString + "'"

            For i As Integer = 0 To ArrayParameter.Count - 1
                sql_st += ", " + ArrayParameter(i) + "='" + InitialValue(i) + "'"
            Next

            Command = conn.CreateCommand()
            Command.CommandText = sql_st
            Command.ExecuteNonQuery()

            ' My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " " + sql_st + vbCrLf, True)

            sql_st = "insert into " + TableName + " select * from " + temptable
            Command = conn.CreateCommand()
            Command.CommandText = sql_st
            Command.ExecuteNonQuery()

            '  My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " " + sql_st + vbCrLf, True)

            bresult = True


        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Exception:" + ex.Message + vbCrLf, True)
            Throw New FaultException(ex.Message)
        Finally
            If Command IsNot Nothing Then
                Command.Dispose()
            End If

            If Not conn Is Nothing Then
                conn.Close()
                conn.Dispose()
            End If

        End Try
        Return bresult
    End Function
    
   '******************************************************************************************************************************
   '******************************************************************************************************************************
   
    Public Function CloneTableRowWithoutInit(ByVal TableName As String, ByVal ID As Integer, ByRef NewID As Integer) As Boolean Implements IService1.CloneTableRowWithoutInit
        Dim bresult As Boolean = False
        Dim sql_st As String = Nothing
        Dim temptable As String = "temp"
        Dim counter As Integer = 1
        NewID = 0

        Dim conn As MySqlConnection = Nothing
        Dim Command As MySqlCommand = Nothing

        ' My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Entry in function" + vbCrLf, True)

        Try
            Dim wrapper As New MyEncryption(MyPassword)
            TableName = wrapper.DecryptData(TableName)
            conn = New MySqlConnection
            conn.ConnectionString = ConnectionString
            conn.Open()
            'While TableExists(temptable, MyPassword) = vbTrue
            '    temptable += counter.ToString
            '    counter += 1
            'End While


            '  My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Passwd OK! Connection OK!" + vbCrLf, True)


            sql_st = "CREATE TEMPORARY TABLE " + temptable + "  Select * FROM " + TableName + " WHERE ID='" + ID.ToString + "'"
            Command = conn.CreateCommand()
            Command.CommandText = sql_st
            Command.ExecuteNonQuery()
            '   My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " " + sql_st + vbCrLf, True)

            NewID = FindNextID_internal(TableName, "ID", conn)
            '  My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " New ID: " + NewID.ToString + vbCrLf, True)

            sql_st = "update " + temptable + " set ID='" + NewID.ToString + "'"
            Command = conn.CreateCommand()
            Command.CommandText = sql_st
            Command.ExecuteNonQuery()

            '       My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " " + sql_st + vbCrLf, True)

            sql_st = "insert into " + TableName + " select * from " + temptable
            Command = conn.CreateCommand()
            Command.CommandText = sql_st
            Command.ExecuteNonQuery()

            '   My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " " + sql_st + vbCrLf, True)

            bresult = True

        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Exception:" + ex.Message + vbCrLf, True)
            Throw New FaultException(ex.Message)
        Finally
            If Command IsNot Nothing Then
                Command.Dispose()
            End If

            If Not conn Is Nothing Then
                conn.Close()
                conn.Dispose()
            End If

        End Try


        Return bresult
    End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
   
    Public Function GetDataInZIPFile(ByVal sqlstr As String, ByVal FieldNo As String, ByVal TextFile As String) As Boolean Implements IService1.GetDataInZIPFile
        Dim result As Boolean = True
        Dim Command As MySqlCommand = Nothing
        Dim sqlReader As MySqlDataReader = Nothing
        Dim conn As MySqlConnection = Nothing
        Dim index As Integer = Convert.ToInt32(FieldNo)
        Dim Filename As String = "\" + TextFile + ".txt"
        Try

            Dim wrapper As New MyEncryption(MyPassword)
            sqlstr = wrapper.DecryptData(sqlstr)
            conn = New MySqlConnection
            conn.ConnectionString = ConnectionString
            conn.Open()


            Command = conn.CreateCommand()
            Command.CommandText = sqlstr
            sqlReader = Command.ExecuteReader()

            If My.Computer.FileSystem.FileExists(FilePath + Filename) Then
                My.Computer.FileSystem.DeleteFile(FilePath + Filename)
                My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "FILE DELETED " & FilePath + Filename + vbCrLf, True)
            Else
                My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "FILE DOESN'T DELETED " & FilePath + Filename + vbCrLf, True)

            End If

            Dim row As String
            If sqlReader.HasRows Then
                While sqlReader.Read()

                    row = ""
                    For i As Integer = 0 To index - 1
                        '         My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "i=" + i.ToString + " " + sqlReader(i).ToString() + "$" + vbCrLf, True)
                        row += sqlReader(i).ToString() + "$"
                    Next
                    '   My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", "row=" + row.ToString() + "$" + vbCrLf, True)

                    row += vbCrLf
                    My.Computer.FileSystem.WriteAllText(FilePath + TextFile + ".txt", row, True)

                End While
            End If

            Dim fileToBeDeflateZipped As FileInfo = New FileInfo(FilePath + TextFile + ".txt")
            Dim deflateZipFileName As FileInfo = New FileInfo(String.Concat(fileToBeDeflateZipped.FullName, ".zip"))
            Using fileToBeZippedAsStream As FileStream = fileToBeDeflateZipped.OpenRead()
                Using deflateZipTargetAsStream As FileStream = deflateZipFileName.Create()
                    Using deflateZipStream As DeflateStream = New DeflateStream(deflateZipTargetAsStream, CompressionMode.Compress)

                        fileToBeZippedAsStream.CopyTo(deflateZipStream)
                    End Using
                End Using
            End Using


        Catch ex As Exception
            result = False
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " Exception:" + ex.Message + vbCrLf, True)
            Throw New FaultException(ex.Message)
        Finally
            If Not sqlReader Is Nothing Then
                sqlReader.Close()
                sqlReader.Dispose()
            End If
            If Not Command Is Nothing Then Command.Dispose()
            If Not conn Is Nothing Then
                conn.Close()
                conn.Dispose()
            End If
        End Try


        Return result
    End Function
    
   '******************************************************************************************************************************
   '******************************************************************************************************************************
   
    Private Function ExecuteNonQueryCommand(ByVal sqlstr As String) As Integer
        Dim result As Integer = 0
        Dim conn As MySqlConnection = Nothing
        Dim Command As MySqlCommand = Nothing
        Try
            Dim wrapper As New MyEncryption(MyPassword)
            sqlstr = wrapper.DecryptData(sqlstr)

            conn = New MySqlConnection
            conn.ConnectionString = ConnectionString
            conn.Open()

            Command = conn.CreateCommand()
            Command.CommandText = sqlstr
            result = Command.ExecuteNonQuery()


        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FilePath + "LogFile.txt", Now().ToString + " 5)Exception:" + ex.Message + vbCrLf, True)
            Throw New FaultException(ex.Message)
        Finally
            If Not conn Is Nothing Then
                conn.Close()
                conn.Dispose()
            End If
            If Not Command Is Nothing Then Command.Dispose()
        End Try

        Return result
    End Function
End Class
