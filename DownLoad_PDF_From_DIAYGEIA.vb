Imports System.IO
Imports MySql.Data.MySqlClient

Public Class DownLoadPDFfromDIAYGEIA
    Private m_DataDir As String
    Private m_DateTo As String
    Private Sub CVS_Reader(ByRef CVSFile As String)


        Dim str, filename, sql_st As String
        Dim ADA, DATE_ANARTISI, THEMA, FOREAS_ID, FOREAS, TYPOS_PRAXIS_ID, TYPOS_PRAXIS, URL, PROTOKOLLO, DATE_PRAXIS, EKDOUSA_ARXH_ID As String
        Dim line, fixed_line, str_sql As String
        Dim index As Integer
        Dim counter As Integer = 0
        Dim total_lines As Integer = 0
        Dim MyText() As String = {""}
        ' Read and display the lines from the file until the end 
        ' of the file is reached.

        Dim sr As StreamReader = New StreamReader(m_DataDir + CVSFile, System.Text.Encoding.Default, True)
        Dim maxID As Integer = 0
        Dim maxDir_ID As Integer = 0
        '  Dim ADA_List As List(Of String) = Nothing
        Try

            Dim dbio As DataBaseIO = New DataBaseIO

            sql_st = "select max(ID) from apofaseis"
            maxID = dbio.SelectScalar(sql_st) + 1

            sql_st = "select max(Dir_ID) from apofaseis"
            maxDir_ID = dbio.SelectScalar(sql_st) + 1

            While Not sr.EndOfStream

                index = -1
                fixed_line = ""
                line = ""

                While index < 0
                    line = sr.ReadLine()
                    If line <> Nothing And line.Length > 10 Then
                        total_lines += 1
                        index = line.LastIndexOf("@") 'PUT THIS CHARACTER AT THE END OF EACH LINE
                        If index > 0 Then
                            line = line.Substring(0, index - 1).Trim
                        End If
                        fixed_line += line
                    Else
                        index = 9999
                    End If
                End While

                If index <> 9999 Then
                    MyText(total_lines - 1) = fixed_line
                    ReDim Preserve MyText(MyText.Length)
                Else
                    total_lines -= 1
                    If total_lines < 0 Then total_lines = 0
                End If
            End While

            If MyText.Length > 1 Then
                ReDim Preserve MyText(MyText.Length - 2)
            End If

            sr.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



        counter = 1
        Dim i As Integer = 0
        Dim records As Integer = 0
        Try
            For i = 0 To MyText.Length - 1

                line = MyText(i)

                line = line.Replace(";;", ";")
                EKDOUSA_ARXH_ID = CVS_lineparser(line)

                DATE_PRAXIS = CVS_lineparser(line)
                index = DATE_PRAXIS.LastIndexOf(" ")
                If index > 0 Then
                    DATE_PRAXIS = DATE_PRAXIS.Substring(0, index).Trim()
                    Dim dtFrom As String() = DATE_PRAXIS.Split(New Char() {"/"c})
                    If dtFrom.Length = 3 Then
                        DATE_PRAXIS = dtFrom(2) & "-" & dtFrom(1) & "-" & dtFrom(0)
                    End If

                End If

                PROTOKOLLO = CVS_lineparser(line)
                URL = CVS_lineparser(line)

                TYPOS_PRAXIS = CVS_lineparser(line)
                TYPOS_PRAXIS_ID = CVS_lineparser(line)
                FOREAS = CVS_lineparser(line)
                FOREAS_ID = CVS_lineparser(line)
                THEMA = CVS_lineparser(line)
                THEMA = THEMA.Replace("'", " ")

                DATE_ANARTISI = CVS_lineparser(line)
                index = DATE_ANARTISI.LastIndexOf(" ")

                If index > 0 Then
                    DATE_ANARTISI = DATE_ANARTISI.Substring(0, index).Trim()
                    Dim dtFrom As String() = DATE_ANARTISI.Split(New Char() {"/"c})
                    If dtFrom.Length = 3 Then
                        DATE_ANARTISI = dtFrom(2) & "-" & dtFrom(1) & "-" & dtFrom(0)
                    End If
                End If

                ADA = line


                index = URL.LastIndexOf("/")
                filename = URL.Substring(index + 1, URL.Length - 1 - index).Trim()


                str_sql = "DELETE FROM apofaseis WHERE ada='" & ADA & "'"

                Dim dbio As DataBaseIO = New DataBaseIO
                records = dbio.DeleteValues(str_sql)


                str = "insert into apofaseis (ID,DIR_ID,ADA,THEMA, FOREAS_ID, EKDOUSA_ARXH_ID,  URL,PROTOKOLLO,
                        DDATE_ANARTISI,DDATE_PRAXIS,DDATE_INSERTION) values ('" &
                        CStr(maxID) & "','" & CStr(maxDir_ID) & "','" & CStr(ADA) & "','" & CStr(THEMA) & "','" & CStr(FOREAS_ID) & "','" &
                        CStr(EKDOUSA_ARXH_ID) & "','" & CStr(URL) & "','" &
                        CStr(PROTOKOLLO) & "','" & DATE_ANARTISI & "','" & DATE_PRAXIS & "','" & m_DateTo & " ')"


                dbio.InsertValues(str)


                maxID += 1
                maxDir_ID += 1


                If Not File.Exists(m_DataDir + filename + ".pdf") Then
                    My.Computer.Network.DownloadFile(URL, m_DataDir + filename + ".pdf")
                End If


            Next

        Catch E As Exception
            ' Let the user know what went wrong.
            MsgBox(E.Message)
        End Try
    End Sub
    Private Function CVS_lineparser(ByRef line As String) As String
        Dim index As Integer
        Dim str_term
        Try
            index = line.LastIndexOf(";")
            If index > 0 Then
                str_term = line.Substring(index + 1, line.Length - 1 - index).Trim()
                line = line.Substring(0, index).Trim()
            Else
                str_term = line
            End If

        Catch E As Exception
            MessageBox.Show(E.Message)
            str_term = ""
        End Try

        Return str_term
    End Function



End Class
