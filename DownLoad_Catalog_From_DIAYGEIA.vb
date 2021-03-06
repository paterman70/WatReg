Imports System.IO
Imports System.Xml
Imports Excel_App = Microsoft.Office.Interop.Excel

Public Class DownLoadFromDIAYGEIA
    Private m_DataDir As String
    Private m_InputFormat As String
    Private m_DateFrom As String
    Private m_DateTo As String

    Private Function DOWNLOAD_APOFASEIS_IN_XLS() As Boolean

        Dim bExist As Boolean = True
        Dim str_sql As String
        Dim counter As Integer = 0
        Dim b_test As Boolean = False
        Dim bresult As Boolean = True

        Dim missing As Object = System.Type.Missing

        Dim path As String = ""
        Dim FileName = "total.csv"
        Dim ListData As New List(Of String())
        Dim MyList As New List(Of String)

        Try
            If b_test Then

            Else


                path = m_DataDir
                If (Not Directory.Exists(path)) Then
                    Directory.CreateDirectory(path)
                End If

                'delete already exist files
                My.Computer.FileSystem.WriteAllText(path + FileName, " " + vbCrLf, False)




                str_sql = "Select FOREAS_ID,EKDOUSA_ARXH_ID,EKDOUSA_ARXH_PERIGRAFI from FOREAS"
                Dim dbio As DataBaseIO = New DataBaseIO
                ListData = dbio.GetValuesinList(str_sql, str_sql.Split(",").Length)



                Select Case m_InputFormat
                    Case "XSL"
                        MyList = GetXLS(ListData, path)

                    Case "XML"
                        MyList = GetXML(ListData, path)
                    Case Else

                End Select
                If MyList.Count = 0 Then bresult = False
                For Each line In MyList
                    line = line.Replace(vbCrLf, " ")
                    line = line.Replace(vbCr, " ")
                    line = line.Replace(vbLf, " ")
                    line = line.Replace(Environment.NewLine, " ")


                    My.Computer.FileSystem.WriteAllText(path + "total.csv", line + vbCrLf, True)
                Next

            End If

        Catch Ex As Exception
            bresult = False


        End Try
        Return bresult
    End Function
    Private Function GetXLS(ByRef Foreas As List(Of String()), ByVal path As String) As List(Of String)
        Dim Mylist As List(Of String) = New List(Of String)
        Dim URL As String = ""
        Dim URLDate As String = ""
        Dim file As String = ""
        Dim worksheetCells As Excel_App.Range

        URLDate = "fq=issueDate:[DT(" & m_DateFrom & "T00:00:00)%20TO%20DT(" & m_DateTo & "T23:59:59)]&wt=xls"



        For Each marray In Foreas

            URL = "https://app.diavgeia.gov.gr/luminapi/api/search/export?q=organizationUid:%22" & marray(0) & "%22&fq=unitUid:%22" & marray(1) & "%22&" & URLDate

            file = marray(2) + ".xls"
            My.Computer.Network.DownloadFile(URL, path + file, "", "", True, 500, True)


            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim book As Microsoft.Office.Interop.Excel.Workbook
            Dim sheet As New Microsoft.Office.Interop.Excel.Worksheet()
            Dim cellValues As Object(,)

            book = excel.Workbooks.Open(path + file)
            sheet = book.Worksheets(1)
            sheet.Activate()
            worksheetCells = sheet.UsedRange()
            If worksheetCells IsNot Nothing Then
                worksheetCells = worksheetCells.Resize(worksheetCells.Rows.Count, worksheetCells.Columns.Count + 2)
                cellValues = worksheetCells.Value

                Dim txt As String = ""
                For firstDim As Integer = 2 To cellValues.GetUpperBound(0)
                    Dim line As String = ""
                    For secondDim As Integer = 1 To cellValues.GetUpperBound(1)
                        If secondDim < cellValues.GetUpperBound(1) - 1 Then
                            txt = cellValues(firstDim, secondDim).ToString + ";"
                        ElseIf secondDim = cellValues.GetUpperBound(1) - 1 Then
                            txt = marray(1).ToString + ";"
                        Else
                            txt = "@"
                        End If
                        line += txt
                    Next
                    Mylist.Add(line)
                Next

            End If

            excel.Workbooks.Close()
            excel.Quit()
        Next

        Return Mylist
    End Function
    Private Function GetXML(ByRef Foreas As List(Of String()), ByVal path As String) As List(Of String)
        Dim Mylist As List(Of String) = New List(Of String)
        Dim URL As String = ""
        Dim URLDate As String = ""
        Dim file As String = ""

        URLDate = "fq=issueDate:[DT(" & m_DateFrom & "T00:00:00)%20TO%20DT(" & m_DateTo & "T23:59:59)]&wt=xml"

        Try
            For Each marray In Foreas
                Try

                    URL = "https://app.diavgeia.gov.gr/luminapi/api/search/export?q=organizationUid:%22" & marray(0) & "%22&fq=unitUid:%22" & marray(1) & "%22&" & URLDate

                    Dim reader As XmlReader = XmlReader.Create(URL)
                    Dim decision As List(Of String) = New List(Of String)
                    Dim line As String
                    While reader.Read()
                        Select Case reader.NodeType
                            Case XmlNodeType.Text
                                decision.Add(reader.Value)
                            Case XmlNodeType.EndElement
                                If reader.Name = "decision" Then

                                    line = decision.Item(0) + ";"
                                    line += decision.Item(4) + ";"
                                    line += decision.Item(8) + ";"
                                    line += decision.Item(6) + ";"
                                    line += decision.Item(5) + ";"
                                    line += decision.Item(2) + ";"
                                    line += decision.Item(1) + ";"
                                    line += decision.Item(3) + ";"
                                    line += decision.Item(7) + ";"
                                    line += decision.Item(9) + ";"

                                    line += marray(1).ToString + ";@"

                                    Mylist.Add(line)
                                    decision.Clear()
                                End If

                        End Select
                    End While

                    reader.Close()
                    '"?><decisionresults>
                    '<decision>
                    '<ada>6??????????1??-4??4</ada>
                    '<decisionTypeLabel>???????????? ???????????????? ?????????????????????? ??????????????</decisionTypeLabel>
                    '<decisionTypeUid>2.4.7.1</decisionTypeUid>
                    '<documentUrl>https://diavgeia.gov.gr/doc/6??????????1??-4??4</documentUrl>
                    '<issueDate>14/07/2020 03:00:00</issueDate>
                    '<organizationLabel>?????????????????????????? ???????????????? ???????????????????? ??? ????????????</organizationLabel>
                    '<organizationUid>50205</organizationUid>
                    '<protocolNumber>25355</protocolNumber>
                    '<subject>?????????????????? ???????????? ???????????? ?????????? ???????? ...??.</subject>
                    '<submissionTimestamp>14/07/2020 11:08:01</submissionTimestamp>
                    '</decision>
                Catch ex As Exception

                End Try

            Next

        Catch ex As Exception

        End Try

        Return Mylist
    End Function

End Class
