Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Text
Imports System
Imports System.IO
Imports System.Threading

Public Class Form1
    Dim thread As New Thread(AddressOf exec)
    Dim thread1 As New Thread(AddressOf exec1)
    Function search(sheet As Excel.Worksheet, keyword As String, row As Integer, column As Integer, right As Integer, down As Integer, ByRef a As Integer, ByRef b As Integer)
        For i = row To row + down
            For j = column To column + right
                If Not (sheet.Cells(i, j).value = Nothing) Then
                    If sheet.Cells(i, j).value.ToString = keyword Then
                        a = i
                        b = j
                        Exit Function
                    End If
                End If
            Next
        Next
        a = -1
        b = -1
    End Function
    Function getvalue(sheet As Excel.Worksheet, row As Integer, col As Integer) As String

        If sheet.Cells(row, col).value = Nothing Then
            Return ""
        End If

        Return sheet.Cells(row, col).value.ToString
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        thread.Start()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Label1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Label1.Text = OpenFileDialog1.FileName
        End If
    End Sub
    Private Sub exec()
        'filling date,datemade,time,confirmdate,id,reason,patdr,med alerts,appt notes
        'this file has extension xls and was converted by pdf to excel

        Dim Dttbl As New System.Data.DataTable

        'destination
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        'source
        Dim xls As New Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet



        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        Dim rd As Integer = 1 ' rownumberindestination
        Dim filno = 0
        For Each file In OpenFileDialog1.FileNames

            Label1.Invoke(Sub()
                              Label1.Text = file
                          End Sub)

            filno = filno + 1

            Label6.Invoke(Sub()
                              Label6.Text = filno
                          End Sub)

            xls.Workbooks.Open(file)
            'get references to first workbook and worksheet



            book = xls.ActiveWorkbook
            sheet = book.ActiveSheet ' this is just to initialize / no need




            ProgressBar1.Invoke(Sub()
                                    ProgressBar1.Minimum = 0
                                    ProgressBar1.Maximum = book.Sheets.Count

                                End Sub)





            Dim i As Integer = 1
            While i <= book.Sheets.Count



                sheet = book.Sheets(i)


                    Label2.Invoke(Sub()
                                      Label2.Text = i

                                  End Sub)

                    ProgressBar1.Invoke(Sub()
                                            ProgressBar1.Value = i
                                        End Sub)
                    For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)



                        rd = rd + 1

                    For s = 1 To 12
                        If Not sheet.Cells(y, s).value = Nothing Then


                            xlWorkSheet.Cells(rd, s) = sheet.Cells(y, s).value.ToString

                        End If
                    Next

                Next




                i = i + 1

            End While
            book.Close()
            xls.Workbooks.Close()

        Next

        xls.Quit()

        releaseObject(sheet)
        releaseObject(book)
        releaseObject(xls)


        xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\merged-" & filno & ".xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()



        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        MsgBox("done")

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        thread1.Start()
    End Sub
    Private Sub exec1()
        'filling date,datemade,time,confirmdate,id,reason,patdr,med alerts,appt notes
        'this file has extension xls and was converted by pdf to excel

        Dim Dttbl As New System.Data.DataTable

        'Source file
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        'Destination file
        Dim xls As New Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet


        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        Dim rd As Integer = 1 ' rownumberindestination

        xlWorkSheet.Cells(1, 2) = "PlanID"
        xlWorkSheet.Cells(1, 3) = "Coverage by class"
        xlWorkSheet.Cells(1, 4) = "percentage"
        xlWorkSheet.Cells(1, 5) = "effective month"
        xlWorkSheet.Cells(1, 6) = "Max coverage"
        xlWorkSheet.Cells(1, 7) = "individual deductible"


        rd = rd + 1


        For Each file In OpenFileDialog1.FileNames

            Label1.Invoke(Sub()
                              Label1.Text = file
                          End Sub)


            xls.Workbooks.Open(file)
            'get references to first workbook and worksheet

            book = xls.ActiveWorkbook
            sheet = book.ActiveSheet ' this is just to initialize / no need









            ProgressBar2.Invoke(Sub()
                                    ProgressBar2.Minimum = 0
                                    ProgressBar2.Maximum = (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                                End Sub)



            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)





                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                    ProgressBar2.Invoke(Sub()
                                            ProgressBar2.Value = y
                                        End Sub)





                    Dim planid, effectivemonth, maxcoverage, individualdeductible As String
                    If Not sheet.Cells(y, 1).value = Nothing Then
                        If sheet.Cells(y, 1).value.ToString = "Plan ID:" Then
                            planid = sheet.Cells(y, 2).value.ToString

                            Dim rowofmaxcoverage = 0
                            Dim colofmaxcoverage = 0
                            Dim rowofeffectivemonth = 0
                            Dim colofeffectivemonth = 0
                            Dim rowofindividualdeductible = 0
                            Dim colofindividualdeductible = 0
                            y = y + 6

                            Try
                                search(sheet, "Effective Month:", y + 7, 6, 10, 3, rowofeffectivemonth, colofeffectivemonth)
                                search(sheet, "Max Coverage:", y + 1, 6, 10, 3, rowofmaxcoverage, colofmaxcoverage)
                                search(sheet, "Individual Deductible:", y + 4, 6, 10, 3, rowofindividualdeductible, colofindividualdeductible)

                                effectivemonth = sheet.Cells(rowofeffectivemonth, colofeffectivemonth + 1).value.ToString
                                maxcoverage = sheet.Cells(rowofmaxcoverage, colofmaxcoverage + 1).value.ToString
                                individualdeductible = sheet.Cells(rowofindividualdeductible, colofindividualdeductible + 1).value.ToString
                            Catch ex As Exception
                                'MsgBox(y)
                            End Try



                            Dim rowofcoveragebyclass = 0
                            Dim colofcoveragebyclass = 0
                            search(sheet, "Coverage By Class", y, 2, 10, 2, rowofcoveragebyclass, colofcoveragebyclass)
                            If Not rowofcoveragebyclass = -1 Then
                                y = rowofcoveragebyclass + 1
                            End If

                            While y < sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row

                                If Not sheet.Cells(y, 2).value = Nothing Then
                                    If sheet.Cells(y, 2).value.ToString = "Code" Then
                                        y = y - 1
                                        Exit While
                                    Else
                                        Dim pl = ""
                                        Dim percent = ""


                                        If sheet.Cells(y, 2).value.ToString.IndexOf("%") > 0 Then
                                            pl = sheet.Cells(y, 2).value.ToString.Remove(sheet.Cells(y, 2).value.ToString.IndexOf("["))
                                            Dim spli = sheet.Cells(y, 2).value.ToString.Split
                                            percent = spli(spli.Length - 2)
                                        ElseIf getvalue(sheet, y, 3).IndexOf("%") > 0 Then
                                            pl = sheet.Cells(y, 2).value.ToString
                                            If sheet.Cells(y, 3).Value.ToString.IndexOf("0%") > 0 Then
                                                percent = "0 %"
                                            Else
                                                Dim spli = sheet.Cells(y, 3).value.ToString.Split(" ")
                                                percent = spli(spli.Length - 2)
                                            End If
                                        ElseIf getvalue(sheet, y, 4).IndexOf("%") > 0 Then
                                            pl = sheet.Cells(y, 2).value.ToString
                                            If sheet.Cells(y, 4).Value.ToString.IndexOf("0%") > 0 Then
                                                percent = "0 %"
                                            Else
                                                Dim spli = sheet.Cells(y, 4).value.ToString.Split(" ")
                                                percent = spli(spli.Length - 2)
                                            End If
                                        End If

                                        Label3.Invoke(Sub()
                                                          Label3.Text = y
                                                      End Sub)
                                        If Not pl = "" And Not percent = "" Then
                                            xlWorkSheet.Cells(rd, 2) = planid
                                            xlWorkSheet.Cells(rd, 3) = pl
                                            xlWorkSheet.Cells(rd, 4) = percent
                                            xlWorkSheet.Cells(rd, 5) = effectivemonth
                                            xlWorkSheet.Cells(rd, 6) = maxcoverage
                                            xlWorkSheet.Cells(rd, 7) = individualdeductible
                                            rd = rd + 1
                                        End If


                                        'MsgBox(rd)
                                    End If
                                End If
                                y = y + 1
                            End While
                        End If
                    End If



                Next





            Next
            book.Close()
            xls.Workbooks.Close()





        Next
        xls.Quit()

        releaseObject(sheet)
        releaseObject(book)
        releaseObject(xls)
        xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\bluebookfinal254.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
             Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()



        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)


        MsgBox("done")
    End Sub
End Class
