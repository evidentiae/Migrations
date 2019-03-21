Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Text
Imports System
Imports System.IO
Imports System.Threading

Public Class Form1
    Dim thread As New Thread(AddressOf exec)
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
                                    ProgressBar1.Minimum = 1
                                    ProgressBar1.Maximum = book.Sheets.Count

                                End Sub)





            Dim i As Integer = 1
            While i <= book.Sheets.Count

                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Value = i
                                    End Sub)

                sheet = book.Sheets(i)


                    Label2.Invoke(Sub()
                                      Label2.Text = i

                                  End Sub)



                xlWorkSheet.Cells(1, 2) = "planid"
                xlWorkSheet.Cells(, 3) = "patientid"
                xlWorkSheet.Cells(1, 4) = "patientname"
                xlWorkSheet.Cells(1, 5) = "g name"
                xlWorkSheet.Cells(1, 6) = "coverage"
                xlWorkSheet.Cells(1, 7) = "ins balance"


                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                    Dim planid = ""

                    If (getvalue(sheet, y, 1) & getvalue(sheet, y, 2)).Contains("[") Then

                        Dim temp = getvalue(sheet, y, 1) & getvalue(sheet, y, 2) & getvalue(sheet, y, 3)

                        planid = temp.Substring(temp.IndexOf("[") + 1)
                        planid = planid.Remove(planid.IndexOf("]"))


                        y = y + 1

                        While Not ((getvalue(sheet, y, 1) & getvalue(sheet, y, 1)).Contains("[")) And y < (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                            Dim starting = 0
                            If IsNumeric(getvalue(sheet, y, 2)) Then
                                starting = 2
                            ElseIf IsNumeric(getvalue(sheet, y, 3)) Then
                                starting = 3
                            ElseIf IsNumeric(getvalue(sheet, y, 1)) Then
                                starting = 1
                            End If

                            If Not starting = 0 Then

                                rd = rd + 1

                                Dim patientid = "", patientname = "", gname = "", coverage = "", insbal = "" ' no need for this
                                patientid = getvalue(sheet, y, starting)
                                patientname = getvalue(sheet, y, starting + 1).Replace("INACTIVE", "").Replace("*", "")
                                Dim status = ""
                                Dim colu = 1
                                For m = 0 To 4
                                    If colu = 4 Then
                                        Exit For
                                    End If
                                    Dim tmp = getvalue(sheet, y, starting + 2 + m)
                                    If Not tmp.Replace("INACTIVE", "").Replace("*", "").Trim = "" Then
                                        xlWorkSheet.Cells(rd, 4 + colu) = tmp.Replace("INACTIVE", "").Replace("*", "").Trim
                                        colu = colu + 1
                                    End If
                                Next



                                If Not getvalue(sheet, y, starting + 2).Contains("INACTIVE") Then
                                        'gname = getvalue(sheet, y, starting + 2)
                                        'coverage = getvalue(sheet, y, starting + 3)
                                        ' insbal = getvalue(sheet, y, starting + 4)
                                        status = "active"
                                    Else
                                        'gname = getvalue(sheet, y, starting + 3).Replace("INACTIVE", "").Replace("*", "")
                                        ' coverage = getvalue(sheet, y, starting + 4)
                                        ' insbal = getvalue(sheet, y, starting + 5)
                                        status = "inactive"
                                    End If


                                xlWorkSheet.Cells(rd, 2) = planid
                                xlWorkSheet.Cells(rd, 3) = patientid
                                xlWorkSheet.Cells(rd, 4) = patientname
                                'xlWorkSheet.Cells(rd, 5) = gname
                                'xlWorkSheet.Cells(rd, 6) = coverage
                                'xlWorkSheet.Cells(rd, 7) = insbal
                                xlWorkSheet.Cells(rd, 8) = status



                                Dim sd = Split(getvalue(sheet, y, starting + 1).Replace("*", ""), " INACTIVE ")

                                If sd.Length > 1 Then
                                    If sd(1).Length > 1 Then


                                        xlWorkSheet.Cells(rd, 4) = sd(0)
                                        xlWorkSheet.Cells(rd, 5) = sd(1)
                                        xlWorkSheet.Cells(rd, 6) = getvalue(sheet, y, starting + 2)
                                        xlWorkSheet.Cells(rd, 7) = getvalue(sheet, y, starting + 3)
                                    End If

                                End If
                            End If
                                y = y + 1
                            Label3.Invoke(Sub()
                                              Label3.Text = y

                                          End Sub)

                        End While
                        y = y - 1
                    End If





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


        xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\patientsbyplan254.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()



        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        MsgBox("done")

    End Sub


End Class
