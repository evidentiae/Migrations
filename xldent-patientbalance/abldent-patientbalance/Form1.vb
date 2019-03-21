Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Text
Imports System
Imports System.IO
Imports System.Threading

Public Class Form1
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Label1.Text = OpenFileDialog1.FileName
        End If
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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Function countAppearence(s As String, c As Char) As Integer
        Dim coun = 0
        For i = 0 To s.Count - 1
            If c = s.Chars(i) Then
                coun = coun + 1
            End If
        Next
        Return coun
    End Function
    Function isname(s As String) As Boolean
        Dim percent As Integer = 0
        For i = 0 To s.Count - 1
            If Char.IsLetter(s.Chars(i)) Then
                percent = percent + 1
            End If
        Next
        If percent > s.Count / 2 And percent > 1 Then
            Return True
        End If
        Return False
    End Function
    Function getvalue(sheet As Excel.Worksheet, row As Integer, col As Integer) As String

        If sheet.Cells(row, col).value = Nothing Then
            Return ""
        End If

        Return sheet.Cells(row, col).value.ToString
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim thread As New Thread(AddressOf exec)
        thread.Start()
    End Sub
    Private Sub exec()
        'filling date,datemade,time,confirmdate,id,reason,patdr,med alerts,appt notes
        'this file has extension xls and was converted by pdf to excel

        Dim COLUMN As Integer = 1

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




        For Each file In OpenFileDialog1.FileNames
            Label1.Invoke(Sub()
                              Label1.Text = file
                          End Sub)

            xls.Workbooks.Open(file)
            'get references to first workbook and worksheet

            book = xls.ActiveWorkbook
            sheet = book.ActiveSheet ' this is just to initialize / no need



            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")


            xlWorkSheet.Cells(1, 1) = "account"
            xlWorkSheet.Cells(1, 2) = "last"
            xlWorkSheet.Cells(1, 3) = "P"
            xlWorkSheet.Cells(1, 4) = "I"
            xlWorkSheet.Cells(1, 4) = "T"
            ProgressBar1.Invoke(Sub()
                                    ProgressBar1.Maximum = book.Sheets.Count
                                    ProgressBar1.Minimum = 1
                                End Sub)

            Dim rd As Integer = 1 ' rownumberindestination

            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)





                rd = rd + 1 ' two lines can be removed if you want to complete on new page 
                Dim col = 1
                Dim vJ = 10
                If getvalue(sheet, 4, 9) = "Total" Then
                    vJ = 9
                End If

                For y = 5 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)


                    Dim dat = ""


                    If isname(getvalue(sheet, y, 2)) Then

                        dat = getvalue(sheet, y, 4).Replace("P:", "")

                        If Not dat.Contains("/") Then
                            dat = getvalue(sheet, y, 5).Replace("P:", "")
                        End If
                        If Not dat.Contains("/") Then
                            dat = ""
                        End If
                        Dim j = y + 1
                        If i = 4 Then
                            MsgBox(dat)
                        End If
                        Dim found = False

                        While Not isname(getvalue(sheet, j, 2)) And j < (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                            found = True
                            If isname(getvalue(sheet, j, 3)) Then
                                Dim vP = getvalue(sheet, j, vJ)
                                Dim vI = getvalue(sheet, j + 1, vJ)
                                Dim vT = getvalue(sheet, j + 2, vJ)

                                If vT.Trim = "" And vI.Trim = "" Then

                                    If getvalue(sheet, 4, 9) = "Total" Then
                                        vI = getvalue(book.Sheets(i + 1), 5, 9)
                                        vT = getvalue(book.Sheets(i + 1), 6, 9)
                                    Else
                                        vI = getvalue(book.Sheets(i + 1), 5, 10)
                                        vT = getvalue(book.Sheets(i + 1), 6, 10)
                                    End If
                                ElseIf vT.Trim = "" Then

                                    If getvalue(sheet, 4, 9) = "Total" Then

                                        vT = getvalue(book.Sheets(i + 1), 5, 9)
                                    Else

                                        vT = getvalue(book.Sheets(i + 1), 5, 10)
                                    End If
                                End If

                                xlWorkSheet.Cells(rd, 1) = getvalue(sheet, j, 3)
                                xlWorkSheet.Cells(rd, 2) = dat
                                xlWorkSheet.Cells(rd, 3) = vP
                                xlWorkSheet.Cells(rd, 4) = vI
                                xlWorkSheet.Cells(rd, 5) = vT
                                rd = rd + 1
                            End If
                            j = j + 1
                        End While
                        If found = False And Not getvalue(sheet, y, 2) = "Account Name" Then
                            Dim vP = getvalue(sheet, y, vJ)
                            Dim vI = getvalue(sheet, y + 1, vJ)
                            Dim vT = getvalue(sheet, y + 2, vJ)
                            If isname(getvalue(sheet, y, 3)) Then
                                xlWorkSheet.Cells(rd, 1) = (getvalue(sheet, y, 2) & getvalue(sheet, y, 3)).Replace("(H)", "")
                            Else
                                xlWorkSheet.Cells(rd, 1) = getvalue(sheet, y, 2).Replace("(H)", "")
                            End If
                            If vT.Trim = "" And vI.Trim = "" Then

                                If getvalue(sheet, 4, 9) = "Total" Then
                                    vI = getvalue(book.Sheets(i + 1), 5, 9)
                                    vT = getvalue(book.Sheets(i + 1), 6, 9)
                                Else
                                    vI = getvalue(book.Sheets(i + 1), 5, 10)
                                    vT = getvalue(book.Sheets(i + 1), 6, 10)
                                End If
                            ElseIf vT.Trim = "" Then

                                If getvalue(sheet, 4, 9) = "Total" Then

                                    vT = getvalue(book.Sheets(i + 1), 5, 9)
                                Else

                                    vT = getvalue(book.Sheets(i + 1), 5, 10)
                                End If
                            End If
                            xlWorkSheet.Cells(rd, 2) = dat
                            xlWorkSheet.Cells(rd, 3) = vP
                            xlWorkSheet.Cells(rd, 4) = vI
                            xlWorkSheet.Cells(rd, 5) = vT
                            rd = rd + 1
                        End If
                    End If



                    Label3.Invoke(Sub()
                                      Label3.Text = y
                                  End Sub)

                Next

                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Value = i
                                    End Sub)


                Label2.Invoke(Sub()
                                  Label2.Text = i
                              End Sub)

            Next
            book.Close()
            xls.Workbooks.Close()
            xls.Quit()

            releaseObject(sheet)
            releaseObject(book)
            releaseObject(xls)


            xlWorkBook.SaveAs("C:\Users\developer\Desktop\bew\tmp\bew012.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)






        Next



        MsgBox("done")

    End Sub

End Class
