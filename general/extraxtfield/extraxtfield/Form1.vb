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
        Dim keyword = TextBox1.Text
        Dim endword1 = TextBox2.Text
        Dim endword2 = TextBox3.Text
        Dim COLUMN As Integer = Val(TextBox4.Text)

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



            Dim continu As Boolean = True

            Dim rd As Integer = 1 ' rownumberindestination

            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)

                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Maximum = (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                                        ProgressBar1.Minimum = 1
                                    End Sub)



                rd = rd + 1 ' two lines can be removed if you want to complete on new page 
                Dim col = 1


                Dim carr = ""

                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                    ' neglecting empty boxes
                    If getvalue(sheet, y, COLUMN).Contains(keyword) Then
                        xlWorkSheet.Cells(rd, 1) = (getvalue(sheet, y, COLUMN) & getvalue(sheet, y, COLUMN + 1)).Replace(keyword, "").Replace(endword1, "").Replace(endword2, "")
                        rd = rd + 1
                    ElseIf getvalue(sheet, y, COLUMN + 1).Contains(keyword) Then
                        xlWorkSheet.Cells(rd, 1) = (getvalue(sheet, y, COLUMN + 1) & getvalue(sheet, y, COLUMN + 2)).Replace(keyword, "").Replace(endword1, "").Replace(endword2, "")
                        rd = rd + 1
                    End If

                    Label3.Invoke(Sub()
                                      Label3.Text = y
                                  End Sub)
                    ProgressBar1.Invoke(Sub()
                                            ProgressBar1.Value = y
                                        End Sub)

                Next



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


            xlWorkBook.SaveAs("C:\Users\developer\Desktop\bew\tmp\bew6.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)






        Next



        MsgBox("done")

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim thread As New Thread(AddressOf exec2)
        thread.Start()
    End Sub
    Private Sub exec2()
        'filling date,datemade,time,confirmdate,id,reason,patdr,med alerts,appt notes
        'this file has extension xls and was converted by pdf to excel
        Dim keyword = TextBox1.Text
        Dim endword1 = TextBox2.Text
        Dim endword2 = TextBox3.Text
        Dim Dttbl As New System.Data.DataTable
        Dim COLUMN As Integer = Val(TextBox4.Text)

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


            Dim continu As Boolean = True

            Dim rd As Integer = 1 ' rownumberindestination

            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)

                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Maximum = (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                                        ProgressBar1.Minimum = 1
                                    End Sub)


                rd = rd + 1 ' two lines can be removed if you want to complete on new page 
                Dim col = 1


                Dim carr = ""

                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                    ' neglecting empty boxes
                    If getvalue(sheet, y, COLUMN).Contains(keyword) Then
                        xlWorkSheet.Cells(rd, COLUMN) = (getvalue(sheet, y + 1, COLUMN) & getvalue(sheet, y + 1, COLUMN + 1)).Replace(keyword, "").Replace(endword1, "").Replace(endword2, "")
                        rd = rd + 1
                    ElseIf getvalue(sheet, y, COLUMN + 1).Contains(keyword) Then
                        xlWorkSheet.Cells(rd, 1) = (getvalue(sheet, y + 1, COLUMN + 1) & getvalue(sheet, y + 1, COLUMN + 2)).Replace(keyword, "").Replace(endword1, "").Replace(endword2, "")
                        rd = rd + 1
                    End If

                    Label3.Invoke(Sub()
                                      Label3.Text = y
                                  End Sub)
                    ProgressBar1.Invoke(Sub()
                                            ProgressBar1.Value = y
                                        End Sub)

                Next



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


            xlWorkBook.SaveAs("C:\Users\developer\Desktop\bew\tmp\bew6.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
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
