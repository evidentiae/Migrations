Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Text
Imports System
Imports System.IO
Imports System.Threading

Public Class Form1
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
        Dim thread As New Thread(AddressOf exec)
        thread.Start()

    End Sub
    Private Sub exec()
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



            xlWorkSheet.Cells(1, 2) = "ID"
            xlWorkSheet.Cells(1, 3) = "Plan name"
            xlWorkSheet.Cells(1, 4) = "Group no"
            xlWorkSheet.Cells(1, 5) = "insurance company id"
            ProgressBar1.Invoke(Sub()
                                    ProgressBar1.Minimum = 0
                                    ProgressBar1.Maximum = book.Sheets.Count

                                End Sub)

            Dim rd As Integer = 1 ' rownumberindestination
            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)
                Dim startingcell = 0

                For k = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                    If Not sheet.Cells(k, 2).value = Nothing Then

                        If IsNumeric(sheet.Cells(k, 2).value.ToString) Then
                            startingcell = k
                            Exit For
                        End If
                    End If
                Next

                For y = startingcell To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)


                    If Not sheet.Cells(y, 2).value = Nothing Then
                        If IsNumeric(sheet.Cells(y, 2).value.ToString) Then
                            If IsNumeric(getvalue(sheet, y + 2, 4)) Then
                                rd = rd + 1
                                xlWorkSheet.Cells(rd, 2) = sheet.Cells(y, 2).value.ToString 'id
                                xlWorkSheet.Cells(rd, 3) = sheet.Cells(y, 3).value.ToString ' plan name
                                xlWorkSheet.Cells(rd, 4) = sheet.Cells(y + 1, 3).value.ToString.Replace("Group No. :", "") 'group no
                                xlWorkSheet.Cells(rd, 5) = sheet.Cells(y + 2, 4).value.ToString 'insu cmp id
                                y = y + 3
                            Else
                                rd = rd + 1
                                xlWorkSheet.Cells(rd, 2) = sheet.Cells(y, 2).value.ToString
                                xlWorkSheet.Cells(rd, 3) = sheet.Cells(y, 3).value.ToString
                                xlWorkSheet.Cells(rd, 4) = sheet.Cells(y + 1, 2).value.ToString.Replace("Group No. :", "")
                                xlWorkSheet.Cells(rd, 5) = sheet.Cells(y + 2, 3).value.ToString
                                y = y + 3
                            End If


                        End If
                        'MsgBox(sheet.Cells(y, 2).value.ToString)
                        'MsgBox(sheet.Cells(y, 2).value.ToString)

                    End If
                Next


                Label2.Invoke(Sub()
                                  Label2.Text = i
                              End Sub)
                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Value = i
                                    End Sub)
            Next
            book.Close()
            xls.Workbooks.Close()
            xls.Quit()

            releaseObject(sheet)
            releaseObject(book)
            releaseObject(xls)


            xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\insurance-plansalpha254.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()



            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)






        Next



        MsgBox("done")
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



End Class
