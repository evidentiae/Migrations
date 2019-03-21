Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Text
Imports System
Imports System.IO
Imports System.Threading

Public Class Form1
    Dim thread As New Thread(AddressOf exec)
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

        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")



        xlWorkSheet.Cells(1, 2) = "ID"
        xlWorkSheet.Cells(1, 3) = "name"
        xlWorkSheet.Cells(1, 4) = "phone"
        xlWorkSheet.Cells(1, 5) = "address"

        Dim rd As Integer = 1 ' rownumberindestination


        For Each file In OpenFileDialog1.FileNames
            Label1.Invoke(Sub()
                              Label1.Text = file
                          End Sub)

            xls.Workbooks.Open(file)
            'get references to first workbook and worksheet

            book = xls.ActiveWorkbook
            sheet = book.ActiveSheet ' this is just to initialize / no need

            ProgressBar1.Invoke(Sub()
                                    ProgressBar1.Minimum = 0
                                    ProgressBar1.Maximum = book.Sheets.Count

                                End Sub)


            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)


                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Value = i
                                    End Sub)

                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)


                    If IsNumeric(getvalue(sheet, y, 2)) Then
                        If getvalue(sheet, y + 1, 2) = "" Then
                            rd = rd + 1
                            xlWorkSheet.Cells(rd, 2) = getvalue(sheet, y, 2) 'id
                            Dim n = getvalue(sheet, y, 3)
                            If n.Contains("Ph 1:") Then
                                n = n.Substring(0, n.IndexOf("Ph 1:"))
                            End If
                            xlWorkSheet.Cells(rd, 3) = n 'name
                            xlWorkSheet.Cells(rd, 5) = getvalue(sheet, y + 1, 3) & " " & getvalue(sheet, y + 2, 3) 'address
                            Dim jk = ""
                            For ui = 4 To 7
                                If getvalue(sheet, y, ui).Contains("(") And getvalue(sheet, y, ui).Contains(")") Then
                                    jk = getvalue(sheet, y, ui)
                                    jk = jk.Replace("Ph 1:", "")
                                    jk = jk.Replace(" ", "")
                                    If jk.Length >= jk.IndexOf(")") + 9 Then
                                        jk = jk.Remove(jk.IndexOf(")") + 9)
                                    End If

                                    Exit For
                                End If
                            Next
                            xlWorkSheet.Cells(rd, 4) = jk 'phone
                            y = y + 2
                        Else
                            rd = rd + 1
                            xlWorkSheet.Cells(rd, 2) = getvalue(sheet, y, 2) 'id
                            Dim n = getvalue(sheet, y, 3)

                            If n.Contains("Ph 1:") Then
                                n = n.Substring(0, n.IndexOf("Ph 1:"))
                            End If
                            If n.Contains("Ph C 1:") Then
                                n = n.Substring(0, n.IndexOf("Ph C 1:"))
                            End If
                            xlWorkSheet.Cells(rd, 3) = n 'name 'name

                            xlWorkSheet.Cells(rd, 5) = getvalue(sheet, y + 1, 2) & " ".Replace("Ph 2:", "") & " " & getvalue(sheet, y + 2, 2) & " ".Replace("FAX:", "") 'address
                            Dim jk = ""
                            For ui = 4 To 7
                                If getvalue(sheet, y, ui).Contains("(") And getvalue(sheet, y, ui).Contains(")") Then
                                    jk = getvalue(sheet, y, ui)
                                    jk = jk.Replace("Ph 1:", "")
                                    jk = jk.Replace(" ", "")
                                    If jk.Length >= jk.IndexOf(")") + 9 Then
                                        jk = jk.Remove(jk.IndexOf(")") + 9)
                                    End If

                                    Exit For
                                End If
                            Next
                            xlWorkSheet.Cells(rd, 4) = jk 'phone
                            y = y + 2
                        End If
                        'MsgBox(sheet.Cells(y, 2).value.ToString)
                        'MsgBox(sheet.Cells(y, 2).value.ToString)

                    End If
                    Label3.Invoke(Sub()
                                      Label3.Text = y
                                  End Sub)
                Next



                Label2.Invoke(Sub()
                                  Label2.Text = i
                              End Sub)
            Next
            book.Close()
            xls.Workbooks.Close()





        Next
        xls.Quit()

        releaseObject(sheet)
        releaseObject(book)
        releaseObject(xls)


        xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\alphacompany254.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()



        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)




        MsgBox("done")
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



End Class
