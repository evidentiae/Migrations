Imports System.Threading
Imports Microsoft.Office.Interop

Public Class Form1
    Function returnsecondslash(s As String) As Integer
        For i = 0 To s.Length - 4
            If s.Chars(i) = "/" And s.Chars(i + 2) = "/" Then
                Return i + 2
            ElseIf s.Chars(i) = "/" And s.Chars(i + 3) = "/" Then
                Return i + 3

            End If
        Next
        Return s.Length
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Label1.Text = OpenFileDialog1.FileName
        End If
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



            xlWorkSheet.Cells(1, 2) = "name"
            xlWorkSheet.Cells(1, 3) = "address1"
            xlWorkSheet.Cells(1, 4) = "address2"
            xlWorkSheet.Cells(1, 5) = "email"
            xlWorkSheet.Cells(1, 6) = "(P)"
            xlWorkSheet.Cells(1, 7) = "(S)"
            xlWorkSheet.Cells(1, 8) = "Phone"
            xlWorkSheet.Cells(1, 9) = "fax"
            xlWorkSheet.Cells(1, 10) = "mobile"
            xlWorkSheet.Cells(1, 11) = "prov"
            xlWorkSheet.Cells(1, 12) = "position"
            xlWorkSheet.Cells(1, 13) = "birth"
            xlWorkSheet.Cells(1, 14) = "ss"
            xlWorkSheet.Cells(1, 15) = "chart"
            xlWorkSheet.Cells(1, 16) = "gender"
            xlWorkSheet.Cells(1, 17) = "status"
            xlWorkSheet.Cells(1, 18) = "phone W"
            xlWorkSheet.Cells(1, 19) = "phone O"

            Dim rd As Integer = 1 ' rownumberindestination
            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)


                Label2.Invoke(Sub()
                                  Label2.Text = i
                              End Sub)
                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Minimum = 0
                                        ProgressBar1.Maximum = (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                                    End Sub)

                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)


                    Label3.Invoke(Sub()
                                      Label3.Text = y
                                  End Sub)

                    ProgressBar1.Invoke(Sub()
                                            ProgressBar1.Value = y

                                        End Sub)


                    Dim appdate As Integer = -1
                    For h = 5 To 7
                        If Not sheet.Cells(y, h).value Is Nothing Then
                            If getvalue(sheet, y, h).Contains("Birth:") Then
                                appdate = h + 1
                                Exit For
                                'MsgBox(appdate)
                            End If

                        End If
                    Next

                    If Not appdate = -1 Then

                        'MsgBox(sheet.Cells(y, 2).value.ToString)
                        rd = rd + 1

                        xlWorkSheet.Cells(rd, 2) = getvalue(sheet, y, 1) 'name
                        xlWorkSheet.Cells(rd, 3) = getvalue(sheet, y + 1, 1) 'address1
                        xlWorkSheet.Cells(rd, 4) = getvalue(sheet, y + 2, 1) 'address2
                        xlWorkSheet.Cells(rd, 5) = getvalue(sheet, y + 3, 1).Replace("E-Mail:", "") 'email
                        xlWorkSheet.Cells(rd, 6) = getvalue(sheet, y + 4, 1).Replace("(P)", "") '(p)
                        xlWorkSheet.Cells(rd, 7) = getvalue(sheet, y + 5, 1).Replace("(S)", "") '(s)
                        xlWorkSheet.Cells(rd, 8) = getvalue(sheet, y, 3).Replace("(H)", "").Replace("(", "").Replace(")", "") 'phone
                        xlWorkSheet.Cells(rd, 9) = getvalue(sheet, y + 3, 3).Replace("(", "").Replace(")", "") 'fax
                        xlWorkSheet.Cells(rd, 10) = getvalue(sheet, y + 4, 3).Replace("(", "").Replace(")", "") 'mobile
                        xlWorkSheet.Cells(rd, 11) = getvalue(sheet, y + 5, 3).Replace("Billing Type:", "") 'prov
                        xlWorkSheet.Cells(rd, 12) = getvalue(sheet, y + 6, 3) 'position
                        xlWorkSheet.Cells(rd, 13) = getvalue(sheet, y, appdate) 'birth
                        xlWorkSheet.Cells(rd, 14) = getvalue(sheet, y + 1, appdate) 'ss
                        xlWorkSheet.Cells(rd, 15) = getvalue(sheet, y + 2, appdate) 'chart
                        xlWorkSheet.Cells(rd, 16) = getvalue(sheet, y + 5, appdate) 'gender
                        xlWorkSheet.Cells(rd, 17) = getvalue(sheet, y + 6, appdate) 'status
                        xlWorkSheet.Cells(rd, 18) = getvalue(sheet, y + 1, 3).Replace("(W)", "").Replace("(", "").Replace(")", "") 'phone W
                        xlWorkSheet.Cells(rd, 19) = getvalue(sheet, y + 2, 3).Replace("(O)", "").Replace("(", "").Replace(")", "") 'phone O






                        y = y + 6
                    ElseIf False Then

                    End If

                Next



            Next
            book.Close()
            xls.Workbooks.Close()
            xls.Quit()

            releaseObject(sheet)
            releaseObject(book)
            releaseObject(xls)


            xlWorkBook.SaveAs("C:\Users\developer\Desktop\bew\easydent\bew\patient list standard\patient337.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
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
