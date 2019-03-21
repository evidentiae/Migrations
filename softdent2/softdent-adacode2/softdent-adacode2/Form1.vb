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



                xlWorkSheet.Cells(1, 2) = "id"
                xlWorkSheet.Cells(, 3) = "description"
                For oiu = 0 To 26
                    xlWorkSheet.Cells(1, oiu + 4) = "Fee " & oiu
                Next



                For y = 5 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                    Dim id = ""
                    Dim description = ""

                    If getvalue(sheet, y, 1) <> "" Then

                        id = getvalue(sheet, y, 1)
                        description = getvalue(sheet, y, 3) & getvalue(sheet, y, 4) & getvalue(sheet, y, 5) & getvalue(sheet, y, 6) & getvalue(sheet, y, 7) & getvalue(sheet, y, 8)

                        y = y + 1
                        Dim feevalues(27) As Integer
                        Dim all = ""
                        For yi = y To y + 8
                            For xi = 2 To 10
                                all = all & getvalue(sheet, yi, xi)
                            Next
                        Next
                        y = y + 8
                        Dim allarray = all.Trim.Split("Fee")

                        Dim index = 0
                        For ii = 0 To 26
                            If allarray(ii).Replace("ee", "").Contains(":") Then
                                feevalues(index) = allarray(ii).Replace("ee", "").Split(":")(1)
                                index = index + 1
                            End If

                        Next
                        rd = rd + 1
                        For iii = 0 To 26
                            xlWorkSheet.Cells(rd, iii + 4) = feevalues(iii)
                        Next
                        xlWorkSheet.Cells(rd, 2) = id
                        xlWorkSheet.Cells(rd, 3) = description
                        Label3.Invoke(Sub()
                                          Label3.Text = y

                                      End Sub)

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


        xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\adacode2.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()



        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        MsgBox("done")

    End Sub


End Class
