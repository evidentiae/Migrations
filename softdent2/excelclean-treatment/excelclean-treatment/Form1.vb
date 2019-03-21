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


            xlWorkSheet.Cells(1, 2) = "id"
            xlWorkSheet.Cells(1, 3) = "plandate"
            xlWorkSheet.Cells(1, 4) = "code"
            xlWorkSheet.Cells(1, 5) = "Type"
            xlWorkSheet.Cells(1, 6) = "doctor"
            xlWorkSheet.Cells(1, 7) = "tooth"
            xlWorkSheet.Cells(1, 8) = "surface"
            xlWorkSheet.Cells(1, 9) = "patient"
            xlWorkSheet.Cells(1, 10) = "insurance"
            xlWorkSheet.Cells(1, 11) = "total"


            Dim i As Integer = 1
            While i <= book.Sheets.Count



                sheet = book.Sheets(i)


                    Label2.Invoke(Sub()
                                      Label2.Text = i

                                  End Sub)





                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                    Dim id = ""

                    If IsNumeric(getvalue(sheet, y, 2)) Then
                        id = getvalue(sheet, y, 2)

                        y = y + 3

                        While Not IsNumeric(getvalue(sheet, y, 2)) And y < (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                            If IsDate(getvalue(sheet, y, 2)) Then
                                rd = rd + 1

                                Dim plandate = "", code = "", type = "", doctor = "", tooth = "", surface = "", patient = "", insurance = "", total = ""
                                plandate = getvalue(sheet, y, 2)
                                Dim temp1 = getvalue(sheet, y, 3).Split
                                code = temp1(0)
                                If temp1.Length > 1 Then


                                    type = temp1(1)
                                    Dim temp2 = getvalue(sheet, y, 4).Split
                                    doctor = temp2(0)
                                    If temp2.Length > 1 Then
                                        tooth = temp2(1)
                                    End If
                                    If getvalue(sheet, y, 5).Length < 5 Then
                                        surface = getvalue(sheet, y, 5)

                                        patient = getvalue(sheet, y, 7)
                                        insurance = getvalue(sheet, y, 8)
                                        total = getvalue(sheet, y, 9)
                                    Else
                                        surface = ""
                                        Dim firstnumeric = 6
                                        For se = 5 To 9
                                            If IsNumeric(getvalue(sheet, y, se)) Then
                                                firstnumeric = se
                                                Exit For
                                            End If
                                        Next
                                        patient = getvalue(sheet, y, firstnumeric)
                                        insurance = getvalue(sheet, y, firstnumeric + 1)
                                        total = getvalue(sheet, y, firstnumeric + 2)
                                    End If

                                Else
                                    type = getvalue(sheet, y, 4)
                                    doctor = getvalue(sheet, y, 5)
                                    tooth = getvalue(sheet, y, 6)


                                    If getvalue(sheet, y, 7).Length < 5 Then
                                        surface = getvalue(sheet, y, 7)

                                        patient = getvalue(sheet, y, 9)
                                        insurance = getvalue(sheet, y, 10)
                                        total = getvalue(sheet, y, 11)
                                    Else
                                        surface = ""
                                        Dim firstnumeric = 8
                                        For se = 7 To 10
                                            If IsNumeric(getvalue(sheet, y, se)) Then
                                                firstnumeric = se
                                                Exit For
                                            End If
                                        Next

                                        patient = getvalue(sheet, y, firstnumeric)
                                        insurance = getvalue(sheet, y, firstnumeric + 1)
                                        total = getvalue(sheet, y, firstnumeric + 2)
                                    End If

                                End If
                                xlWorkSheet.Cells(rd, 2) = id
                                xlWorkSheet.Cells(rd, 3) = plandate
                                xlWorkSheet.Cells(rd, 4) = code
                                xlWorkSheet.Cells(rd, 5) = type
                                xlWorkSheet.Cells(rd, 6) = doctor
                                xlWorkSheet.Cells(rd, 7) = tooth
                                xlWorkSheet.Cells(rd, 8) = surface
                                xlWorkSheet.Cells(rd, 9) = patient
                                xlWorkSheet.Cells(rd, 10) = insurance
                                xlWorkSheet.Cells(rd, 11) = total
                            End If
                            y = y + 1
                            Label3.Invoke(Sub()
                                              Label3.Text = y

                                          End Sub)

                        End While
                        y = y - 1
                    End If





                Next

                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Value = i
                                    End Sub)


                i = i + 1

            End While
            book.Close()
            xls.Workbooks.Close()

        Next

        xls.Quit()

        releaseObject(sheet)
        releaseObject(book)
        releaseObject(xls)


        xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\treatment254.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()



        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        MsgBox("done")

    End Sub


End Class
