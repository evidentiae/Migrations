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


            xlWorkSheet.Cells(1, 1) = "carrier"
            xlWorkSheet.Cells(1, 2) = "patient name"
            xlWorkSheet.Cells(1, 3) = "birth date"
            xlWorkSheet.Cells(1, 4) = "chart no"
            xlWorkSheet.Cells(1, 5) = "subscription no "
            xlWorkSheet.Cells(1, 6) = "subscriber chart no"
            xlWorkSheet.Cells(1, 7) = "renewal"
            xlWorkSheet.Cells(1, 8) = "address"
            xlWorkSheet.Cells(1, 9) = "phone"
            xlWorkSheet.Cells(1, 10) = "group name"
            xlWorkSheet.Cells(1, 11) = "group number"
            xlWorkSheet.Cells(1, 12) = "employer"



            Dim continu As Boolean = True

            Dim rd As Integer = 1 ' rownumberindestination
            Dim carr = ""
            Dim carrAddress = ""
            Dim carrGroupNum = ""
            Dim carrGroupName = ""
            Dim carrPhoneNumber = ""
            Dim employer = ""
            Dim subscriber = ""
            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)


                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Maximum = sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
                                        ProgressBar1.Minimum = 1
                                    End Sub)

                rd = rd + 1 ' two lines can be removed if you want to complete on new page 
                Dim col = 1


                Dim appdate As String = ""
                Dim dateindex = 0

                subscriber = ""
                Dim renewal = ""



                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                    continu = True


                    If getvalue(sheet, y, 2).Contains("CARRIER:") Then 'carrier /  group name / group number
                        carr = getvalue(sheet, y, 2)
                        If carr.Replace("CARRIER:", "").Length < 2 Then
                            carr = getvalue(sheet, y, 3)
                        End If
                        carrGroupName = (getvalue(sheet, y, 4) & " " & getvalue(sheet, y, 5)).Replace("GROUP NAME:", "").Replace("DED S/P/O:", "")
                        carrGroupNum = (getvalue(sheet, y + 1, 4) & " " & getvalue(sheet, y + 1, 5)).Replace("GROUP NUM:", "").Replace("IND,", "")
                        employer = (getvalue(sheet, y + 5, 4) & " " & getvalue(sheet, y + 5, 5)).Replace("EMPLOYER:", "").Replace("LAST UPDATE:", "")

                    End If


                    If getvalue(sheet, y, 2).Contains("ADDRESS:") Then 'address
                        Dim kj = 0
                        carrAddress = ""
                        While Not getvalue(sheet, y + kj, 2).Contains("PHONE:")
                            carrAddress = carrAddress & " " & getvalue(sheet, y + kj, 2) & getvalue(sheet, y + kj, 3)

                            kj = kj + 1
                        End While
                        carrAddress = carrAddress.Replace("GROUP NUM:", "").Replace("ADDRESS:", "").Replace("GROUP NAME:", "")

                    End If



                    If getvalue(sheet, y, 2).Contains("PHONE:") Then 'phone

                        carrPhoneNumber = (getvalue(sheet, y, 2) & " " & getvalue(sheet, y, 3)).Replace("PHONE:", "").Replace("CLAIM FORMAT:", "").Replace("TIME LIMIT: 0 days", "")

                    End If



                    If getvalue(sheet, y, 2).Contains("PATIENT NAME") Then
                        y = y + 1
                        For h = 3 To 7
                            If getvalue(sheet, y - 5, h).Length > 1 And getvalue(sheet, y - 5, h).Length < 5 Then
                                renewal = getvalue(sheet, y - 5, h)
                                Exit For
                            End If
                        Next

                        While Not getvalue(sheet, y, 2).Contains("CARRIER:")
                            Dim s = ""
                            For u = 2 To 7
                                s = s & " " & getvalue(sheet, y, u - 1)
                            Next

                            Dim k = s.Split
                            If k.Length = 0 Then
                                Exit While
                            End If
                            Dim col1 = ""
                            Dim col2 = ""
                            Dim col3 = ""
                            Dim col4 = ""
                            For p = 0 To k.Length - 1
                                If Not IsDate(k(p)) And Not IsNumeric(k(p)) Then
                                    col1 = col1 + " " + k(p)
                                Else
                                    If k.Length > p Then
                                        If IsDate(k(p)) Then
                                            col2 = k(p)
                                            p = p + 1
                                        End If

                                    End If

                                    If k.Length > p Then


                                        If k(p).Contains("12:00:00") Then
                                            p = p + 1
                                        End If
                                    End If

                                    If k.Length > p Then


                                        If k(p).Contains("AM") Then
                                            p = p + 1
                                        End If
                                    End If


                                    If k.Length > p Then
                                        col3 = k(p)
                                        p = p + 1
                                    End If
                                    If k.Length > p Then
                                        col4 = k(p)
                                        p = p + 1
                                    End If
                                    Exit For

                                End If

                            Next
                            If continu Then
                                subscriber = col3.Replace("Married", "").Replace("Single", "").Replace("Child", "")

                                continu = False
                            End If
                            xlWorkSheet.Cells(rd, 2) = col1.Replace("*", "").Replace("(P)", "").Replace("(S)", "")
                            xlWorkSheet.Cells(rd, 3) = col2
                            xlWorkSheet.Cells(rd, 4) = col3.Replace("Married", "").Replace("Single", "").Replace("Child", "")
                            xlWorkSheet.Cells(rd, 5) = col4.Replace("Married", "").Replace("Single", "").Replace("Child", "")
                            xlWorkSheet.Cells(rd, 6) = subscriber
                            xlWorkSheet.Cells(rd, 8) = carrAddress
                            xlWorkSheet.Cells(rd, 9) = carrPhoneNumber
                            xlWorkSheet.Cells(rd, 10) = carrGroupName
                            xlWorkSheet.Cells(rd, 11) = carrGroupNum
                            xlWorkSheet.Cells(rd, 12) = employer

                            ' For u = 2 To 11
                            'xlWorkSheet.Cells(rd, u) = getvalue(sheet, y, u - 1)
                            'Next
                            xlWorkSheet.Cells(rd, 1) = carr.Replace("GROUP NAME:", "").Replace("CARRIER:", "")
                            xlWorkSheet.Cells(rd, 7) = renewal


                            rd = rd + 1
                            y = y + 1

                        End While
                        y = y - 1
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


            xlWorkBook.SaveAs("C:\Users\developer\Desktop\bew\tmp\bew3.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
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
