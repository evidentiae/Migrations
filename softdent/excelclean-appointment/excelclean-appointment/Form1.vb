Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Text
Imports System
Imports System.IO
Imports System.Threading

Public Class Form1
    Dim thread As New Thread(AddressOf exec)
    Dim thread1 As New Thread(AddressOf exec1)
    Dim thread2 As New Thread(AddressOf exec2)
    Dim thread3 As New Thread(AddressOf exec3)
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



            xlWorkSheet.Cells(1, 2) = "appdate"
            xlWorkSheet.Cells(1, 3) = "datemade"
            xlWorkSheet.Cells(1, 4) = "time"
            xlWorkSheet.Cells(1, 5) = "confirm dat"
            xlWorkSheet.Cells(1, 7) = "reason"
            xlWorkSheet.Cells(1, 8) = "patDR"
            xlWorkSheet.Cells(1, 9) = "Med alerts"
            xlWorkSheet.Cells(1, 10) = "appt notes"



            ProgressBar1.Invoke(Sub()
                                    ProgressBar1.Minimum = 0
                                    ProgressBar1.Maximum = book.Sheets.Count
                                End Sub)
            Dim rd As Integer = 1 ' rownumberindestination
            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)

                Dim appdate As String = ""
                For h = 1 To sheet.UsedRange.Columns.Count
                    If Not sheet.Cells(3, h).value Is Nothing Then
                        If IsDate(sheet.Cells(3, h).value.ToString) Then
                            appdate = sheet.Cells(3, h).value.ToString
                            'MsgBox(appdate)
                        End If

                    End If
                Next

                Dim reas As Integer = 0
                For u = 0 To sheet.UsedRange.Columns.Count
                    ' MsgBox(sheet.Cells(4, u + 1).value.ToString)
                    If sheet.Cells(4, u + 1).value.ToString = "REASON" Then
                        reas = u
                        Exit For
                    End If

                Next
                If reas = 0 Then
                    MsgBox("error 'reason' was not found")
                End If


                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Value = i
                                    End Sub)

                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)


                    If Not sheet.Cells(y, 2).value = Nothing Then

                        'MsgBox(sheet.Cells(y, 2).value.ToString)

                        Dim beginrow = sheet.Cells(y, 2).value.ToString.Split

                        If IsDate(beginrow(0) & "m") Then
                            rd = rd + 1
                            xlWorkSheet.Cells(rd, 2) = appdate 'appdate


                            Dim datemade = sheet.Cells(y + 1, 2).value.ToString.Split(":")
                            xlWorkSheet.Cells(rd, 3) = datemade(1) 'datemade
                            'MsgBox(datemade(1))
                            xlWorkSheet.Cells(rd, 4) = beginrow(0) 'time
                            ' MsgBox(beginrow(0))
                            Dim confirmdate
                            Dim temppchosen = 0
                            If Not sheet.Cells(y, 8).value = Nothing Then
                                confirmdate = sheet.Cells(y, 8).value.ToString.Split
                                temppchosen = 8
                            ElseIf Not sheet.Cells(y, 9).value = Nothing Then
                                confirmdate = sheet.Cells(y, 9).value.ToString.Split
                                temppchosen = 9
                            Else
                                confirmdate = sheet.Cells(y, 10).value.ToString.Split
                                temppchosen = 10
                            End If

                            For tt = temppchosen To 9
                                If IsDate(confirmdate(0)) Then
                                    Exit For
                                Else
                                    confirmdate = sheet.Cells(y, tt + 1).value.ToString.Split
                                End If
                            Next
                            If IsDate(confirmdate(0)) Then
                                xlWorkSheet.Cells(rd, 5) = confirmdate(0) 'confirm dat
                            End If


                            xlWorkSheet.Cells(rd, 6) = beginrow(1) ' ID


                            'If Not sheet.Cells(y, reas + 1).value = Nothing Then
                            'xlWorkSheet.Cells(rd, 7) = sheet.Cells(y, reas + 1).value.ToString  'reason
                            'End If
                            ' If Not sheet.Cells(y, reas + 2).value = Nothing Then
                            'xlWorkSheet.Cells(rd, 8) = sheet.Cells(y, reas + 2).value.ToString 'patDR
                            'End If
                            Dim h As Integer
                            Try




                                For h = sheet.UsedRange.Columns.Count To 1 Step -1
                                    If Not sheet.Cells(y, h).value Is Nothing Then
                                        If IsNumeric(sheet.Cells(y, h).value.ToString) Then

                                            xlWorkSheet.Cells(rd, 8) = sheet.Cells(y, h).value.ToString 'patDR
                                            xlWorkSheet.Cells(rd, 7) = sheet.Cells(y, h - 1).value.ToString  'reason
                                            Exit For
                                        Else
                                            xlWorkSheet.Cells(rd, 8) = sheet.Cells(y, h).value.ToString.Substring(sheet.Cells(y, h).value.ToString.LastIndexOf(" "))
                                            Dim alpha = sheet.Cells(y, h).value.ToString.Replace(sheet.Cells(y, h).value.ToString.Substring(sheet.Cells(y, h).value.ToString.LastIndexOf(" ")), "")
                                            Dim ou As Integer = 0
                                            While ou < alpha.Length

                                                If IsNumeric(alpha.Chars(ou)) Or alpha.Chars(ou) = "." Then

                                                    alpha = alpha.Remove(0, 1)
                                                    ou = ou - 1

                                                End If
                                                ou = ou + 1
                                            End While
                                            xlWorkSheet.Cells(rd, 7) = alpha  'reason
                                            Exit For
                                        End If
                                    End If
                                Next
                            Catch ex As Exception
                                MsgBox(ex.ToString)

                            End Try

                        ElseIf beginrow.Length > 1 Then
                            Dim comp As String = sheet.Cells(y, 2).value.ToString
                            If comp.Contains("MED. ALERTS") Then

                                Dim med = Split(comp, ":", 2)
                                xlWorkSheet.Cells(rd, 9) = med(1) 'Med alerts

                            ElseIf comp.Contains("Appt. Notes") Then

                                Dim appt = comp.Split(":")
                                xlWorkSheet.Cells(rd, 10) = appt(1) 'appt notes

                            End If
                        End If
                    End If
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


            xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\app254.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)






        Next



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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Label1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        thread1.Start()
    End Sub
    Private Sub exec1()
        'filling dr,hug,asst,op
        'this file was converted by adobe and has extension xlsx

        Dim Dttbl As New System.Data.DataTable

        'destination file
        Dim xls1 As New Excel.Application
        Dim book1 As Excel.Workbook
        Dim sheet1 As Excel.Worksheet

        'source file

        Dim xls As New Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        For Each file In OpenFileDialog1.FileNames
            Label1.Invoke(Sub()
                              Label1.Text = file
                          End Sub)



            xls.Workbooks.Open(file)
            xls1.Workbooks.Open(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\app254.xlsx")
            'get references to first workbook and worksheet

            book = xls.ActiveWorkbook
            sheet = book.ActiveSheet ' this is just to initialize / no need

            book1 = xls1.ActiveWorkbook
            sheet1 = book1.ActiveSheet ' this is just to initialize / no need



            sheet1.Cells(1, 11) = "dr"
            sheet1.Cells(1, 12) = "hug"
            sheet1.Cells(1, 13) = "asst"
            sheet1.Cells(1, 14) = "op"


            Dim rd As Integer = 1 ' rownumberindestination
            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)

                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Minimum = 0
                                        ProgressBar1.Maximum = sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
                                    End Sub)

                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                    ProgressBar1.Invoke(Sub()
                                            ProgressBar1.Value = y
                                        End Sub)

                    If Not sheet.Cells(y, 1).value = Nothing Then

                        'MsgBox(sheet.Cells(y, 2).value.ToString)

                        Dim beginrow = sheet.Cells(y, 1).value.ToString.Trim.Split


                        If IsDate(beginrow(0) & "m") Then
                            rd = rd + 1

                            Dim alltext = ""

                            For p = 1 To sheet.Columns.Count


                                If Not sheet.Cells(y, p).value = Nothing Then

                                    alltext = alltext & " " & sheet.Cells(y, p).value.ToString
                                ElseIf p > 54 Then

                                    Exit For

                                End If

                            Next

                            alltext = alltext.Replace("bew", "")
                            Dim checking = alltext.Trim.Split

                            If checking(1) = getvalue(sheet1, rd, 6) Then

                                alltext = alltext.Substring(returnsecondslash(alltext))

                                alltext = alltext.Replace(vbLf, " ")
                                alltext = alltext.Replace(vbCrLf, " ")
                                alltext = alltext.Replace(Environment.NewLine, " ")
                                alltext = alltext.Replace("\n", " ")

                                Dim arr = alltext.Split(" ")
                                ' If rd = 179 Then
                                'MsgBox(alltext)
                                ' End If

                                Dim four As Integer = 0
                                If IsNumeric(arr(0).Trim.Replace("/", "")) Then

                                    If Val(arr(0).Replace("/", "")) > 100 And Val(arr(0).Replace("/", "")) < 1000 Then 'if it is on 3 digit
                                        four = 1
                                    End If

                                End If
                                For u = 0 To arr.Length - 1

                                    If IsNumeric(arr(u)) Then
                                        'If rd = 179 Then
                                        'MsgBox(arr(0))
                                        'End If


                                        If Not four = 0 Then ' ignoring first value
                                            sheet1.Cells(rd, four + 10) = arr(u)

                                        End If

                                        If four = 4 Then
                                            Exit For
                                        End If
                                        four = four + 1
                                    End If
                                Next
                                Label3.Invoke(Sub()
                                                  Label3.Text = y
                                              End Sub)

                            Else


                                sheet1.Cells(rd, 11) = "error1"
                                sheet1.Cells(rd, 12) = "error1"
                                sheet1.Cells(rd, 13) = "error1"
                                sheet1.Cells(rd, 14) = "error1"
                                sheet1.Cells(rd, 15) = checking(1)
                                ' MsgBox("first " & alltext)
                            End If
                        ElseIf beginrow(0) = "TIME" Then

                            If sheet.Cells(y, 1).value.ToString.Contains("REASON") And sheet.Cells(y, 1).value.ToString.Contains("Pat Dr ") Then
                                For l = 0 To beginrow.Length - 1
                                    If IsDate(beginrow(l) & "m") Then
                                        rd = rd + 1

                                        Dim alltext = sheet.Cells(y, 1).value.ToString.Substring(sheet.Cells(y, 1).value.ToString.IndexOf("Pat Dr ") + 7)


                                        alltext = alltext.Replace("bew", "")
                                        Dim checking = alltext.Trim.Split

                                        If checking(1) = getvalue(sheet1, rd, 6) Then



                                            While alltext.Length > 1 And alltext.Contains("/")


                                                If alltext.Substring(alltext.IndexOf("/")).IndexOf("/") - alltext.Substring(alltext.IndexOf("/") + 1).IndexOf("/") = -2 Then
                                                    alltext = alltext.Substring(alltext.IndexOf("/") + 1)
                                                    alltext = alltext.Substring(alltext.IndexOf("/"))
                                                    Exit While

                                                Else
                                                    alltext = alltext.Substring(alltext.IndexOf("/") + 1)
                                                    'MsgBox(alltext)
                                                End If
                                            End While

                                            Dim arr = alltext.Split(" ")
                                            Dim four As Integer = 0
                                            For u = 0 To arr.Length - 1
                                                If IsNumeric(arr(u)) Then
                                                    If Not four = 0 Then ' ignoring first value
                                                        sheet1.Cells(rd, four + 10) = arr(u)
                                                    End If

                                                    If four = 4 Then
                                                        Exit For
                                                    End If
                                                    four = four + 1
                                                End If
                                            Next
                                            Label3.Invoke(Sub()
                                                              Label3.Text = y
                                                          End Sub)

                                        Else
                                            sheet1.Cells(rd, 11) = "error"
                                            sheet1.Cells(rd, 12) = "error"
                                            sheet1.Cells(rd, 13) = "error"
                                            sheet1.Cells(rd, 14) = "error"
                                            sheet1.Cells(rd, 15) = checking(1)
                                            'MsgBox("secondt " & alltext)
                                        End If
                                        Exit For
                                    End If
                                Next


                            End If
                        End If
                        Label3.Invoke(Sub()
                                          Label3.Text = y
                                      End Sub)

                        Application.DoEvents()
                    End If
                Next



                Label2.Invoke(Sub()
                                  Label2.Text = i
                              End Sub)
            Next

            xls.Workbooks.Close()
            xls.Quit()

            releaseObject(sheet)
            releaseObject(book)
            releaseObject(xls)





            releaseObject(sheet1)
            releaseObject(book1)
            releaseObject(xls1)

            book1.Save()
            book1.Close()
            'xls1.Quit()





        Next



        MsgBox("done")
    End Sub

    Private Sub exec2()
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

            ProgressBar1.Invoke(Sub()
                                    ProgressBar1.Minimum = 0
                                    ProgressBar1.Maximum = book.Sheets.Count
                                End Sub)

            Dim rd As Integer = 1 ' rownumberindestination
            For i = 1 To book.Sheets.Count
                sheet = book.Sheets(i)

                Dim appdate As String = ""




                Dim IDD = ""


                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Value = i
                                    End Sub)
                Dim first = True
                For y = 1 To (sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)
                    Label3.Invoke(Sub()
                                      Label3.Text = y
                                  End Sub)

                    If Not sheet.Cells(y, 2).value = Nothing Then

                        'MsgBox(sheet.Cells(y, 2).value.ToString)

                        Dim beginrow = sheet.Cells(y, 2).value.ToString.Split

                        If IsDate(beginrow(0) & "m") Then
                            rd = rd + 2
                            IDD = beginrow(1) ' ID
                            xlWorkSheet.Cells(rd, 2) = IDD
                            first = True

                        ElseIf getvalue(sheet, y, 2).Contains("Code") Then
                            If Not first = True Then
                                rd = rd + 1
                            End If
                            first = False
                            Dim code = "", tooth = "", surface = ""
                            For ii = 2 To 17
                                If getvalue(sheet, y, ii).Contains("Code") And getvalue(sheet, y, ii).Contains(":") Then
                                    Dim texts = Split(getvalue(sheet, y, ii), ":", 2)
                                    If texts.Length > 1 Then
                                        code = texts(1)
                                    End If
                                    code = code & " " & getvalue(sheet, y, ii + 1)
                                ElseIf getvalue(sheet, y, ii).Contains("Tooth Number") And getvalue(sheet, y, ii).Contains(":") Then
                                    Dim texts = Split(getvalue(sheet, y, ii), ":", 2)
                                    If texts.Length > 1 Then
                                        tooth = texts(1)
                                    End If
                                    tooth = tooth & " " & getvalue(sheet, y, ii + 1)
                                ElseIf getvalue(sheet, y, ii).Contains("Tooth Surface") And getvalue(sheet, y, ii).Contains(":") Then
                                    Dim texts = Split(getvalue(sheet, y, ii), ":", 2)
                                    If texts.Length > 1 Then
                                        surface = texts(1)
                                    End If
                                    surface = surface & " " & getvalue(sheet, y, ii + 1)
                                End If

                            Next

                            tooth = tooth.Replace("Tooth Surface", "")
                            If surface <> "" Then
                                tooth = tooth.Replace(surface, "")
                            End If

                            xlWorkSheet.Cells(rd, 2) = IDD

                            xlWorkSheet.Cells(rd, 3) = code
                            xlWorkSheet.Cells(rd, 4) = tooth

                            xlWorkSheet.Cells(rd, 5) = surface


                        End If
                    End If
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


            xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\app254-1.xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)


        Next



        MsgBox("done")

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Label1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        thread2.Start()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        thread3.Start()
    End Sub
    Private Sub exec3()
        'filling dr,hug,asst,op
        'this file was converted by adobe and has extension xlsx

        Dim Dttbl As New System.Data.DataTable

        'destination file
        Dim xls1 As New Excel.Application
        Dim book1 As Excel.Workbook
        Dim sheet1 As Excel.Worksheet

        'source file

        Dim xls As New Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        For Each file In OpenFileDialog1.FileNames
            Label1.Invoke(Sub()
                              Label1.Text = file
                          End Sub)



            xls.Workbooks.Open(file)
            xls1.Workbooks.Open(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\app254-1.xlsx")
            'get references to first workbook and worksheet

            book = xls.ActiveWorkbook
            sheet = book.ActiveSheet ' this is just to initialize / no need

            book1 = xls1.ActiveWorkbook
            sheet1 = book1.ActiveSheet ' this is just to initialize / no need
            sheet1.Cells(1, 2) = "ID"
            sheet1.Cells(1, 3) = "CODE"
            sheet1.Cells(1, 4) = "TOOTH"
            sheet1.Cells(1, 5) = "SURFACE"
            sheet1.Cells(1, 6) = "APPDATE"
            sheet1.Cells(1, 7) = "DATEMADE"
            sheet1.Cells(1, 8) = "TIME"
            sheet1.Cells(1, 9) = "CONFIRM DATE"
            sheet1.Cells(1, 10) = "ID"
            sheet1.Cells(1, 11) = "REASON"
            sheet1.Cells(1, 12) = "PATDR"
            sheet1.Cells(1, 13) = "MED ALERTS"
            sheet1.Cells(1, 14) = "APPT NOTES"
            sheet1.Cells(1, 15) = "DR"
            sheet1.Cells(1, 16) = "HUG"
            sheet1.Cells(1, 17) = "ASST"
            sheet1.Cells(1, 18) = "OP"


            Dim sheetarray(sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row, 13) As String
            For l = 2 To sheet.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
                For o = 2 To 14
                    sheetarray(l - 2, o - 2) = getvalue(sheet, l, o)
                    If sheet.Cells(l, o).Value IsNot Nothing Then
                        sheetarray(l - 2, o - 2) = sheet.Cells(l, o).value.ToString
                    Else
                        sheetarray(l - 2, o - 2) = ""
                    End If

                Next
                Label3.Invoke(Sub()
                                  Label3.Text = l
                              End Sub)
            Next

            book.Close()
            Dim rd As Integer = -1 ' rownumberindestination
            For i = 1 To book1.Sheets.Count
                    sheet = book1.Sheets(i)

                    ProgressBar1.Invoke(Sub()
                                            ProgressBar1.Minimum = 0
                                            ProgressBar1.Maximum = sheet1.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
                                        End Sub)

                For y = 1 To (sheet1.Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row)

                    ProgressBar1.Invoke(Sub()
                                            ProgressBar1.Value = y
                                        End Sub)
                    If sheet1.Cells(y, 2).value IsNot Nothing Then


                        If IsNumeric(sheet1.Cells(y, 2).value.ToString) Then
                            rd = rd + 1
                            'MsgBox(sheet.Cells(y, 2).value.ToString)

                            While sheet1.Cells(y, 2).value IsNot Nothing
                                For f = 0 To 12

                                    sheet1.Cells(y, f + 6) = sheetarray(rd, f)
                                Next
                                y = y + 1
                                Label3.Invoke(Sub()
                                                  Label3.Text = y
                                              End Sub)
                                ProgressBar1.Invoke(Sub()
                                                        ProgressBar1.Value = y
                                                    End Sub)
                            End While


                            Label3.Invoke(Sub()
                                              Label3.Text = y
                                          End Sub)
                        End If
                    End If
                Next



                Label2.Invoke(Sub()
                                      Label2.Text = i
                                  End Sub)
                Next

                xls.Workbooks.Close()
                xls.Quit()

                releaseObject(sheet)
                releaseObject(book)
                releaseObject(xls)





                releaseObject(sheet1)
                releaseObject(book1)
                releaseObject(xls1)

                book1.Save()
                book1.Close()
                'xls1.Quit()





            Next



            MsgBox("done")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Label1.Text = OpenFileDialog1.FileName
        End If
    End Sub
End Class
