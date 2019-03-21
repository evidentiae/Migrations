Imports System.Globalization

Public Class dob
    Public year As Integer
    Public month As Integer
    Public day As Integer
    Function dateString(country) As String

        Try
            Dim m As String = month
            Dim d As String = day
            If month < 10 Then
                m = "0" & month
            End If
            If day < 10 Then
                d = "0" & day
            End If

            If IsDate(d & "/" & m & "/" & year) Then
                'Dim provider As New CultureInfo("en-US")
                'Return DateTime.ParseExact(year & "-" & m & "-" & d, "yyyy-MM-dd", provider).ToString("dd/MM/yyyy")
                If country = "US" Then
                    Return d & "/" & m & "/" & year
                ElseIf country = "CA" Then
                    Return d & "/" & m & "/" & year
                End If

            End If
            'Return m & "/" & d & "/" & year
            Return "01/01/1917"
        Catch ex As Exception
            Dim c1 As New c1
            c1.log("error code ; 666141 in datestring function in dob class " & ex.Message)
            Return "01/01/1917"
        End Try
    End Function

End Class
