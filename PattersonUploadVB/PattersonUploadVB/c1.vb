Imports System.IO
Imports System.Net
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Leadtools.Dicom
Imports Leadtools.Dicom.DicomDataSet
Imports System.Reflection
Imports Leadtools
Imports System.Net.Mail
Imports System.Globalization

Public Class c1
    Public Const MedicalServerKey As String = ""

    Function sendemail(from As String, too As String, body As String) As Boolean
        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(from, "Bewbew1989")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.ipage.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress(from)
            e_mail.To.Add(too)
            e_mail.Subject = "Error Upload"
            e_mail.IsBodyHtml = False
            e_mail.Body = body
            Smtp_Server.Send(e_mail)
            Return True
        Catch error_t As Exception
            log("error code ; 666125 sending email" & error_t.Message)
            Return False
        End Try
    End Function

    Function log(s As String) As Boolean
        My.Computer.FileSystem.WriteAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "/logEV.txt", s & vbNewLine, True)
        Return True
    End Function

    Public Shared Function SetLicense(ByVal silent As Boolean) As Boolean
        Try
            Dim licenseFilePath As String = Application.StartupPath & "\LEADTOOLS.LIC"
            Dim developerKey As String = "0Mx8XtVTo5btbTbDM9FCk8KRAsBZLgAIsN2p3rlplDOpoL/YFmPX1zOYtXqhCnFE"
            RasterSupport.SetLicense(licenseFilePath, developerKey)
        Catch ex As Exception
            Dim c As New c1
            c.log("error code ; 666126 lciense leadtools" & ex.Message)
        End Try

        If RasterSupport.KernelExpired Then
            Dim licenseName = "LEADTOOLS.LIC"
            Dim keyName = "LEADTOOLS.LIC.KEY"
            Dim _assembly As Assembly
            Dim developerKey As String = Nothing
            Dim licenseFile As Byte() = Nothing

            Try
                _assembly = Assembly.GetExecutingAssembly()

                Using _imageStream As Stream = _assembly.GetManifestResourceStream(keyName)

                    If _imageStream IsNot Nothing Then

                        Using reader As StreamReader = New StreamReader(_imageStream)

                            If reader IsNot Nothing Then
                                developerKey = reader.ReadToEnd()
                            End If
                        End Using
                    End If
                End Using

                Using _imageStream As Stream = _assembly.GetManifestResourceStream(licenseName)

                    If _imageStream IsNot Nothing Then
                        Dim br As BinaryReader = New BinaryReader(_imageStream)
                        licenseFile = New Byte(_imageStream.Length - 1) {}
                        _imageStream.Read(licenseFile, 0, licenseFile.Length)
                        br.Close()
                        _imageStream.Close()
                    End If
                End Using

            Catch ex As Exception
                licenseFile = Nothing
                Dim c As New c1
                c.log("error code ; 666127 lciense leadtools" & ex.Message)
            End Try

            If (developerKey IsNot Nothing) AndAlso (licenseFile IsNot Nothing) Then

                Try
                    RasterSupport.SetLicense(licenseFile, developerKey)
                Catch ex As Exception
                    Dim c As New c1
                    c.log("error code ; 666128 lciense leadtools" & ex.Message)
                End Try
            End If
        End If

        If RasterSupport.KernelExpired Then

            If silent = False Then
                Dim msg As String = "Your license file is missing, invalid or expired. LEADTOOLS will not function. Please contact LEAD Sales for information on obtaining a valid license."
                Dim logmsg As String = String.Format("*** NOTE: {0} ***{1}", msg, Environment.NewLine)
                System.Diagnostics.Debugger.Log(0, Nothing, "*******************************************************************************" & Environment.NewLine)
                System.Diagnostics.Debugger.Log(0, Nothing, logmsg)
                System.Diagnostics.Debugger.Log(0, Nothing, "*******************************************************************************" & Environment.NewLine)
                MessageBox.Show(Nothing, msg, "No LEADTOOLS License", MessageBoxButtons.OK, MessageBoxIcon.[Stop])
                System.Diagnostics.Process.Start("https://www.leadtools.com/downloads/evaluation-form.asp?evallicenseonly=true")
            End If

            Return False
        End If

        Return True
    End Function

    Public Shared Function SetLicense() As Boolean
        Return SetLicense(False)
    End Function

    Function login(user As String, pass As String, realm As String) As CookieContainer
        Dim logincookie As CookieContainer
        Try
            Dim postData As String = "{""username"":""" & user & """,""password"":""" & pass & """}"
            Dim tempCookies As New CookieContainer
            Dim encoding As New UTF8Encoding
            Dim byteData As Byte() = encoding.GetBytes(postData)

            Dim postReq As HttpWebRequest = DirectCast(WebRequest.Create("https://" & realm & "/api/auth/login"), HttpWebRequest)
            postReq.Method = "POST"
            postReq.KeepAlive = True
            postReq.CookieContainer = tempCookies
            postReq.ContentType = "application/json; charset=utf-8"
            postReq.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 5.1; de; rv:1.9.0.13) Gecko/2009073022 Firefox/3.0.13"
            postReq.ContentLength = byteData.Length

            Dim postreqstream As Stream = postReq.GetRequestStream()
            postreqstream.Write(byteData, 0, byteData.Length)
            postreqstream.Close()
            Dim postresponse As HttpWebResponse

            postresponse = DirectCast(postReq.GetResponse(), HttpWebResponse)
            tempCookies.Add(postresponse.Cookies)
            logincookie = tempCookies
            Dim postreqreader As New StreamReader(postresponse.GetResponseStream())

            Dim answer As String = postreqreader.ReadToEnd

            Dim rss As JObject = JObject.Parse(answer)

            Dim p As New connectionStatus
            p = rss.ToObject(Of connectionStatus)
            If Not p.success Then
                log("failed login (" & DateTime.Now & ") : " & p.message)
                Return Nothing
            End If
            log("code ; 555000 success login (" & DateTime.Now & ") : " & p.message)
            Return logincookie
        Catch ex As Exception
            log("error code ; 666130 login (" & DateTime.Now & ") : " & ex.Message)
            Return Nothing
        End Try


    End Function

    Function cleanName(name As String) As String
        If name.Contains("(") And name.Contains(")") Then
            name = name.Remove(name.IndexOf("("), name.IndexOf(")") - name.IndexOf("(") + 1).Trim
        End If
        If name.Contains("(") Then
            name = name.Replace("(", "")
        End If

        name = name.Replace("(", "").Replace(")", "").Replace(".", "").Replace(",", "")

        Return name
    End Function

    Function getPatientuui(realm As String, cookie As CookieContainer, name As String, dob As String, inactive As Boolean, country As String) As String
        Try
            Dim nameSplitted = cleanName(name)
            If cleanName(name).Split(" ").Count > 2 Then
                nameSplitted = cleanName(name).Split(" ")(0) & " " & cleanName(name).Split(" ")(2)  'might be changed to index 1 
            End If

            Dim getReq As HttpWebRequest = DirectCast(WebRequest.Create("https://" & realm & "/api/kois/patient/recent/1?query=" & nameSplitted & "&inactive=" & inactive), HttpWebRequest)
            getReq.Method = "GET"
            getReq.KeepAlive = True
            getReq.CookieContainer = cookie
            getReq.ContentType = "application/json; charset=utf-8"
            getReq.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 5.1; de; rv:1.9.0.13) Gecko/2009073022 Firefox/3.0.13"
            Dim getresponse As HttpWebResponse
            getresponse = DirectCast(getReq.GetResponse(), HttpWebResponse)
            Dim getreqreader As New StreamReader(getresponse.GetResponseStream())
            Dim answer As String = getreqreader.ReadToEnd()

            Dim j As JArray = JArray.Parse(answer)
            Dim j1 = j.ToArray
            If j1.Count = 0 Then
                log("code ; 555565 no patients found for name:" & name & " dob:" & dob & " inactive: " & inactive)
                Return Nothing
            End If
            Dim dupcount = 0
            Dim dupvalue As Integer
            For i = 0 To j1.Count - 1
                Dim rss As JObject = JObject.Parse(j1(i).ToString)
                Dim p As New patientaddress
                p = rss.ToObject(Of patientaddress)
                'MsgBox(p.patient.patientInfo.dob.dateString(country))
                'MsgBox(dob)
                'MsgBox(p.patient.patientInfo.lastName & " " & p.patient.patientInfo.firstName)
                'MsgBox(name)
                If (p.patient.patientInfo.lastName & " " & p.patient.patientInfo.firstName).ToLower = name.ToLower Then
                    'MsgBox(p.patient.patientInfo.dob.dateString)
                    'MsgBox(dob)
                    If p.patient.patientInfo.dob.dateString(country) = dob Or j.Count = 1 Then

                        log("code ; 555898 Patient found uuid:" & p.patient.uuid & "; name:" & name & " inactive: " & inactive)
                        Return p.patient.uuid
                    ElseIf getdayfromstring(p.patient.patientInfo.dob.dateString(country)) = getmonthfromstring(dob) _
                         And getMonthFromString(p.patient.patientInfo.dob.dateString(country)) = getDayFromString(dob) _
                         And getYearFromString(p.patient.patientInfo.dob.dateString(country)) = getYearFromString(dob) Then
                        log("code ; 555738 Patient found with same name but different dob uuid:" & p.patient.uuid & "; name:" & name & "dicom dob:" & dob & "recieved dob: " & p.patient.patientInfo.dob.dateString(country) & " inactive: " & inactive)
                        Return p.patient.uuid
                    Else
                        log("code ; 555007 Patient duplicate the date recieved is " & p.patient.patientInfo.dob.dateString(country) & " dicom dob:" & dob & " uuid:" & p.patient.uuid & "; name:" & name & " inactive: " & inactive)
                        dupcount = dupcount + 1
                        dupvalue = i

                    End If
                ElseIf j.Count = 1 And p.patient.patientInfo.dob.dateString(country) = dob Then
                    log("code ; 555818 Patient found with different names uuid:" & p.patient.uuid & "; name:" & name & " inactive: " & inactive)

                    Return p.patient.uuid
                ElseIf j.Count = 1 And getdayfromstring(p.patient.patientInfo.dob.dateString(country)) = getmonthfromstring(dob) _
                    And getMonthFromString(p.patient.patientInfo.dob.dateString(country)) = getDayFromString(dob) _
                    And getYearFromString(p.patient.patientInfo.dob.dateString(country)) = getYearFromString(dob) Then
                    log("code ; 555838 Patient found with different names and dob uuid:" & p.patient.uuid & "; name:" & name & " dicom dob:" & dob & "recieved dob: " & p.patient.patientInfo.dob.dateString(country) & " inactive: " & inactive)

                    Return p.patient.uuid

                End If
            Next

            If dupcount = 1 And j.Count <> 1 Then
                Dim rss As JObject = JObject.Parse(j1(dupvalue).ToString)
                Dim p As New patientaddress
                p = rss.ToObject(Of patientaddress)
                log("code ; 555778 Patient found but not same dob uuid:" & p.patient.uuid & "; name:" & name & " inactive: " & inactive)
                Return p.patient.uuid
            End If
            log("code ; 555971 Patient not found with duplicate " & dupcount & "uuid:" & "; name:" & name & " inactive: " & inactive)
            Return ""
        Catch ex As Exception
            log("error code ; 666531 return uui (" & DateTime.Now & ") : " & ex.Message)
            Return ""
        End Try
    End Function

    Function getPatientID(realm As String, cookie As CookieContainer, name As String, dob As String, inactive As Boolean, country As String) As String
        Try
            Dim nameSplitted = cleanName(name)
            If cleanName(name).Split(" ").Count > 2 Then
                nameSplitted = cleanName(name).Split(" ")(0) & " " & cleanName(name).Split(" ")(2)  'might be changed to index 1 
            End If

            Dim getReq As HttpWebRequest = DirectCast(WebRequest.Create("https://" & realm & "/api/kois/patient/recent/1?query=" & nameSplitted & "&inactive=" & inactive), HttpWebRequest)
            getReq.Method = "GET"
            getReq.KeepAlive = True
            getReq.CookieContainer = cookie
            getReq.ContentType = "application/json; charset=utf-8"
            getReq.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 5.1; de; rv:1.9.0.13) Gecko/2009073022 Firefox/3.0.13"
            Dim getresponse As HttpWebResponse
            getresponse = DirectCast(getReq.GetResponse(), HttpWebResponse)
            Dim getreqreader As New StreamReader(getresponse.GetResponseStream())
            Dim answer As String = getreqreader.ReadToEnd()
            Dim j As JArray = JArray.Parse(answer)
            Dim j1 = j.ToArray
            If j1.Count = 0 Then
                log("code ; 444565 no patients found for name:" & name & " dob:" & dob & " inactive: " & inactive)
                Return Nothing
            End If
            Dim dupcount = 0
            Dim dupvalue As Integer
            For i = 0 To j1.Count - 1
                Dim rss As JObject = JObject.Parse(j1(i).ToString)
                Dim p As New patientaddress
                p = rss.ToObject(Of patientaddress)
                'MsgBox(p.patient.patientInfo.dob.dateString(country))
                'MsgBox(dob)
                'MsgBox(p.patient.patientInfo.lastName & " " & p.patient.patientInfo.firstName)
                'MsgBox(name)
                If (p.patient.patientInfo.lastName & " " & p.patient.patientInfo.firstName).ToLower = name.ToLower Then
                    'MsgBox(p.patient.patientInfo.dob.dateString)
                    'MsgBox(dob)
                    If p.patient.patientInfo.dob.dateString(country) = dob Or j.Count = 1 Then

                        log("code ; 444898 Patient found id:" & p.patient.id & "; name:" & name & " inactive: " & inactive)
                        Return p.patient.id
                    ElseIf getDayFromString(p.patient.patientInfo.dob.dateString(country)) = getMonthFromString(dob) _
                         And getMonthFromString(p.patient.patientInfo.dob.dateString(country)) = getDayFromString(dob) _
                         And getYearFromString(p.patient.patientInfo.dob.dateString(country)) = getYearFromString(dob) Then
                        log("code ; 444738 Patient found with same name but different dob id:" & p.patient.id & "; name:" & name & "dicom dob:" & dob & "recieved dob: " & p.patient.patientInfo.dob.dateString(country) & " inactive: " & inactive)
                        Return p.patient.id
                    Else
                        log("code ; 444007 Patient duplicate the date recieved is " & p.patient.patientInfo.dob.dateString(country) & " dicom dob:" & dob & " id:" & p.patient.id & "; name:" & name & " inactive: " & inactive)
                        dupcount = dupcount + 1
                        dupvalue = i

                    End If
                ElseIf j.Count = 1 And p.patient.patientInfo.dob.dateString(country) = dob Then
                    log("code ; 444818 Patient found with different names id:" & p.patient.id & "; name:" & name & " inactive: " & inactive)

                    Return p.patient.id
                ElseIf j.Count = 1 And getDayFromString(p.patient.patientInfo.dob.dateString(country)) = getMonthFromString(dob) _
                    And getMonthFromString(p.patient.patientInfo.dob.dateString(country)) = getDayFromString(dob) _
                    And getYearFromString(p.patient.patientInfo.dob.dateString(country)) = getYearFromString(dob) Then
                    log("code ; 444838 Patient found with different names and dob id:" & p.patient.id & "; name:" & name & " dicom dob:" & dob & "recieved dob: " & p.patient.patientInfo.dob.dateString(country) & " inactive: " & inactive)

                    Return p.patient.id

                End If
            Next

            If dupcount = 1 And j.Count <> 1 Then
                Dim rss As JObject = JObject.Parse(j1(dupvalue).ToString)
                Dim p As New patientaddress
                p = rss.ToObject(Of patientaddress)
                log("code ; 444778 Patient found but not same dob id:" & p.patient.id & "; name:" & name & " inactive: " & inactive)
                Return p.patient.id
            End If
            log("code ; 444971 Patient not found with duplicate " & dupcount & "id:" & "; name:" & name & " inactive: " & inactive)
            Return ""
        Catch ex As Exception
            log("error code ; 777531 return uui (" & DateTime.Now & ") : " & ex.Message)
            Return ""
        End Try
    End Function

    Function uploadImageDicom(DicomDataset As DicomDataSet, realm As String, cookie As CookieContainer, uuid As String) As Boolean
        Try

            Dim de2 As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.SeriesDescription, Nothing)
            Dim de3 As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.PatientID, Nothing)
            Dim de4 As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.SOPInstanceUID, Nothing)
            Dim de5 As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.AcquisitionDate, Nothing) 'remove this
            If de2 Is Nothing Then
                DicomDataset.InsertElement(de3, False, DicomTag.SeriesDescription, DicomVRType.SH, False, 0)
                de2 = DicomDataset.FindFirstElement(Nothing, DicomTag.SeriesDescription, Nothing)
            End If

            If de5 Is Nothing Then
            Else
                If DicomDataset.GetDateValue(de5, 0, 1)(0).ToDateTime < Convert.ToDateTime("01-01-1975") Then
                    Dim dttt(1) As Date
                    dttt(0) = Convert.ToDateTime("01-01-1975")
                    DicomDataset.SetDateValue(de5, dttt)
                End If

            End If

            DicomDataset.SetStringValue(de2, "Radiographs")
            DicomDataset.SetStringValue(de3, realm.Split(".")(0) & "-" & uuid)

            Dim mstream As MemoryStream = New MemoryStream()
            DicomDataset.Save(mstream, DicomDataSetSaveFlags.ExplicitVR Or DicomDataSetSaveFlags.MetaHeaderPresent)
            Dim dataSetBytes As Byte() = mstream.ToArray()
            'Dim encoding As New UTF8Encoding
            Dim byteData As Byte() = dataSetBytes


            Dim postReq As HttpWebRequest = DirectCast(WebRequest.Create("https://" & realm & "/api/kois/dicom/instances"), HttpWebRequest)
            postReq.Method = "POST"
            postReq.KeepAlive = True
            postReq.CookieContainer = cookie
            Dim s As Stream = postReq.GetRequestStream()
            s.Write(dataSetBytes, 0, dataSetBytes.Length)
            s.Flush()
            s.Close()
            postReq.ContentType = "application/octet-stream"
            postReq.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 5.1; de; rv:1.9.0.13) Gecko/2009073022 Firefox/3.0.13"
            Dim postresponse As HttpWebResponse
            postresponse = DirectCast(postReq.GetResponse(), HttpWebResponse)
            Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
            Dim answer As String = postreqreader.ReadToEnd()
            If answer.Contains("patientId") Then ' serialize thhis and get patient id in response
                log("code ; 999999 image uploaded for patient " & realm.Split(".")(0) & "-" & uuid & " SOP uuid " & DicomDataset.GetStringValue(de4, 0))
                Return True
            End If
            log("code ; 888777 image failed to upload for patient " & realm.Split(".")(0) & "-" & uuid & " SOP uuid " & DicomDataset.GetStringValue(de4, 0) & " with message: " & answer)
            Return False
        Catch ex As Exception
            log("error code ; 666131 upload image (" & DateTime.Now & ") : " & ex.Message)
            Return False
        End Try
    End Function

    Function getPatientInfoFromDicom(DicomDataset As DicomDataSet) As String()
        Dim s(2) As String
        s(0) = ""
        s(1) = ""
        Dim sopinstanceUID = ""

        Try
            Dim de As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.PatientName, Nothing)
            Dim de1 As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.PatientBirthDate, Nothing)
            Dim de2 As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.SOPInstanceUID, Nothing)

            s(0) = DicomDataset.GetStringValue(de, 0).Replace("^", " ")


            If de2 Is Nothing Then
                log("code ; 555336 no sop instance uuid for : name " & s(0))
            Else
                sopinstanceUID = DicomDataset.GetStringValue(de2, 0)

            End If

            If de1 Is Nothing Then
                log("code ; 555443 getting patient info from dicom of id " & sopinstanceUID & " (" & DateTime.Now & ") : name " & s(0) & " no dob ")
            Else
                If de1.Length > 0 Then
                    s(1) = DicomDataset.GetDateValue(de1, 0, 1)(0).ToDateTime.ToString("dd/MM/yyyy")
                    log("code ; 555444 getting patient info from dicom of id " & sopinstanceUID & " (" & DateTime.Now & ") : name " & s(0) & " dob " & s(1))
                End If
            End If



            Return s
        Catch ex As Exception
            log("error code ; 666124 getting patient info from dicom (" & DateTime.Now & ") : " & ex.Message)
            Return s
        End Try
    End Function

    Function getPatientInfoFromJPEG(path As String) As String()

        Try
            If path.Contains("-") And path.IndexOf("-") <> path.LastIndexOf("-") Then
                Return path.Substring(path.LastIndexOf("\") + 1, path.LastIndexOf(".") - path.LastIndexOf("\") - 1).Split("-")
            End If
        Catch ex As Exception
            log("error code ; 666184 getting patient info from JPEG (" & DateTime.Now & ") : " & ex.Message)
            Dim se(2) As String
            se(0) = ""
            se(1) = ""
            se(2) = ""
            Return se
        End Try

        Dim s(2) As String
        s(0) = ""
        s(1) = ""
        s(2) = ""
        Return s

    End Function

    Function ReadIDFile(path As String) As String()
        If Not File.Exists(path) Then
            Return Nothing
        End If
        Return File.ReadAllLines(path)
    End Function

    Function GetPatientName(oldID As String, arr As String()) As String()
        Dim s(2) As String
        s(0) = ""
        s(1) = ""
        For i = 0 To arr.Length - 1
            If arr(i).Contains(oldID) Then
                If arr(i).Split(",")(0).Replace("""", "").Trim = oldID Then
                    s(0) = arr(i).Split(",")(1).Replace("""", "").Trim
                    s(1) = arr(i).Split(",")(2).Replace("""", "").Trim
                    Return s
                End If
            End If
        Next
        Return s
    End Function

    Function renameFile(path As String, flag As Integer, ByRef path1 As String) As Boolean

        '-1 to remove our naming convention 
        '1 for uploaded
        '2 for ignored
        '3 for error 
        Try
            path1 = ""
            Dim newName = ""
            Dim newNameLong = ""
            If flag = -1 Then

                newName = path.Substring(path.LastIndexOf("\") + 1).Replace(".Oryx-Uploaded", "").Replace(".Oryx-Ignored", "").Replace(".Oryx-PNF", "").Replace(".Oryx-Error", "").Replace(".Oryx-Dup", "")
                newNameLong = path.Substring(0, path.LastIndexOf("\") + 1) & newName
                If newNameLong = path Then
                    path1 = path
                Else
                    My.Computer.FileSystem.RenameFile(path, newName)
                    path1 = newNameLong
                End If
                log("code ; 555869 renaming with flag & " & flag & " (" & DateTime.Now & ") : ")
                Return True
            ElseIf flag = 1 Then

                newName = path.Substring(path.LastIndexOf("\") + 1) & ".Oryx-Uploaded"
                newNameLong = path.Substring(0, path.LastIndexOf("\") + 1) & newName

                My.Computer.FileSystem.RenameFile(path, newName)
                path1 = newNameLong
                log("code ; 555871 renaming with flag & " & flag & " (" & DateTime.Now & ") : ")
                Return True
            ElseIf flag = 2 Then
                newName = path.Substring(path.LastIndexOf("\") + 1) & ".Oryx-Ignored"
                newNameLong = path.Substring(0, path.LastIndexOf("\") + 1) & newName

                My.Computer.FileSystem.RenameFile(path, newName)
                path1 = newNameLong
                log("code ; 555872 renaming with flag & " & flag & " (" & DateTime.Now & ") : ")
                Return True
            ElseIf flag = 3 Then
                newName = path.Substring(path.LastIndexOf("\") + 1) & ".Oryx-Error"
                newNameLong = path.Substring(0, path.LastIndexOf("\") + 1) & newName

                My.Computer.FileSystem.RenameFile(path, newName)
                path1 = newNameLong
                log("code ; 555873 renaming with flag & " & flag & " (" & DateTime.Now & ") : ")
                Return True
            ElseIf flag = 4 Then
                newName = path.Substring(path.LastIndexOf("\") + 1) & ".Oryx-PNF"
                newNameLong = path.Substring(0, path.LastIndexOf("\") + 1) & newName

                My.Computer.FileSystem.RenameFile(path, newName)
                path1 = newNameLong
                log("code ; 555874 renaming with flag & " & flag & " (" & DateTime.Now & ") : ")
                Return True
            ElseIf flag = 5 Then
                newName = path.Substring(path.LastIndexOf("\") + 1) & ".Oryx-Dup"
                newNameLong = path.Substring(0, path.LastIndexOf("\") + 1) & newName
                My.Computer.FileSystem.RenameFile(path, newName)
                path1 = newNameLong
                log("code ; 555875 renaming with flag & " & flag & " (" & DateTime.Now & ") : ")
                Return True
            End If

            Return False
        Catch ex As Exception
            log("error code ; 666794 renaming with flag & " & flag & " (" & DateTime.Now & ") : " & ex.Message)
            Return False
        End Try


    End Function

    Function IsUploaded(path As String) As Boolean
        If path.Contains(".Oryx-Uploaded") Then
            Return True
        End If
        Return False

    End Function

    Function loadFileExplorer(path As String, search As String) As Boolean

        Try

            Dim p As New ProcessStartInfo

            p.FileName = "runExplorerWithSerach.exe"

            p.Arguments = """" & path & """" & " " & """" & search & """"

            ' Use a hidden window
            p.WindowStyle = ProcessWindowStyle.Hidden

            ' Start the process
            Process.Start(p)
            log("code ; 555994 runExplorerWithserach.exe  " & path & search & " (" & DateTime.Now & ") : ")
        Catch ex As Exception
            log("code ; 666994 runExplorerWithserach.exe  " & path & search & " (" & DateTime.Now & ") : " & ex.Message)
        End Try


    End Function

    Function uploadImageJpeg(Path As String, realm As String, cookie As CookieContainer, id As String, filename As String, name As String) As Boolean
        Try



            Dim mstream As MemoryStream = New MemoryStream()
            Dim m As Image = Image.FromFile(Path)
            m.Save(mstream, Imaging.ImageFormat.Png)

            Dim dataSetBytes As Byte() = mstream.ToArray()

            'Dim stream As FileStream = File.OpenRead(Path)
            'Dim fileBytes As Byte() = New Byte(stream.Length - 1) {}
            'stream.Read(fileBytes, 0, fileBytes.Length)
            'stream.Close()

            'Dim encoding As New UTF8Encoding
            Dim byteData As Byte() = dataSetBytes

            MsgBox(name)
            MsgBox(filename)

            Dim str = String.Format("---------7e333ad708c6
Content-Disposition: form-data; name=""name""

" & name & "
---------7e333ad708c6
Content-Disposition: form-data; name=""template""

ancillary tests
---------7e333ad708c6
Content-Disposition: form-data; name=""file""; filename=""" & filename & """
Content-Type: image/png

")

            Dim byteData1 As Byte() = System.Text.Encoding.ASCII.GetBytes(str)
            Dim byteData2 As Byte() = System.Text.Encoding.ASCII.GetBytes("  

---------7e333ad708c6--")
            Dim one As Byte() = byteData1
            Dim two As Byte() = byteData
            Dim three As Byte() = byteData2
            Dim combined As Byte() = New Byte(one.Length + two.Length - 1) {}
            Dim combined1 As Byte() = New Byte(combined.Length + byteData2.Length - 1) {}

            For i As Integer = 0 To combined.Length - 1
                combined(i) = If(i < one.Length, one(i), two(i - one.Length))
            Next
            For i As Integer = 0 To combined1.Length - 1
                combined1(i) = If(i < combined.Length, combined(i), byteData2(i - combined.Length))
            Next




            Dim postReq As HttpWebRequest = DirectCast(WebRequest.Create("https://" & realm & "/api/kois/upload/document/patientId/" & id), HttpWebRequest)
            postReq.Method = "POST"
            postReq.KeepAlive = True
            postReq.CookieContainer = cookie

            Dim s As Stream = postReq.GetRequestStream()


            's.Write(dataSetBytes, 0, dataSetBytes.Length)
            s.Write(combined1, 0, combined1.Length)


            s.Flush()
            s.Close()

            postReq.ContentType = "multipart/form-data; boundary=-------7e333ad708c6"
            postReq.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.119 Safari/537.36"
            Dim postresponse As HttpWebResponse
            postresponse = DirectCast(postReq.GetResponse(), HttpWebResponse)
            Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
            Dim answer As String = postreqreader.ReadToEnd()
            If answer.Contains("patientId") Then ' serialize thhis and get patient id in response
                log("code ; 499999 image uploaded for patient " & realm.Split(".")(0) & "-" & id & " SOP uuid ") '& DicomDataSet.GetStringValue(de4, 0))
                Return True
                MsgBox("True")
            End If
            log("code ; 488777 image failed to upload for patient " & realm.Split(".")(0) & "-" & id & " SOP uuid ") '& DicomDataSet.GetStringValue(de4, 0) & " with message: " & answer)
            Return False
        Catch ex As Exception
            log("error code ; 466131 upload image (" & DateTime.Now & ") : " & ex.Message)
            'Return False
            MsgBox(ex.ToString)
        End Try
    End Function

    Function byteArray(filename As String)
        Dim oFileStream As FileStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
        Dim FileByteArrayData As Byte() = New Byte(oFileStream.Length - 1) {}
        oFileStream.Read(FileByteArrayData, 0, System.Convert.ToInt32(oFileStream.Length))
        oFileStream.Close()
        Return FileByteArrayData
    End Function

    Function getDayFromString(s As String) As String
        If s.IndexOf("/") = 2 And s.LastIndexOf("/") = 5 And s.Length = 10 Then
            Return s.Substring(0, 2)
        End If
        Return "00"
    End Function

    Function getMonthFromString(s As String) As String
        If s.IndexOf("/") = 2 And s.LastIndexOf("/") = 5 And s.Length = 10 Then
            Return s.Substring(3, 2)
        End If
        Return "00"
    End Function

    Function getYearFromString(s As String) As String
        If s.IndexOf("/") = 2 And s.LastIndexOf("/") = 5 And s.Length = 10 Then
            Return s.Substring(6, 4)
        End If
        Return "00"
    End Function

    Function hashFile(ByVal filepath As String) As String
        Try


            Using reader As New System.IO.FileStream(filepath, IO.FileMode.Open, IO.FileAccess.Read)
                Using md5 As New System.Security.Cryptography.MD5CryptoServiceProvider
                    Dim hash() As Byte = md5.ComputeHash(reader)
                    Return System.Text.Encoding.Unicode.GetString(hash)
                End Using
            End Using
        Catch ex As Exception
            log("error 168" & ex.Message)
        End Try
        Return ""
    End Function

    Function CompareFiles(ByVal file1FullPath As String, ByVal file2FullPath As String) As Boolean
        Try


            If Not File.Exists(file1FullPath) Or Not File.Exists(file2FullPath) Then
                'One or both of the files does not exist.
                Return False
            End If

            If file1FullPath = file2FullPath Then
                ' fileFullPath1 and fileFullPath2 points to the same file...
                Return True
            End If

            Try
                Dim file1Hash As String = hashFile(file1FullPath)
                Dim file2Hash As String = hashFile(file2FullPath)

                If file1Hash = file2Hash Then
                    Return True
                Else
                    Return False
                End If

            Catch ex As Exception
                Return False
            End Try
        Catch ex As Exception
            log("error 168" & ex.Message)
        End Try
    End Function

    Function removeDupFromDirectory(files As String(), path As String)
        Dim filesCount = files.Count
        For i = 0 To filesCount - 1

            Dim f = Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(Function(saa) saa.Split("-")(0) = files(i).Split("-")(0) And IO.Path.GetExtension(saa) <> ".Oryx-Dup" And saa <> files(i))

            For j = 0 To f.Count - 1
                If CompareFiles(f(j), files(i)) Then
                    Dim ty As String = ""
                    renameFile(files(i), 5, ty)
                    Exit For
                End If
            Next
        Next

    End Function


End Class
