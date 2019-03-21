
Imports System.IO
Imports System.Net
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Leadtools.Dicom
Imports Leadtools.Dicom.DicomDataSet
Imports System.Reflection
Imports Leadtools
Imports System.Threading
Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports Leadtools.Codecs
Imports Leadtools.ImageProcessing.Effects

Public Class Form1
    Dim c1 As New c1
    Dim th As New Thread(AddressOf bw_DoWork)
    Dim th1 As New Thread(AddressOf bw1_DoWork)
    Dim thLoad As New Thread(AddressOf bw_DoWorkLoad)
    Dim th1Load As New Thread(AddressOf bw1_DoWorkLoad)
    Dim lock = False
    Public Country = "US"
    Dim files() As String
    Public logincookie As CookieContainer
    Dim UploadPreviouslyUploaded = False
    Dim dateFormat As String = ""
    Dim DNF = Date.Now.ToString("dd-MM-yyyy")

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MsgBox(c1.CompareFiles("C:\Users\developer\Desktop\hoo.jpeg", "C:\Users\developer\Desktop\hoo1.jpeg"))
        dateFormat = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\International", True).GetValue("sShortDate")
        Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\International", True).SetValue("sShortDate", "dd/MM/yyyy")

        RadioButton1.Checked = True
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 100
        Label9.Text = ""
        th.Priority = ThreadPriority.Highest
        th1.Priority = ThreadPriority.Highest
        c1.SetLicense()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label6.Text = 0
        Label8.Text = 0

        'files = Directory.GetFiles(Label2.Text, "*", SearchOption.AllDirectories)
        'MsgBox(files.Count)
        ProgressBar1.Value = 0
        If th.ThreadState = ThreadState.Running Then
            Exit Sub
        End If
        If th1.ThreadState = ThreadState.Running Then
            Exit Sub
        End If

        If files Is Nothing Then
            MsgBox("No Files to Process")
            Exit Sub
        End If

        If files.Count < 1 Then
            MsgBox("No Files to Process")
            Exit Sub
        End If


        If th.ThreadState = ThreadState.Stopped Or th.ThreadState = ThreadState.Aborted Then
            th = New Thread(AddressOf bw_DoWork)
        End If
        If thLoad.ThreadState = ThreadState.Stopped Or thLoad.ThreadState = ThreadState.Aborted Then
            thLoad = New Thread(AddressOf bw_DoWorkLoad)
        End If


        th.Start()
        thLoad.Start()

    End Sub

    Private Sub bw_DoWork()
        Dim UploadedFiles = 0
        Dim ErrorFiles = 0
        Dim PNF = 0
        Dim IgnoredFile = 0
        Dim AlreadyUploaded = 0

        Label6.Invoke(Sub()
                          Label6.Text = UploadedFiles
                      End Sub)
        Label8.Invoke(Sub()
                          Label8.Text = ErrorFiles
                      End Sub)
        Label11.Invoke(Sub()
                           Label11.Text = IgnoredFile
                       End Sub)
        Label13.Invoke(Sub()
                           Label13.Text = PNF
                       End Sub)
        Label15.Invoke(Sub()
                           Label15.Text = AlreadyUploaded
                       End Sub)

        Try

            TextBox1.Invoke(Sub()
                                TextBox1.Text = "training-us.evidentiae.com"
                                TextBox2.Text = "salehrania@gmail.com"
                                TextBox3.Text = "Password1!"
                            End Sub)



            Dim realm = TextBox1.Text.Trim
            Dim user = TextBox2.Text.Trim
            Dim pass = TextBox3.Text.Trim

            If realm.Length < 1 Or user.Length < 1 Or pass.Length < 1 Then

                MsgBox("Please Fill the Parameters")
                Exit Sub

            End If
            logincookie = c1.login(user, pass, realm)

            If logincookie Is Nothing Then

                MsgBox("Failed to Login")
                Exit Sub
            End If
            Dim filesCount = files.Count
            DicomEngine.Startup()
            For i = 0 To filesCount - 1
                c1.log(vbNewLine & vbNewLine & vbNewLine)
                Dim Pat = files(i)

                Label10.Invoke(Sub()
                                   Label10.Text = "[ " & i & " ] " & ("Processing File : " & Pat).Replace(Label2.Text, "- - -")
                               End Sub)
                'Thread.Sleep(500)



                If Not c1.IsUploaded(Pat) Or UploadPreviouslyUploaded Then

                    c1.renameFile(Pat, -1, Pat)
                    'MsgBox(Pat)
                    c1.log("code ; 555121 loading image Pat : " & Pat)
                    Dim ext = Path.GetExtension(Pat)
                    If ext.Trim = "" Or ext.ToLower = ".dicom" Or ext.ToLower = ".dcm" Then

                        Dim DicomDataset As New DicomDataSet()

                        DicomDataset.Load(Pat, DicomDataSetLoadFlags.None)

                        'Dim mstream As MemoryStream = New MemoryStream()
                        'DicomDataset.Save(mstream, DicomDataSetSaveFlags.ExplicitVR Or DicomDataSetSaveFlags.MetaHeaderPresent)
                        'Dim dataSetBytes As Byte() = mstream.ToArray()


                        'Dim Codecs As RasterCodecs = New RasterCodecs()
                        'Codecs.Options.Jpeg.Save.QualityFactor = 15
                        'Codecs.ThrowExceptionsOnInvalidImages = True
                        'Dim pixelDataElement As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.PixelData, True)
                        'Dim Image As RasterImage = DicomDataset.GetImage(pixelDataElement, 0, 0, RasterByteOrder.Rgb, DicomGetImageFlags.AllowRangeExpansion Or DicomGetImageFlags.AutoApplyModalityLut Or DicomGetImageFlags.AutoApplyVoiLut)

                        'Dim Command As SharpenCommand = New SharpenCommand(0)

                        'Command.Run(Image)


                        'mstream = New MemoryStream()
                        'Codecs.Save(Image, mstream, RasterImageFormat.DicomJpegColor, 8)
                        'dataSetBytes = mstream.ToArray()
                        'My.Computer.FileSystem.WriteAllBytes("C:\Users\developer\Desktop\hooooo.dcm", mstream.ToArray(), False)
                        'Dim mstream1 As New MemoryStream()
                        'DicomDataset.Save(mstream1, DicomDataSetSaveFlags.None)
                        'Dim codecs As RasterCodecs = New RasterCodecs()
                        'codecs.Load(mstream1)

                        'Dim pixelDataElement As DicomElement = DicomDataset.FindFirstElement(Nothing, DicomTag.PixelData, True)
                        'Dim Image As RasterImage = codecs.Load(mstream1) ' DicomDataset.GetImage(pixelDataElement, 0, 0, RasterByteOrder.Rgb, DicomGetImageFlags.AllowRangeExpansion Or DicomGetImageFlags.AutoApplyModalityLut Or DicomGetImageFlags.AutoApplyVoiLut)

                        'Dim command As SharpenCommand = New SharpenCommand(0)
                        'command.Run(Image)


                        ''MsgBox("1")
                        'Dim dicomdataset1 As DicomDataSet = New DicomDataSet()

                        'Dim mstream As New MemoryStream()
                        'codecs.Save(Image, mstream, RasterImageFormat.DicomGray, 8)
                        '' MsgBox("2")
                        'My.Computer.FileSystem.WriteAllBytes("C:\Users\developer\Desktop\hooooo.dcm", mstream.ToArray(), False)
                        '' MsgBox("3")
                        'dicomdataset1.Save("C:\Users\developer\Desktop\hooooo.dcm", DicomDataSetSaveFlags.ImplicitVR)
                        ''MsgBox("4")


                        'Exit Sub




                        Dim s(2) As String

                        s = c1.getPatientInfoFromDicom(DicomDataset)
                        If s(0) = "" Then
                            c1.log("code ; 555122 no name found")
                        Else

                            Dim inactive As Boolean = CheckBox2.Checked

                            Dim uuid = c1.getPatientuui(realm, logincookie, s(0), s(1), inactive, Country)
                            'MsgBox(uuid)
                            If (uuid Is Nothing Or uuid = "") And inactive Then
                                uuid = c1.getPatientuui(realm, logincookie, s(0), s(1), Not inactive, Country)
                                'MsgBox("kk" & uuid)
                            End If

                            lock = True
                            If uuid IsNot Nothing And uuid <> "" Then


                                If c1.uploadImageDicom(DicomDataset, realm, logincookie, uuid) Then

                                    DicomDataset.Dispose()

                                    UploadedFiles = UploadedFiles + 1
                                    c1.renameFile(Pat, 1, Pat)
                                    'If deleteAfetrUpload Then
                                    '    File.Delete(Pat)
                                    'End If
                                Else

                                    DicomDataset.Dispose()
                                    c1.renameFile(Pat, 3, Pat)
                                    ErrorFiles = ErrorFiles + 1
                                End If
                            Else
                                c1.log("code : 555825 patient was not found in both active and inactive")
                                DicomDataset.Dispose()
                                If CheckBox3.Checked Then
                                    My.Computer.FileSystem.CopyFile(Pat, My.Computer.FileSystem.SpecialDirectories.MyDocuments & "/PNF " & DNF & "/" & Regex.Replace(s(0), "[^A-Z a-z0-9\-/]", "") & "/" & DateTime.Now.ToString("dd-MM-yyyy HH-mm-ss-fff"))
                                    My.Computer.FileSystem.WriteAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "/PNF " & DNF & "/" & Regex.Replace(s(0), "[^A-Z a-z0-9\-/]", "") & "/info.txt", s(0) & "---" & s(1), False)
                                End If

                                c1.renameFile(Pat, 4, Pat)
                                PNF = PNF + 1
                            End If
                            lock = False
                        End If
                    Else

                        c1.log("code ; 555125 wrong file format : " & Pat)
                        c1.renameFile(Pat, 2, Pat)
                        IgnoredFile = IgnoredFile + 1
                    End If
                Else
                    AlreadyUploaded = AlreadyUploaded + 1
                    c1.log("code ; 555633 uploaded before File : " & Pat)
                End If
                'MsgBox(th.ThreadState.ToString)
                ProgressBar1.Invoke(Sub()
                                        ProgressBar1.Value = Math.Round((i + 1) / filesCount * 100)

                                    End Sub)
                Label6.Invoke(Sub()
                                  Label6.Text = UploadedFiles
                              End Sub)
                Label8.Invoke(Sub()
                                  Label8.Text = ErrorFiles
                              End Sub)
                Label11.Invoke(Sub()
                                   Label11.Text = IgnoredFile
                               End Sub)
                Label13.Invoke(Sub()
                                   Label13.Text = PNF
                               End Sub)
                Label15.Invoke(Sub()
                                   Label15.Text = AlreadyUploaded
                               End Sub)


            Next
            files = Nothing
            DicomEngine.Shutdown()
        Catch ex As Exception
            c1.log("error code ; 666123  " & ex.Message)
        End Try
    End Sub

    Private Sub loadFiles()

        If th.ThreadState = ThreadState.Running Then
            files = Nothing
            Exit Sub
        End If
        If th1.ThreadState = ThreadState.Running Then
            files = Nothing
            Exit Sub
        End If

        If Not Directory.Exists(Label2.Text) Then
            MsgBox("Please select a valid Directory")
            files = Nothing
            Exit Sub
        End If

        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            Label2.Text = FolderBrowserDialog1.SelectedPath
        End If

        If th1.ThreadState = ThreadState.Stopped Or th1.ThreadState = ThreadState.Aborted Then
            th1 = New Thread(AddressOf bw1_DoWork)
        End If
        If th1Load.ThreadState = ThreadState.Stopped Or th1Load.ThreadState = ThreadState.Aborted Then
            th1Load = New Thread(AddressOf bw1_DoWorkLoad)
        End If

        th1.Start()
        th1Load.Start()

    End Sub


    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        loadFiles()
    End Sub


    Private Sub bw1_DoWork()
        files = Directory.GetFiles(Label2.Text, "*", SearchOption.AllDirectories)
        Label4.Invoke(Sub()
                          Label4.Text = files.Count
                      End Sub)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then

            UploadPreviouslyUploaded = True
        Else
            UploadPreviouslyUploaded = False
        End If
    End Sub

    Private Sub bw1_DoWorkLoad()
        Dim i = 1

        While th1.ThreadState = ThreadState.Running
            'MsgBox("pr")
            Dim tex = "Loading Files"
            For j = 0 To i
                tex = tex & " ."
            Next
            Label9.Invoke(Sub()
                              Label9.Text = tex
                          End Sub)
            Thread.Sleep(500)
            i = (i + 1) Mod 5
        End While
        Label9.Invoke(Sub()
                          Label9.Text = ""
                      End Sub)
    End Sub

    Private Sub bw_DoWorkLoad()
        Dim i = 1

        While th.ThreadState <> ThreadState.Stopped And th.ThreadState <> ThreadState.Aborted
            'MsgBox("pr")
            Dim tex = "Processing"
            For j = 0 To i
                tex = tex & " ."
            Next
            Label9.Invoke(Sub()
                              Label9.Text = tex
                          End Sub)
            Thread.Sleep(500)
            i = (i + 1) Mod 5
        End While
        Label9.Invoke(Sub()
                          Label9.Text = th.ThreadState.ToString
                      End Sub)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        While lock

        End While
        If th.ThreadState = ThreadState.Running Then
            th.Abort()
        End If
        If th1.ThreadState = ThreadState.Running Then
            th1.Abort()
        End If
        If thLoad.ThreadState = ThreadState.Running Then
            thLoad.Abort()
        End If
        If th.ThreadState = ThreadState.Running Then
            th1Load.Abort()
        End If
        files = Nothing
    End Sub

    Private Sub frmSimple_Disposed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim result = MsgBox("Are you sure you want to Exit ?", vbYesNo)

        If result = Windows.Forms.DialogResult.Yes Then

            Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Control Panel\International", True).SetValue("sShortDate", dateFormat)
            If th.ThreadState = ThreadState.Running Then
                th.Abort()
            End If
            If th1.ThreadState = ThreadState.Running Then
                th1.Abort()
            End If
            If thLoad.ThreadState = ThreadState.Running Then
                thLoad.Abort()
            End If
            If th.ThreadState = ThreadState.Running Then
                th1Load.Abort()
            End If

        Else
            e.Cancel = True
        End If

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.DoubleClick
        If Directory.Exists(Label2.Text) And Label2.Text <> "." Then
            c1.loadFileExplorer(Label2.Text, "*.Oryx-Uploaded")
        End If
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.DoubleClick
        If Directory.Exists(Label2.Text) And Label2.Text <> "." Then
            c1.loadFileExplorer(Label2.Text, "*.Oryx-Error")
        End If
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.DoubleClick
        If Directory.Exists(Label2.Text) And Label2.Text <> "." Then
            c1.loadFileExplorer(Label2.Text, "*.Oryx-Ignored")
        End If
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.DoubleClick
        If Directory.Exists(Label2.Text) And Label2.Text <> "." Then
            c1.loadFileExplorer(Label2.Text, "*.Oryx-PNF")
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Country = "US"
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            Country = "CA"
        End If
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim c As CookieContainer = c1.login("salehrania@gmail.com", "Password1!", "training-us.evidentiae.com")
        c1.uploadImageJpeg("C:\Users\developer\Desktop\3.jpeg", "training-us.evidentiae.com", c, "2204", "3.jpeg", "kha")
    End Sub

End Class
