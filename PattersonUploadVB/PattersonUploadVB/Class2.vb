Imports Leadtools
Imports Leadtools.Codecs
Imports Leadtools.ImageProcessing
Imports Leadtools.ImageProcessing.Color
Imports Leadtools.Drawing
Imports Leadtools.Svg
Imports System.IO

Public Class Class2


    Public Sub RasterCodecsExample()

        Dim codecs As RasterCodecs = New RasterCodecs()

        Dim srcFileName As String = Path.Combine(LEAD_VARS.ImagesDir, "Image1.cmp")
        Dim dstFileName As String = Path.Combine(LEAD_VARS.ImagesDir, "Image1_test.jpg")

        Dim image As RasterImage = codecs.Load(srcFileName)
        codecs.Save(image, dstFileName, RasterImageFormat.Jpeg, 0)

        ' Clean up 
        image.Dispose()
        codecs.Dispose()
    End Sub

    Public NotInheritable Class LEAD_VARS
        Public Const ImagesDir As String = "C:\Users\Public\Documents\LEADTOOLS Images"
    End Class
End Class
