Imports System.Drawing
Imports System.Configuration

Public Class AImgAdd
    Inherits System.Web.UI.Page
    Protected strFileName As String = "" : Protected strSavePath As String = ""
    Protected str_mappath As String = ""
    Dim field_name, type As String
    Protected WithEvents valPicFile As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents UploadFile1 As System.Web.UI.HtmlControls.HtmlInputFile

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '*** this if working only with chrome, in case of IE use the next if shuld be deleted ***
        'If InStr(Request.ServerVariables("HTTP_REFERER"), Request.ServerVariables("HTTP_HOST")) Then
        'Else
        '    Response.Redirect("../../../")
        'End If
        'If IsNothing(Request.Cookies(ConfigurationSettings.AppSettings("AdminCookieName"))) Then
        If IsNothing(Request.Cookies("bizpegasus")("UserId")) Then
            Response.Redirect("../../../")
        End If
        'If Not IsNumeric(Request.Cookies(ConfigurationSettings.AppSettings("AdminCookieName"))("workerid")) Then
        If Not IsNumeric(Request.Cookies("bizpegasus")("UserId")) Then
            Response.Redirect("../../../")
        End If

        field_name = Trim(Request.QueryString("field_name"))
        type = Trim(Request.QueryString("type"))
        If Len(type) = 0 Then
            type = "image"
        End If

        str_mappath = Application("VirDir") & "/download/pictures"

        Server.ScriptTimeout = 3000

        Page.Validate()

        If Request.Files.Count > 0 And Page.IsValid Then

            If Request.Files("UploadFile1").ContentLength > 0 Then


                Dim Ext As String = IO.Path.GetExtension(Request.Files("UploadFile1").FileName)
                Dim PicType As String = LCase(Request.Files("UploadFile1").ContentType.ToString())

                Dim FileNameWithoutExt As String = IO.Path.GetFileNameWithoutExtension(Request.Files("UploadFile1").FileName)
                Dim fs As IO.File
                Dim i As Integer = 1
                strFileName = FileNameWithoutExt
                Do While IO.File.Exists(Server.MapPath(str_mappath & "/" & strFileName & Ext))
                    strFileName = FileNameWithoutExt & "_" & CStr(i)
                    i = i + 1
                Loop
                strFileName = strFileName & Ext

                If (PicType.StartsWith("image/")) Then
                    'Check , create new directory
                    strSavePath = Server.MapPath(str_mappath & "/" & strFileName)

                    'Read in the uploaded file into a Stream
                    Dim objStream As System.IO.Stream = Request.Files("UploadFile1").InputStream
                    Dim objImage As Image = Image.FromStream(objStream)
                    SavePictureThumbnail(1800, 652, strSavePath, objImage)
                    objImage.Dispose()
                    objStream.Close()
                End If
                fs = Nothing

                'Dim scrp As String = "<SCRIPT LANGUAGE=javascript>" & vbCrLf & _
                '   "<!--" & vbCrLf & _
                '   "   window.opener.document.forms[0].elements['" & field_name & "'].value = '" & "http://" & Request.ServerVariables("Server_name") & str_mappath & "/" & strFileName & "';" & vbCrLf & _
                '   "   self.close(); " & vbCrLf & _
                '   "//-->" & vbCrLf & _
                '   "</SCRIPT>"
                Dim scrp As String = "<SCRIPT LANGUAGE=javascript>" & vbCrLf & _
                   "<!--" & vbCrLf & _
                   "   window.opener.document.getElementById('" & field_name & "').value = '" & "https://" & Request.ServerVariables("Server_name") & str_mappath & "/" & strFileName & "';" & vbCrLf & _
                   "   self.close(); " & vbCrLf & _
                   "//-->" & vbCrLf & _
                   "</SCRIPT>"
                '   ClientScript.RegisterStartupScript(Me.GetType(), "onload", scrp)
                Page.RegisterStartupScript("onload", scrp)

                'RegisterStartupScript("onload", scrp)
            End If
        End If
    End Sub

    Public Sub SavePictureThumbnail(ByVal WantedHeight As Integer, ByVal WantedWidth As Integer, _
ByVal strPicPath As String, ByVal objImage As Image)

        'assign photo height
        Dim OriginalPhotoHeight, NewPhotoHeight As Integer
        OriginalPhotoHeight = objImage.Height

        'assign photo width
        Dim OriginalPhotoWidth, NewPhotoWidth As Integer
        OriginalPhotoWidth = objImage.Width

        If OriginalPhotoHeight > OriginalPhotoWidth Then
            If OriginalPhotoHeight > WantedHeight Then
                NewPhotoHeight = WantedHeight
            Else
                NewPhotoHeight = OriginalPhotoHeight
            End If

            NewPhotoWidth = CInt((OriginalPhotoWidth * NewPhotoHeight) / OriginalPhotoHeight)
        Else
            If OriginalPhotoWidth > WantedWidth Then
                NewPhotoWidth = WantedWidth
            Else
                NewPhotoWidth = OriginalPhotoWidth
            End If

            NewPhotoHeight = CInt((OriginalPhotoHeight * NewPhotoWidth) / OriginalPhotoWidth)
        End If

        If ((OriginalPhotoHeight <= NewPhotoHeight) And (OriginalPhotoWidth <= NewPhotoWidth)) _
        And (objImage.RawFormat.Equals(Drawing.Imaging.ImageFormat.Gif) Or objImage.RawFormat.Equals(Drawing.Imaging.ImageFormat.Jpeg)) Then
            objImage.Save(strPicPath)
        Else
            SavePicture(WantedHeight, WantedWidth, strPicPath, objImage)
        End If

        objImage.Dispose()

    End Sub

    Public Sub SavePicture(ByVal WantedHeight As Integer, ByVal WantedWidth As Integer, ByVal strPicPath As String, ByVal objImage As Image)

        'assign photo height
        Dim OriginalPhotoHeight, NewPhotoHeight As Integer
        OriginalPhotoHeight = objImage.Height

        'assign photo width
        Dim OriginalPhotoWidth, NewPhotoWidth As Integer
        OriginalPhotoWidth = objImage.Width

        If OriginalPhotoHeight > OriginalPhotoWidth Then
            If OriginalPhotoHeight > WantedHeight Then
                NewPhotoHeight = WantedHeight
            Else
                NewPhotoHeight = OriginalPhotoHeight
            End If

            NewPhotoWidth = CInt((OriginalPhotoWidth * NewPhotoHeight) / OriginalPhotoHeight)
        Else
            If OriginalPhotoWidth > WantedWidth Then
                NewPhotoWidth = WantedWidth
            Else
                NewPhotoWidth = OriginalPhotoWidth
            End If

            NewPhotoHeight = CInt((OriginalPhotoHeight * NewPhotoWidth) / OriginalPhotoWidth)
        End If

        Dim newPic As Drawing.Bitmap = New Drawing.Bitmap(NewPhotoWidth, NewPhotoHeight, Imaging.PixelFormat.Format24bppRgb)
        Dim g As Drawing.Graphics
        g = Graphics.FromImage(newPic)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
        g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
        g.FillRectangle(Brushes.White, 0, 0, NewPhotoWidth, NewPhotoHeight)
        g.DrawImage(objImage, 0, 0, NewPhotoWidth, NewPhotoHeight)

        'Set the image quality 
        Dim objImageCodecInfo As Imaging.ImageCodecInfo
        Dim objEncoder0 As Imaging.Encoder
        Dim objEncoderParameter0 As Imaging.EncoderParameter
        Dim objEncoderParameters As Imaging.EncoderParameters

        'Get an ImageCodecInfo object that represents the JPEG codec. 
        objImageCodecInfo = GetEncoderInfo("image/jpeg")

        'Create an EncoderParameters object. 
        'An EncoderParameters object has an array of EncoderParameter 
        'objects. In this case, there is only one 
        ' EncoderParameter object in the array. 
        objEncoderParameters = New Imaging.EncoderParameters(1)

        'Create an Encoder object based on the GUID 
        'for the Quality parameter category. 
        objEncoder0 = Imaging.Encoder.Quality

        'Save the bitmap as a JPEG file with quality level 85. 
        objEncoderParameter0 = New Imaging.EncoderParameter(objEncoder0, 85L)
        objEncoderParameters.Param(0) = objEncoderParameter0

        newPic.SetResolution(150, 150)
        newPic.Save(strPicPath, objImageCodecInfo, objEncoderParameters)

        newPic.Dispose()
        g.Dispose()
    End Sub

    Function GetEncoderInfo(ByRef mimeType As String) As Imaging.ImageCodecInfo

        For Each ice As Imaging.ImageCodecInfo In Imaging.ImageCodecInfo.GetImageEncoders()
            If ice.MimeType.Equals(mimeType) Then
                Return ice
            End If
        Next
        Throw New Exception(mimeType + " mime type not found in ImageCodecInfo")
    End Function

    Private Sub InitializeComponent()

    End Sub
End Class
