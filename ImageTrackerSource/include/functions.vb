Imports System.Web.HttpContext
Imports System.Text
Imports System.IO
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions


Namespace include

    Friend Class StringWriterWithEncoding
        Inherits StringWriter

        Private m_encoding As Encoding

        Public Sub New(ByVal sb As StringBuilder, ByVal encoding As Encoding)
            MyBase.New(sb)
            m_encoding = encoding
        End Sub

        Public Overrides ReadOnly Property Encoding() As Encoding
            Get
                Return m_encoding
            End Get
        End Property

    End Class


    Public Class funcs
        Public Function altFix(ByVal word As Object) As String
            Dim temp As String
            If Not IsDBNull(word) Then
                temp = CStr(word)
            Else
                Return ""
                Exit Function
            End If

            temp = Replace(temp, "'", "\x27")    ' JScript encode apostrophes 
            temp = Replace(temp, """", "\x22")   ' JScript encode double-quotes 
            temp = HttpUtility.HtmlEncode(temp)  ' encode chars special to HTML 

            Return temp

        End Function
        Public Function checkDevice() As Boolean ''//detecting mobile devices

            '    If Request.QueryString.Count = 0 Or Request.QueryString("checkDevice") = "true" Then
            Dim user_agent, mobile_ua As String
            Dim mobile_browser As Integer = 0
            Dim Size As Integer = 0
            Dim match As Boolean
            user_agent = LCase(Current.Request.ServerVariables("HTTP_USER_AGENT"))
            ' Current.Response.Write("user_agent=" & user_agent)


            mobile_browser = 0

            Dim RegexVar As Regex

            If Trim(user_agent & "") <> "" Then
                match = RegexVar.IsMatch(user_agent, "(android)")
                If match Then
                    mobile_browser = 0
                    Return False

                End If
                match = RegexVar.IsMatch(user_agent, "(up.browser|up.link|mmp|symbian|smartphone|midp|wap|phone|windows ce|pda|mobile|mini|palm)")

                If match Then mobile_browser = mobile_browser + 1

                If InStr(Current.Request.ServerVariables("HTTP_ACCEPT"), "application/vnd.wap.xhtml+xml") Or Not IsNothing(Current.Request.ServerVariables("HTTP_X_PROFILE")) Or Not IsNothing(Current.Request.ServerVariables("HTTP_PROFILE")) Then
                    mobile_browser = mobile_browser + 1
                End If

                Dim mobile_agents() As String = {"w3c ", "acs-", "alav", "alca", "amoi", "audi", "avan", "benq", "bird", "blac", "blaz", "brew", "cell", "cldc", "cmd-", "dang", "doco", "eric", "hipt", "inno", "ipaq", "java", "jigs", "kddi", "keji", "leno", "lg-c", "lg-d", "lg-g", "lge-", "maui", "maxo", "midp", "mits", "mmef", "mobi", "mot-", "moto", "mwbp", "nec-", "newt", "noki", "oper", "palm", "pana", "pant", "phil", "play", "port", "prox", "qwap", "sage", "sams", "sany", "sch-", "sec-", "send", "seri", "sgh-", "shar", "sie-", "siem", "smal", "smar", "sony", "sph-", "symb", "t-mo", "teli", "tim-", "tosh", "tsm-", "upg1", "upsi", "vk-v", "voda", "wap-", "wapa", "wapi", "wapp", "wapr", "webc", "winw", "winw", "xda", "xda-"}
                Size = UBound(mobile_agents)
                mobile_ua = LCase(Left(user_agent, 4))

                ''Response.Write(mobile_browser & "<br><br>")

                Dim i As Integer
                For i = 0 To Size
                    If mobile_agents(i) = mobile_ua Then
                        mobile_browser = mobile_browser + 1
                        Exit For
                    End If
                Next

                If mobile_browser > 0 Then
                    Return True

                End If
            End If


            ''Response.Write(mobile_browser & "<br><br>")
            ''Response.End()
            '   End If
        End Function

        Public Shared Function sFix(ByVal myString As Object) As String
            Dim quote, newQuote, temp As String
            If Not IsDBNull(myString) Then
                temp = CStr(myString)
            Else
                Return ""
                Exit Function
            End If

            If temp <> "" Then
                quote = "'"
                newQuote = "''"
                If InStr(myString, "'") Then
                    temp = Replace(temp, quote, newQuote)
                End If
            End If
            Return Trim(temp)

        End Function
        Function GetGenderName(ByVal Gender, ByVal LangID)
            If Not IsDBNull(Gender) Then
                Select Case Gender
                    Case "1" : GetGenderName = "בן"
                    Case "2" : GetGenderName = "בת"
                End Select
            Else
                GetGenderName = ""
            End If
        End Function

        Function getEventsProgramId(ByVal friendlyURL)
            If friendlyURL <> "" Then
                Dim strsql As String
                Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

                strsql = "SELECT Program_Id FROM Events " & _
                "WHERE (Events.rewrite_friendlyUrl = N'" & sFix(friendlyURL) & "' or Events.Event_Title = N'" & sFix(friendlyURL) & "')"
                'Current.Response.Write("<BR>" & strsql)
                ' Current.Response.End()
                con.Open()
                Dim cmdSelect = New System.Data.SqlClient.SqlCommand(strsql, con)
                Dim pageId = cmdSelect.ExecuteScalar
                con.Close()
                Return pageId
            End If
        End Function

        Function getEventsEventID(ByVal friendlyURL)
            If friendlyURL <> "" Then
                Dim strsql As String
                Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

                strsql = "SELECT Event_Id FROM Events " & _
                "WHERE (Events.rewrite_friendlyUrl = N'" & sFix(friendlyURL) & "' or Events.Event_Title = N'" & sFix(friendlyURL) & "')"
                'Current.Response.Write("<BR>" & strsql)
                ' Current.Response.End()
                con.Open()
                Dim cmdSelect = New System.Data.SqlClient.SqlCommand(strsql, con)
                Dim pageId = cmdSelect.ExecuteScalar
                con.Close()
                Return pageId
            End If
        End Function
        Function getMediaMediaID(ByVal friendlyURL)
            If friendlyURL <> "" Then
                Dim strsql As String
                Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

                strsql = "SELECT Id FROM Pg " & _
                "WHERE (Pg.rewrite_friendlyUrl = N'" & sFix(friendlyURL) & "' or Pg.hdr = N'" & sFix(friendlyURL) & "')"
                ' Current.Response.Write("<BR>" & strsql)
                'Current.Response.End()
                con.Open()
                Dim cmdSelect = New System.Data.SqlClient.SqlCommand(strsql, con)
                Dim pageId = cmdSelect.ExecuteScalar
                con.Close()
                Return pageId
            End If


        End Function
        Function getMediaProgramId(ByVal friendlyURL)
            If friendlyURL <> "" Then
                Dim strsql As String
                Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

                strsql = "SELECT ProgramId FROM Pg " & _
                "WHERE (Pg.rewrite_friendlyUrl = N'" & sFix(friendlyURL) & "' or Pg.hdr = N'" & sFix(friendlyURL) & "')"
                'Current.Response.Write("<BR>" & strsql)
                ' Current.Response.End()
                con.Open()
                Dim cmdSelect = New System.Data.SqlClient.SqlCommand(strsql, con)
                Dim pageId = cmdSelect.ExecuteScalar
                con.Close()
                Return pageId
            End If



        End Function

        ''single quote
        Public Shared Function sqFix(ByVal myString As Object) As String
            Dim quote, newQuote, temp As String
            If Not IsDBNull(myString) Then
                temp = CStr(myString)
            Else
                Return ""
                Exit Function
            End If

            If temp <> "" Then
                quote = "'"
                newQuote = "''"
                If InStr(myString, "'") Then
                    temp = Replace(temp, quote, newQuote)
                End If
            End If
            Return Trim(temp)

        End Function

        Public Shared Function killChars(ByVal strWords As String) As String

            Dim badChars() As String = {"select", "drop", "--", "insert", "delete", "xp_", "exec", "cmdshell", "drop", "database"}
            Dim newChars As String
            Dim i As Integer
            newChars = strWords

            For i = 0 To UBound(badChars)
                newChars = Replace(newChars, badChars(i), "")
            Next
            Return newChars

        End Function
        Public Shared Function killCharsFriendlyUrl(ByVal strWords As String) As String

            Dim badChars() As String = {"""", ":", "&", "+", "!", "?", ",", ".", "<", ">", "(", ")"}
            Dim newChars As String
            Dim i As Integer
            newChars = strWords

            For i = 0 To UBound(badChars)
                newChars = Replace(newChars, badChars(i), "")
            Next
            Return newChars

        End Function
        Public Shared Function vFix(ByVal word As Object) As String
            Dim temp As String
            If Not IsDBNull(word) Then
                temp = CStr(word)
            Else
                Return ""
                Exit Function
            End If
            temp = Replace(temp, "<", "&lt;")
            temp = Replace(temp, ">", "&gt;")
            temp = Replace(temp, Chr(34), "&quot;")

            Return Trim(temp)

        End Function

        Function nFix(ByVal word As Object) As Object
            If Not IsDBNull(word) And IsNumeric(word) Then
                Return word
            Else
                Return DBNull.Value
            End If
        End Function
        Function Truncate(ByVal o As Object, ByVal L As Integer) As String
            If Len(o.ToString()) > L Then
                Return Left(o.ToString(), L) & "&nbsp; ...&nbsp;"
            Else
                Return o.ToString()

            End If

        End Function


        Public Function HtmlBreak(ByVal strWords As String) As String
            Return Replace(strWords, vbCrLf, "<br>")
        End Function

        Function NmbrFormat(ByVal nbr As Object, ByVal strTmp As String) As Object
            If Not IsDBNull(nbr) And IsNumeric(nbr) Then
                If Trim(nbr) = "0" Then
                    Return ""
                ElseIf IsNumeric(nbr) Then
                    Return FormatNumber(nbr, 0, TriState.True, TriState.True, TriState.True).ToString() & strTmp
                End If
            Else
                Return nbr
            End If
        End Function
        Public Function hebrewUrlEncode(ByVal str As String) As String

            hebrewUrlEncode = HttpUtility.UrlEncode(str, System.Text.Encoding.GetEncoding("windows-1255"))

        End Function

        Public Function hebrewUrlDecode(ByVal str As String) As String

            hebrewUrlDecode = HttpUtility.UrlDecode(str, System.Text.Encoding.GetEncoding("windows-1255"))

        End Function
        Function Decode(ByVal sIn As String)
            Dim x, y, abfrom, abto
            Decode = "" : abfrom = ""

            For x = 0 To 25 : abfrom = abfrom & Chr(65 + x) : Next
            For x = 0 To 25 : abfrom = abfrom & Chr(97 + x) : Next
            For x = 0 To 9 : abfrom = abfrom & CStr(x) : Next

            abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
            For x = 1 To Len(sIn) : y = InStr(abto, Mid(sIn, x, 1))
                If y = 0 Then
                    Decode = Decode & Mid(sIn, x, 1)
                Else
                    Decode = Decode & Mid(abfrom, y, 1)
                End If
            Next
        End Function

        Function Encode(ByVal sIn)
            Dim x, y, abfrom, abto
            Encode = "" : abfrom = ""

            For x = 0 To 25 : abfrom = abfrom & Chr(65 + x) : Next
            For x = 0 To 25 : abfrom = abfrom & Chr(97 + x) : Next
            For x = 0 To 9 : abfrom = abfrom & CStr(x) : Next

            abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
            For x = 1 To Len(sIn) : y = InStr(abfrom, Mid(sIn, x, 1))
                If y = 0 Then
                    Encode = Encode & Mid(sIn, x, 1)
                Else
                    Encode = Encode & Mid(abto, y, 1)
                End If
            Next
        End Function

        Function breaks(ByVal word)
            If Not IsDBNull(word) Then
                word = Replace(word, Chr(13) & Chr(10), "<br>")
                breaks = word
            Else
                breaks = ""
            End If
        End Function

        Function cutLength(ByVal word As String, ByVal Length As Integer) As String
            Dim temp As String
            If Len(word) <= Length Then
                temp = word
            Else
                temp = Left(word, Length) & "..."
            End If

            Return temp

        End Function

        Function GetEncoderInfo(ByRef mimeType As String) As System.Drawing.Imaging.ImageCodecInfo

            For Each ice As System.Drawing.Imaging.ImageCodecInfo In System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                If ice.MimeType.Equals(mimeType) Then
                    Return ice
                End If
            Next
            Throw New Exception(mimeType + " mime type not found in ImageCodecInfo")
        End Function
        Public Function ShowPicture(ByVal pic_url) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    Return "<img src='" & Current.Application("VirDir") & "/download/users/" & pic_url & "' border=0 hspace=0 vspace=0>"
                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function

        Public Sub SavePictureBanner(ByVal WantedHeight As Integer, ByVal WantedWidth As Integer, _
                ByVal strPicPath As String, ByVal objImage As System.Drawing.Image)
            Dim NewPhotoWidth, NewPhotoHeight As Integer
            NewPhotoWidth = WantedWidth
            NewPhotoHeight = WantedHeight
            Dim newPic As Drawing.Bitmap = New Drawing.Bitmap(NewPhotoWidth, NewPhotoHeight, Imaging.PixelFormat.Format64bppPArgb)
            Dim g As Drawing.Graphics
            g = g.FromImage(newPic)
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
            g.FillRectangle(Brushes.White, 0, 0, NewPhotoWidth, NewPhotoHeight)
            g.DrawImage(objImage, 0, 0, NewPhotoWidth, NewPhotoHeight)
            g.Dispose()

            'Set the image quality 
            Dim objImageCodecInfo As System.Drawing.Imaging.ImageCodecInfo
            Dim objEncoder0 As System.Drawing.Imaging.Encoder
            Dim objEncoderParameter0 As System.Drawing.Imaging.EncoderParameter
            Dim objEncoderParameters As System.Drawing.Imaging.EncoderParameters

            Dim strMime As String = "image/jpeg" : Dim ext = ".jpg"

            objImageCodecInfo = GetEncoderInfo(strMime)

            'Create an EncoderParameters object. 
            'An EncoderParameters object has an array of EncoderParameter 
            'objects. In this case, there is only one 
            ' EncoderParameter object in the array. 
            objEncoderParameters = New System.Drawing.Imaging.EncoderParameters(1)

            'Create an Encoder object based on the GUID 
            'for the Quality parameter category. 
            objEncoder0 = System.Drawing.Imaging.Encoder.Quality

            'Save the bitmap as a JPEG file with quality level 85. 
            objEncoderParameter0 = New System.Drawing.Imaging.EncoderParameter(objEncoder0, 85L)
            objEncoderParameters.Param(0) = objEncoderParameter0

            newPic.SetResolution(150, 150)

            newPic.Save(strPicPath, objImageCodecInfo, objEncoderParameters)

            newPic.Dispose()


        End Sub


        Public Sub SavePicture(ByVal WantedHeight As Integer, ByVal WantedWidth As Integer, _
        ByVal strPicPath As String, ByVal objImage As System.Drawing.Image)

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

                If NewPhotoHeight > WantedHeight Then
                    NewPhotoHeight = WantedHeight
                    NewPhotoWidth = CInt((OriginalPhotoWidth * NewPhotoHeight) / OriginalPhotoHeight)
                End If
            End If


            Dim newPic As Drawing.Bitmap = New Drawing.Bitmap(NewPhotoWidth, NewPhotoHeight, Imaging.PixelFormat.Format64bppPArgb)
            Dim g As Drawing.Graphics
            g = g.FromImage(newPic)
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
            g.FillRectangle(Brushes.White, 0, 0, NewPhotoWidth, NewPhotoHeight)
            g.DrawImage(objImage, 0, 0, NewPhotoWidth, NewPhotoHeight)
            g.Dispose()

            'Set the image quality 
            Dim objImageCodecInfo As System.Drawing.Imaging.ImageCodecInfo
            Dim objEncoder0 As System.Drawing.Imaging.Encoder
            Dim objEncoderParameter0 As System.Drawing.Imaging.EncoderParameter
            Dim objEncoderParameters As System.Drawing.Imaging.EncoderParameters

            Dim strMime As String = "image/jpeg" : Dim ext = ".jpg"

            objImageCodecInfo = GetEncoderInfo(strMime)

            'Create an EncoderParameters object. 
            'An EncoderParameters object has an array of EncoderParameter 
            'objects. In this case, there is only one 
            ' EncoderParameter object in the array. 
            objEncoderParameters = New System.Drawing.Imaging.EncoderParameters(1)

            'Create an Encoder object based on the GUID 
            'for the Quality parameter category. 
            objEncoder0 = System.Drawing.Imaging.Encoder.Quality

            'Save the bitmap as a JPEG file with quality level 85. 
            objEncoderParameter0 = New System.Drawing.Imaging.EncoderParameter(objEncoder0, 85L)
            objEncoderParameters.Param(0) = objEncoderParameter0

            newPic.SetResolution(150, 150)

            newPic.Save(strPicPath, objImageCodecInfo, objEncoderParameters)

            newPic.Dispose()

        End Sub

        Public Sub SavePictureThumbnail(ByVal WantedHeight As Integer, ByVal WantedWidth As Integer, _
ByVal strPicPath As String, ByVal objImage As System.Drawing.Image)

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
            And objImage.RawFormat.Equals(Drawing.Imaging.ImageFormat.Gif) Then
                objImage.Save(strPicPath)
            Else
                SavePicture(WantedHeight, WantedWidth, strPicPath, objImage)
            End If

            objImage.Dispose()

        End Sub
        Public Function ShowMemberFoto(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/members/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function
        Public Function ShowHighLightPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/highlight/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function
        Public Function ShowHighLightRightPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/HighLightRight/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function


        Public Function ShowSectionPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/sections/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function
        Public Function ShowEventsPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function
        Public Function ShowTagPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/tags/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function

        Public Function ShowBlogPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/blogs/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function

        Public Function ShowForumPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/forums/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function

        Public Function ShowNewsPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/news/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function
        Public Function ShowPeoplesPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/peoples/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function


        Public Function ShowProgramsPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/programs/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=0 vspace=0 >"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function
        Public Function ShowMediaPicture(ByVal pic_url, ByVal width, ByVal height) As String
            If Not IsDBNull(pic_url) Then
                If Len(pic_url) > 0 Then
                    ' Return "<img src='include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode("../download/events/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 hspace=2 >"
                    Return "<img src='" & Current.Application("VirDir") & "/include/Thumbnail.aspx?filename=" & Current.Server.HtmlEncode(Current.Application("VirDir") & "/download/highlight/" & pic_url) & "&width=" & width & "&height=" & height & "' border=0 style=margin:3px>"

                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function



        Public Shared Function cFix(ByVal fromWord) As String
            Dim newWord As String
            If Trim(fromWord) <> "" Then
                newWord = LCase(fromWord)
                newWord = Replace(newWord, "select", "")
                newWord = Replace(newWord, "exec", "")
                newWord = Replace(newWord, "drop", "")
                newWord = Replace(newWord, ";", "")
                newWord = Replace(newWord, "--", "")
                newWord = Replace(newWord, "insert", "")
                newWord = Replace(newWord, "delete", "")
                newWord = Replace(newWord, "xp_", "")
                newWord = Replace(newWord, "#", "")
                newWord = Replace(newWord, "%", "")
                newWord = Replace(newWord, "&", "")
                newWord = Replace(newWord, "(", "")
                newWord = Replace(newWord, ")", "")
                'newWord = Replace(newWord, "/", "")
                newWord = Replace(newWord, "\", "")
                newWord = Replace(newWord, ":", "")
                newWord = Replace(newWord, ";", "")
                newWord = Replace(newWord, "<", "")
                newWord = Replace(newWord, ">", "")
                newWord = Replace(newWord, "=", "")
                newWord = Replace(newWord, "[", "")
                newWord = Replace(newWord, "]", "")
                newWord = Replace(newWord, "?", "")
                newWord = Replace(newWord, "`", "")
                newWord = Replace(newWord, "|", "")
                'newWord = Replace(newWord, "'", "'+char(39)+'")
                newWord = Replace(newWord, "'", "''")
                'newWord = Replace(newWord, "'", "")
                cFix = newWord
            Else
                cFix = fromWord
            End If
        End Function
        Public Shared Function searchFix(ByVal fromWord) As String
            Dim newWord As String
            If Trim(fromWord) <> "" Then
                newWord = LCase(fromWord)
                newWord = Replace(newWord, "select", "")
                newWord = Replace(newWord, "javascript", "")
                newWord = Replace(newWord, "document.location", "")
                newWord = Replace(newWord, "script", "")
                newWord = Replace(newWord, "*", "")
                newWord = Replace(newWord, "exec", "")
                newWord = Replace(newWord, "drop", "")
                newWord = Replace(newWord, "declare", "")
                newWord = Replace(newWord, ";", "")
                newWord = Replace(newWord, "--", "")
                newWord = Replace(newWord, "insert", "")
                newWord = Replace(newWord, "delete", "")
                newWord = Replace(newWord, "xp_", "")
                newWord = Replace(newWord, "#", "")
                newWord = Replace(newWord, "%", "")
                newWord = Replace(newWord, "&", "")
                ' newWord = Replace(newWord, "(", "")
                ' newWord = Replace(newWord, ")", "")
                'newWord = Replace(newWord, "/", "")
                newWord = Replace(newWord, "\", "")
                newWord = Replace(newWord, ":", "")
                newWord = Replace(newWord, ";", "")
                newWord = Replace(newWord, "<", "")
                newWord = Replace(newWord, ">", "")
                newWord = Replace(newWord, "=", "")
                newWord = Replace(newWord, "[", "")
                newWord = Replace(newWord, "]", "")
                ' newWord = Replace(newWord, "?", "")
                newWord = Replace(newWord, "`", "")
                newWord = Replace(newWord, "|", "")
                'newWord = Replace(newWord, "'", "'+char(39)+'")
                ' newWord = Replace(newWord, "'", "''")
                'newWord = Replace(newWord, "'", "")
                searchFix = newWord
            Else
                searchFix = fromWord
            End If
        End Function
        Function ShowStar(ByVal grade1, ByVal grade2, ByVal grade3, ByVal grade4, ByVal grade5, ByVal allCount, ByVal Id) As String
            Dim d_grade As Double
            Dim output, img_name As String

            If IsDBNull(grade1) Then
                grade1 = 0
            Else
                d_grade = CDbl(grade1)
            End If
            If IsDBNull(grade2) Then
                grade2 = 0
            Else
                d_grade = d_grade + CDbl(grade2) * 2
            End If
            If IsDBNull(grade3) Then
                grade3 = 0
            Else
                d_grade = d_grade + CDbl(grade3) * 3
            End If
            If IsDBNull(grade4) Then
                grade4 = 0
            Else
                d_grade = d_grade + CDbl(grade4) * 4
            End If
            If IsDBNull(grade5) Then
                grade5 = 0
            Else
                d_grade = d_grade + CDbl(grade5) * 5
            End If


            If IsNumeric(d_grade) And allCount > 0 Then
                d_grade = Math.Round(d_grade / allCount, 1)
            Else
                d_grade = 0
            End If
            'Current.Response.Write("m=" & d_grade)

            output = ""

            Dim count As Integer = Fix(d_grade)
            Dim j, g As Integer
            Dim half As Integer = 0
            If d_grade Mod count > 0 Then
                half = 1
            End If

            For i As Integer = 0 To count - 1
                j = j + 1
                output = output & "<a title='דרג/י את השידור' href=" & Current.Application("VirDir") & "/programs/scale.aspx?MediaId=" & Id & "&Star=" & i & " onmouseover=""showStars('" & Id & "','" & i & "')""   onMouseOut=""RefreshStars('" & Id & "','" & count & "','" & half & "')"" ><img src='" & Current.Application("VirDir") & "/images/star.gif' Id='Star_" & Id & "_" & i & "' name='Star_" & Id & "_" & i & "' border=0 width=15 height=16 style='position:relative;top:2px;'></a>"
            Next
            If d_grade Mod count > 0 Then
                ' half = 1
                j = j + 1
                output = output & "<a title='דרג/י את השידור' href=" & Current.Application("VirDir") & "/programs/scale.aspx?MediaId=" & Id & "&Star=" & j - 1 & " onmouseover=""showStars('" & Id & "','" & j - 1 & "')""   onMouseOut=""RefreshStars('" & Id & "','" & count & "','" & half & "')""><img src='" & Current.Application("VirDir") & "/images/star_1_2.gif'  Id='Star_" & Id & "_" & j - 1 & "' name='Star_" & Id & "_" & j - 1 & "' border=0 width=15 height=16  style='position:relative;top:2px;'></a>"
            End If
            If j <= 4 Then
                For g = j To 4
                    output = output & "<a title='דרג/י את השידור' href=" & Current.Application("VirDir") & "/programs/scale.aspx?MediaId=" & Id & "&Star=" & g & " onmouseover=""showStars('" & Id & "','" & g & "')""   onMouseOut=""RefreshStars('" & Id & "','" & count & "','" & half & "')""><img src='" & Current.Application("VirDir") & "/images/star_dis.gif'  Id='Star_" & Id & "_" & g & "' name='Star_" & Id & "_" & g & "' border=0 width=15 height=16 style='position:relative;top:2px;'></a>"

                Next


            End If

            If Len(output) = 0 Then
                output = "<span class=new_desc>&nbsp;</span>"
            End If
            '    output = d_grade & "===" & allCount
            ShowStar = output

        End Function

        Public Function getTags(ByVal strTable As String) As String
            Dim strTmp As New Text.StringBuilder("")

            Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            Dim cmdSelect As New SqlClient.SqlCommand("get_Tags_List", con)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@Table", strTable)
            con.Open()
            Dim dt As New DataTable
            Dim dta = New SqlClient.SqlDataAdapter(cmdSelect)
            dt.BeginLoadData()
            dta.Fill(dt)
            dt.EndLoadData()
            con.Close()

            For i As Integer = 0 To dt.DefaultView.Count - 1
                Dim strCount As New Text.StringBuilder(Trim(dt.DefaultView.Item(i).Item("TagsCount")))
                Dim strName As New Text.StringBuilder(Trim(dt.DefaultView(i).Item("Tag_Name")))
                strName.Replace("'", "\'")

                Select Case strCount.Length
                    Case 1 : strCount.Insert(0, Space(5))
                    Case 2 : strCount.Insert(0, Space(4))
                    Case 3 : strCount.Insert(0, Space(3))
                    Case 4 : strCount.Insert(0, Space(2))
                    Case 5 : strCount.Insert(0, Space(1))
                End Select

                strTmp.Append("'")
                strTmp.Append(strName)
                strTmp.Append(strCount)
                strTmp.Append("',")
            Next

            Return strTmp.ToString(0, strTmp.Length - 1)

        End Function



        Function ShowNewStar(ByVal grade, ByVal allCount, ByVal Id) As String
            Dim d_grade As Double
            If IsNumeric(grade) Then
                d_grade = grade
            Else
                d_grade = 0
            End If
            Dim output, img_name As String




            If IsNumeric(d_grade) And allCount > 0 Then
                d_grade = Math.Round(d_grade / allCount, 1)
            Else
                d_grade = 0
            End If
            'Current.Response.Write("m=" & d_grade)
            'Current.Response.End()
            output = ""

            Dim count As Integer = Fix(d_grade)
            Dim j, g As Integer
            Dim half As Integer = 0
            If d_grade Mod count > 0 Then
                half = 1
            End If

            For i As Integer = 0 To count - 1
                j = j + 1
                output = output & "<a title='דרג/י את השידור' href=" & Current.Application("VirDir") & "/programs/scale.aspx?MediaId=" & Id & "&Star=" & i & " onmouseover=""showStars('" & Id & "','" & i & "')""   onMouseOut=""RefreshStars('" & Id & "','" & count & "','" & half & "')"" ><img src='" & Current.Application("VirDir") & "/images/star.gif' Id='Star_" & Id & "_" & i & "' name='Star_" & Id & "_" & i & "' border=0 width=15 height=16 style='position:relative;top:2px;'></a>"
            Next
            If d_grade Mod count > 0 Then
                ' half = 1
                j = j + 1
                output = output & "<a title='דרג/י את השידור' href=" & Current.Application("VirDir") & "/programs/scale.aspx?MediaId=" & Id & "&Star=" & j - 1 & " onmouseover=""showStars('" & Id & "','" & j - 1 & "')""   onMouseOut=""RefreshStars('" & Id & "','" & count & "','" & half & "')""><img src='" & Current.Application("VirDir") & "/images/star_1_2.gif'  Id='Star_" & Id & "_" & j - 1 & "' name='Star_" & Id & "_" & j - 1 & "' border=0 width=15 height=16  style='position:relative;top:2px;'></a>"
            End If
            If j <= 4 Then
                For g = j To 4
                    output = output & "<a title='דרג/י את השידור' href=" & Current.Application("VirDir") & "/programs/scale.aspx?MediaId=" & Id & "&Star=" & g & " onmouseover=""showStars('" & Id & "','" & g & "')""   onMouseOut=""RefreshStars('" & Id & "','" & count & "','" & half & "')""><img src='" & Current.Application("VirDir") & "/images/star_dis.gif'  Id='Star_" & Id & "_" & g & "' name='Star_" & Id & "_" & g & "' border=0 width=15 height=16 style='position:relative;top:2px;'></a>"

                Next


            End If

            If Len(output) = 0 Then
                output = "<span class=new_desc>&nbsp;</span>"
            End If
            '    output = d_grade & "===" & allCount
            ShowNewStar = output

        End Function
        Public Function Show_FullName(ByVal MemberCode As String) As String
            Dim result As String
            Dim conR As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            Dim cmdSelect As New SqlClient.SqlCommand("Member_get_fullname", conR)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@MemberId", SqlDbType.Int).Value = CInt(MemberCode)
            conR.Open()
            Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
            If dr.Read() Then
                If Not IsDBNull(dr("FullName")) Then
                    result = Trim(dr("FullName"))
                Else
                    result = ""
                End If
            End If
            dr.Close()
            conR.Close()
            Show_FullName = result
        End Function

        ''''''''Public Function ShowLink(ByVal MemberCode As String) As String
        ''''''''    Dim result As String
        ''''''''    Dim cookieCode, username As String
        ''''''''    If Not IsNothing(Current.Request.Cookies("103FM")) Then
        ''''''''        cookieCode = Current.Request.Cookies("103FM")("memberCode")
        ''''''''    Else
        ''''''''        cookieCode = "0"
        ''''''''    End If
        ''''''''    If Not IsNothing(Current.Request.Cookies("103FM")) Then
        ''''''''        username = Current.Request.Cookies("103FM")("username")
        ''''''''    Else
        ''''''''        username = ""
        ''''''''    End If

        ''''''''    If Trim(MemberCode) = Trim(cookieCode) Then
        ''''''''        result = "href='/members/home.aspx'"
        ''''''''    ElseIf Trim(MemberCode) <> "0" Then
        ''''''''        result = "href='/membersPages/default.aspx?" & Encode("memberCode=" & MemberCode) & "'"
        ''''''''    Else
        ''''''''        result = "nohref"
        ''''''''    End If
        ''''''''    Return result
        ''''''''End Function
        Public Sub SetSiteCookie(ByVal MemberID As String)

            If IsNumeric(MemberID) Then

                Dim dtrChk As SqlClient.SqlDataReader
                Dim cmdStr, email, password, name As String
                Dim con = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
                con.Open()
                cmdStr = "Select email, password, fname + ' ' + lname as name " & _
                " From members Where member_id = " & MemberID
                Dim cmdSelect As New SqlClient.SqlCommand(cmdStr, con)
                dtrChk = cmdSelect.ExecuteReader()
                If dtrChk.Read() Then
                    If Not IsDBNull(dtrChk.Item("email")) Then
                        email = dtrChk.Item("email")
                    End If
                    If Not IsDBNull(dtrChk.Item("password")) Then
                        password = dtrChk.Item("password")
                    End If
                    If Not IsDBNull(dtrChk.Item("name")) Then
                        name = dtrChk.Item("name")
                    End If
                    '                    Current.Response.Write("name=" & name)

                    Current.Request.Cookies.Clear()
                    Dim objC As HttpCookieCollection = Current.Request.Cookies
                    If Not IsNothing(objC("103FM")) Then
                        objC.Remove("103FM")
                    End If
                    Dim motCookie As New HttpCookie("103FM")
                    motCookie(Encode("MemberID")) = Encode(MemberID)
                    motCookie(Encode("Email")) = Encode(email)
                    motCookie(Encode("Password")) = Encode(password)
                    motCookie("MemberName") = hebrewUrlEncode(CStr(name))
                    'Current.Response.Write("tt=" & hebrewUrlEncode(name))
                    'Current.Response.End()
                    motCookie.Expires = Now.AddDays(100)
                    If Not IsNothing(Current.Response.Cookies("103FM")) Then
                        Current.Response.Cookies.Remove("103FM")
                    End If

                    Current.Response.Cookies.Add(motCookie)

                End If
                dtrChk.Close()
                con.Close()
                ' Current.Response.Write("m=" & MemberID)
                ' Current.Response.End()


            End If
        End Sub





    End Class
End Namespace