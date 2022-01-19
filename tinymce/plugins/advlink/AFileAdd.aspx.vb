Imports System.Configuration
Public Class AFileAdd
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
        If Not IsNumeric(Request.Cookies(Request.Cookies("bizpegasus")("UserId"))) Then
            Response.Redirect("../../../")
        End If

        field_name = Trim(Request.QueryString("field_name"))
        type = Trim(Request.QueryString("type"))
        If Len(type) = 0 Then
            type = "file"
        End If

        str_mappath = Application("VirDir") & "/download/files"

        Server.ScriptTimeout = 3000

        Page.Validate()

        If Request.Files.Count > 0 And Page.IsValid Then

            If Request.Files("UploadFile1").ContentLength > 0 Then
                Dim Ext As String = IO.Path.GetExtension(Request.Files("UploadFile1").FileName)
                Dim FileNameWithoutExt As String = IO.Path.GetFileNameWithoutExtension(Request.Files("UploadFile1").FileName)
                Dim fs As IO.File
                Dim i As Integer = 1
                strFileName = FileNameWithoutExt
                Do While IO.File.Exists(Server.MapPath(str_mappath & "/" & strFileName & Ext))
                    strFileName = FileNameWithoutExt & "_" & CStr(i)
                    i = i + 1
                Loop
                strFileName = strFileName & Ext
                strSavePath = Server.MapPath(str_mappath & "/" & strFileName)
                Request.Files("UploadFile1").SaveAs(strSavePath)
                fs = Nothing

                ' Dim scrp As String = "<SCRIPT LANGUAGE=javascript>" & vbCrLf & _
                '"<!--" & vbCrLf & _
                '"   window.opener.document.forms[0].elements['" & field_name & "'].value = '" & "http://" & Request.ServerVariables("Server_name") & str_mappath & "/" & strFileName & "';" & vbCrLf & _
                '"   self.close(); " & vbCrLf & _
                '"//-->" & vbCrLf & _
                '"</SCRIPT>"
                Dim scrp As String = "<SCRIPT LANGUAGE=javascript>" & vbCrLf & _
              "<!--" & vbCrLf & _
              "   window.opener.document.getElementById('" & field_name & "').value = '" & "https://" & Request.ServerVariables("Server_name") & str_mappath & "/" & strFileName & "';" & vbCrLf & _
              "   self.close(); " & vbCrLf & _
              "//-->" & vbCrLf & _
              "</SCRIPT>"
                '  Page.ClientScript.RegisterStartupScript(Page.GetType, , "onload", scrp)
                Page.RegisterStartupScript("onload", scrp)

         
                'RegisterStartupScript("onload", scrp)
            End If
        End If
    End Sub

    Private Sub InitializeComponent()

    End Sub
End Class
