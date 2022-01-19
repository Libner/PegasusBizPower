Imports System.Text.RegularExpressions
Public Class addResponse
    Inherits System.Web.UI.Page
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected UserID, OrgID, sqlstr, UseBizLogo, PageLang, _
    PRODUCT_TYPE, fromEmail, fromName, quest_id, questName, pr_language, _
    subjectEmail, FILE_ATTACHMENT, PageTitle, pagesource, formLink, first_slice, second_slice As String
    Protected prodId, responseID, lang_id As Integer
    Protected func As New bizpower.cfunc
    Protected strLocal As String = ""
    Dim con As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New System.Data.SqlClient.SqlCommand
    Public Response_Title, Response_Content As String
    Public dir_var, align_var, dir_obj_var As String
    Protected WithEvents btnSubmit As HtmlControls.HtmlInputButton
    Protected WithEvents aFile1, aFile2, aFile3 As Web.UI.HtmlControls.HtmlAnchor
    Protected fileuploadF1, fileuploadF2, fileuploadF3 As HtmlControls.HtmlInputFile
    Protected WithEvents btnDeleteFile1, btnDeleteFile2, btnDeleteFile3 As System.Web.UI.WebControls.Button
    Public str_alert As String
    Public arrTitles() As String
    Public arrButtons() As String

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        UserID = Trim(Trim(Request.Cookies("bizpegasus")("UserID")))
        OrgID = Trim(Trim(Request.Cookies("bizpegasus")("OrgID")))
        lang_id = func.fixNumeric(Request.Cookies("bizpegasus")("LANGID"))

        If lang_id = 0 Then
            lang_id = 1
        End If
        If lang_id = 2 Then
            dir_var = "rtl" : align_var = "left" : dir_obj_var = "ltr"
        Else
            dir_var = "ltr" : align_var = "right" : dir_obj_var = "rtl"
        End If
        Server.ScriptTimeout = 6000

        strLocal = "http://" & Trim(Request.ServerVariables("SERVER_NAME"))
        If Len(Trim(Request.ServerVariables("SERVER_PORT"))) > 0 Then
            strLocal = strLocal & ":" & Trim(Request.ServerVariables("SERVER_PORT"))
        End If
        strLocal = strLocal & Application("VirDir") & "/"

        btnDeleteFile1.Attributes.Add("onclick", "return window.confirm('האם ברצונך למחוק את הקובץ');")
        btnDeleteFile1.Visible = False

          sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 71 Order By word_id"
        cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
        con.Open()
        Dim arrTitlesD As String = ""
        Dim rstitle As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        Do While rstitle.Read
            arrTitlesD = arrTitlesD & ";" & Trim(rstitle("word"))
        Loop
        con.Close()

        'If arrTitlesD <> "" Then
        '    arrTitlesD = arrTitlesD.Substring(1)
        'End If
        arrTitles = arrTitlesD.Split(";")

        btnSubmit.Value = arrTitles(1)



        sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
        cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
        con.Open()
        Dim arrButtonsD As String = ""
        Dim rsButton As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        Do While rsButton.Read
            arrButtonsD = arrButtonsD & ";" & Trim(rsButton("value"))
        Loop
        con.Close()
        'If arrButtonsD <> "" Then
        '    arrButtonsD = arrButtonsD.Substring(1)
        'End If
        arrButtons = Split(arrButtonsD, ";")


        prodId = func.fixNumeric(Request.QueryString("prodID"))
        responseID = func.fixNumeric(Request.QueryString("responseID"))

        If Not Page.IsPostBack Then
            If UserID <> "" Then

                If Request.QueryString("responseID") <> "" Then
                     If Len(responseID) > 0 Then
                        sqlstr = "Select Response_Title, Response_Content,Response_File_1,Response_File_2,Response_File_3 From Product_Responses Where Response_ID = " & _
       responseID & " And Product_ID = " & prodId
                        cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
                        con.Open()
                        Dim rs_product As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
                        If rs_product.Read() Then
                            Response_Title = Trim(rs_product("Response_Title"))
                            Response_Content = Trim(rs_product("Response_Content"))

                            If Not IsDBNull(rs_product("Response_File_1")) Then
                                Dim FileSrc As String = Trim(rs_product("Response_File_1"))
                                If Len(FileSrc) > 0 Then
                                    aFile1.HRef = "../../../download/files/" & CStr(FileSrc)
                                    aFile1.InnerText = CStr(FileSrc)
                                    btnDeleteFile1.Visible = True
                                End If
                            End If
                            If Not IsDBNull(rs_product("Response_File_2")) Then
                                Dim FileSrc As String = Trim(rs_product("Response_File_2"))
                                If Len(FileSrc) > 0 Then
                                    aFile2.HRef = "../../../download/files/" & CStr(FileSrc)
                                    aFile2.InnerText = CStr(FileSrc)
                                    btnDeleteFile2.Visible = True
                                End If
                            End If
                            If Not IsDBNull(rs_product("Response_File_3")) Then
                                Dim FileSrc As String = Trim(rs_product("Response_File_3"))
                                If Len(FileSrc) > 0 Then
                                    aFile3.HRef = "../../../download/files/" & CStr(FileSrc)
                                    aFile3.InnerText = CStr(FileSrc)
                                    btnDeleteFile3.Visible = True
                                End If
                            End If
                        End If
                        con.Close()

                    End If
                End If
            End If
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.ServerClick
        SaveData(True)
    End Sub

    Private Sub SaveData(ByVal Redirect)
        prodId = func.fixNumeric(Request.Form("prodID"))
        responseID = func.fixNumeric(Request.Form("responseID"))
        If Trim(Request.Form("responseID")) = 0 Then ' add type
            sqlstr = "Insert into Product_Responses (Product_ID,Organization_ID,Response_Title,Response_Content) values (" & _
   prodId & "," & OrgID & ",'" & func.sFix(Request.Form("Response_Title")) & "','" & func.sFix(Request.Form("Response_Content")) & "');select @@IDENTITY"
            Dim cmdInsert As New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            Try
                responseID = cmdInsert.ExecuteScalar
            Catch ex As Exception
            End Try
            con.Close()

        Else ' update type
            responseID = Trim(Request.Form("responseID"))
            sqlstr = "Update Product_Responses set Response_Title = '" & func.sFix(Request.Form("Response_Title")) & _
            "', Response_Content = '" & func.sFix(Request.Form("Response_Content")) & "'" & _
            " Where Response_ID = " & responseID & " And Product_ID = " & prodId
            Dim cmdInsert As New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            Try
                cmdInsert.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            con.Close()
        End If


        If CInt(responseID) > 0 Then

            '----------------------------------
            'file
            Dim pFile As System.IO.File

            If Not (fileuploadF1.PostedFile Is Nothing) And (fileuploadF1.PostedFile.ContentLength > 0) Then 'Check to make sure we actually have a file to upload
                If IsNumeric(responseID) Then
                    'Delete existing file
                    If pFile.Exists(Server.MapPath(aFile1.HRef)) Then
                        pFile.Delete(Server.MapPath(aFile1.HRef))
                    End If
                End If
                Dim strFileName As String = fileuploadF1.PostedFile.FileName
                Dim strFileNameExt As String = IO.Path.GetExtension(fileuploadF1.PostedFile.FileName)
                If pFile.Exists(Server.MapPath("../../../download/files/" & strFileName)) Then
                    Dim i As Integer = 0
                    Do While pFile.Exists(Server.MapPath("../../../download/files/" & strFileName))
                        i = i + 1
                        strFileName = IO.Path.GetFileNameWithoutExtension(fileuploadF1.PostedFile.FileName) & "_" & i & strFileNameExt
                    Loop
                End If
                Dim strSavePath As String = Server.MapPath("../../../download/files/" & strFileName)

                fileuploadF1.PostedFile.SaveAs(strSavePath)

                Dim cmdUpdate As New SqlClient.SqlCommand("Update Product_Responses Set Response_File_1 = '" & func.sFix(strFileName) & "' Where Response_ID = " & responseID, con)
                con.Open()
                cmdUpdate.ExecuteNonQuery()
                cmdUpdate.Dispose()
                con.Close()
            End If

            If Not (fileuploadF2.PostedFile Is Nothing) And (fileuploadF2.PostedFile.ContentLength > 0) Then 'Check to make sure we actually have a file to upload
                If IsNumeric(responseID) Then
                    'Delete existing file
                    If pFile.Exists(Server.MapPath(aFile2.HRef)) Then
                        pFile.Delete(Server.MapPath(aFile2.HRef))
                    End If
                End If
                Dim strFileName As String = fileuploadF2.PostedFile.FileName
                Dim strFileNameExt As String = IO.Path.GetExtension(fileuploadF2.PostedFile.FileName)
                If pFile.Exists(Server.MapPath("../../../download/files/" & strFileName)) Then
                    Dim i As Integer = 0
                    Do While pFile.Exists(Server.MapPath("../../../download/files/" & strFileName))
                        i = i + 1
                        strFileName = IO.Path.GetFileNameWithoutExtension(fileuploadF2.PostedFile.FileName) & "_" & i & strFileNameExt
                    Loop
                End If
                Dim strSavePath As String = Server.MapPath("../../../download/files/" & strFileName)

                fileuploadF1.PostedFile.SaveAs(strSavePath)

                Dim cmdUpdate As New SqlClient.SqlCommand("Update Product_Responses Set Response_File_2 = '" & func.sFix(strFileName) & "' Where Response_ID = " & responseID, con)
                con.Open()
                cmdUpdate.ExecuteNonQuery()
                cmdUpdate.Dispose()
                con.Close()
            End If

            If Not (fileuploadF3.PostedFile Is Nothing) And (fileuploadF3.PostedFile.ContentLength > 0) Then 'Check to make sure we actually have a file to upload
                If IsNumeric(responseID) Then
                    'Delete existing file
                    If pFile.Exists(Server.MapPath(aFile3.HRef)) Then
                        pFile.Delete(Server.MapPath(aFile3.HRef))
                    End If
                End If
                Dim strFileName As String = fileuploadF3.PostedFile.FileName
                Dim strFileNameExt As String = IO.Path.GetExtension(fileuploadF3.PostedFile.FileName)
                If pFile.Exists(Server.MapPath("../../../download/files/" & strFileName)) Then
                    Dim i As Integer = 0
                    Do While pFile.Exists(Server.MapPath("../../../download/files/" & strFileName))
                        i = i + 1
                        strFileName = IO.Path.GetFileNameWithoutExtension(fileuploadF3.PostedFile.FileName) & "_" & i & strFileNameExt
                    Loop
                End If
                Dim strSavePath As String = Server.MapPath("../../../download/files/" & strFileName)

                fileuploadF3.PostedFile.SaveAs(strSavePath)

                Dim cmdUpdate As New SqlClient.SqlCommand("Update Product_Responses Set Response_File_3 = '" & func.sFix(strFileName) & "' Where Response_ID = " & responseID, con)
                con.Open()
                cmdUpdate.ExecuteNonQuery()
                cmdUpdate.Dispose()
                con.Close()
            End If


        End If

        If Redirect Then
            Response.Redirect("addresponse.aspx?responseID=" & responseID & "&prodID=" & prodId)
        Else
            Dim cScript As String
            cScript = "<script language='javascript'> opener.location.href = opener.location.href; self.close(); </script>"
            RegisterStartupScript("ReloadScrpt", cScript)
        End If
        

    End Sub

    Sub Delete_File(ByVal sender As Object, ByVal e As EventArgs) Handles btnDeleteFile1.Click, btnDeleteFile2.Click, btnDeleteFile3.Click

        Dim numFile As String = CType(sender, Button).CommandArgument
        Dim aFile As HtmlAnchor = CType(Page.FindControl("aFile" & numFile), HtmlAnchor)
           If aFile.HRef.Length > 0 Then
            Dim con As SqlClient.SqlConnection
            Dim cmdUpdate As SqlClient.SqlCommand

            Dim pFile As System.IO.File
            If pFile.Exists(Server.MapPath(aFile.HRef)) Then
                pFile.Delete(Server.MapPath(aFile.HRef))
            End If

            con = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            cmdUpdate = New SqlClient.SqlCommand("Update Product_Responses Set Response_File_" & numFile & " = NULL WHERE Response_ID = " & responseID, con)
            cmdUpdate.CommandType = CommandType.Text
            con.Open()
            cmdUpdate.ExecuteNonQuery()
            cmdUpdate.Dispose()
            con.Close()
        End If
        Response.Redirect("AddResponse.aspx?prodId=" & CStr(prodId) & "&responseID=" & responseID)
    End Sub
End Class
