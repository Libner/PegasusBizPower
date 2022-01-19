Public Class UploadFile
    Inherits System.Web.UI.Page
    Protected DepartureId As Integer
    Public func As New bizpower.cfunc
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected conBIns As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

    Protected fileupload1 As HtmlControls.HtmlInputFile
    Protected WithEvents fselect As HtmlSelect
    Protected dateP, UserName As HtmlInputText
    Protected Userid, WorkerId As String
    Protected fileDescr As HtmlTextArea
    Protected checkSup As PlaceHolder
    Protected DepatureCode, supplierEmail As String
    Protected isGuaranteed As Boolean
    Protected GLog As Repeater
    Protected EmailContent As String
    Protected divfselectError, divSelect As PlaceHolder




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
         If IsNumeric(Request("Userid")) Then
            Userid = CInt(Request("Userid"))
        End If
        '    Response.Write("u=" & Request("Userid") & ":" & Userid)


        dateP.Value = Now()
        WorkerId = Trim(Request.Cookies("bizpegasus")("UserId"))
        UserName.Value = func.GetSelectUserName(WorkerId)

        If Page.IsPostBack Then
            SaveData()
        End If
    End Sub
   
 
    Public Sub SaveData()
        Dim WorkerId = Trim(Request.Cookies("bizpegasus")("UserId"))
        Dim Userid = Request("Userid")
        Dim path As String = Server.MapPath("~/Download/usersFiles/")
        Dim fileOK As Boolean = False

        '  Response.Write("22=" & fileupload1.PostedFile.ContentLength)

        Dim fileName As String = ""
        If Not (fileupload1.PostedFile Is Nothing) And (fileupload1.PostedFile.ContentLength > 0) Then 'Check to make sure we actually have a file to upload
            '' Response.Write(fileupload1.PostedFile.ContentLength)

            Dim fileExtension As String
            fileExtension = System.IO.Path.GetExtension(fileupload1.PostedFile.FileName).ToLower()

            Dim allowedExtensions As String() = _
                {".jpg", ".jpeg", ".png", ".gif", ".pdf", ".docx", ".doc", ".xls", ".xlsx", ".txt", ".msg"}
            For i As Integer = 0 To allowedExtensions.Length - 1
                ''Response.Write(allowedExtensions(i) & "<BR>")
                If fileExtension = allowedExtensions(i) Then
                    fileOK = True
                End If
            Next
        End If

        If fileOK Then
            Dim originName = Replace(System.IO.Path.GetFileName(fileupload1.PostedFile.FileName), System.IO.Path.GetExtension(fileupload1.PostedFile.FileName), "")
            '  Response.Write(originName)
            '  Response.Write("Userid=" & Userid)

            '  Response.End()
            '
            fileName = originName & "_" & Userid & "_" & Day(Now()) & Month(Now()) & Year(Now) & "_" & Hour(Now) & "_" & Minute(Now()) & "_" & Second(Now()) & System.IO.Path.GetExtension(fileupload1.PostedFile.FileName)
            '       Response.Write("<BR>" & fileName)
            '      Response.End()
            ' fileupload1.PostedFile.SaveAs(path & _
            '      fileupload1.PostedFile.FileName)
            fileupload1.PostedFile.SaveAs(path & _
                 fileName)
            '''    fileName = System.IO.Path.GetExtension(fileupload1.PostedFile.FileName)
        Else
            'Response.Write("file error")
            'Response.End()

            RegisterStartupScript("scriptName_FreeText", "<script type=""text/javascript"">alert('סוג קובץ לא תקין')</script>")
             End If


        ''If Request.Form("fselect") > 0 Then   'save only one supplier_Id
        ''    supplier_Id = Request.Form("fselect")
        ''    supplierName = func.GetSelectSupplierName(supplier_Id)
        ''    supplierUserName = func.GetSelectSupplierUserName(supplier_Id)
        ''    supplierEmail = func.GetSelectSupplierUserEmail(supplier_Id)
        ''    'supplierGUID = func.GetSelectSupplierGUID(supplier_Id)
        ''    ''Insert ''

        If fileOK Then
            cmdSelect = New SqlClient.SqlCommand("Ins_UploadUserFile", conB)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(Userid)
            cmdSelect.Parameters.Add("@WorkerId", SqlDbType.Int).Value = CInt(WorkerId)
            cmdSelect.Parameters.Add("@FileName", SqlDbType.NVarChar, 200).Value = fileName
            cmdSelect.Parameters.Add("@date", SqlDbType.DateTime).Value = Now()
            cmdSelect.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = fileDescr.Value

            conB.Open()
            cmdSelect.ExecuteNonQuery()
            conB.Close()
        End If
        ''    '''----
        ''    ''---  Send Mail



        If fileOK Then

            Dim cScript As String

            cScript = "<script language='javascript'>alert('הקובץ הועלה בהצלחה');window.opener.location.reload(false);self.close();</script>"
            RegisterStartupScript("ReloadScrpt", cScript)

        End If

    End Sub

End Class
