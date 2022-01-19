Imports System.Data.SqlClient
Imports System.Web.Mail
Public Class SendMail1
    Inherits System.Web.UI.Page
    Protected topIncludeU As topInclude
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dtUsers, dtContacts, dtBCCUsers, dtDep As New DataTable
    Dim primKeydtUsers(0), primKeydtBCCUsers(0), primKeydtContacts(0), primKeydtDep(0) As Data.DataColumn
    Protected rowId, AirlineId As Integer
    Protected AirlineCode, DEPARTURECode, QUANTITYNow As String
    Dim objMail As New System.Web.Mail.MailMessage
    Protected func As New bizpower.cfunc
    Dim UserEmail, selWaitingAirlineResponse As String
    Protected PNR As String
    Public UserId, DepIds As String
    Protected pwidth As String
    Protected BaseUrl As String
    Protected WorkerId As Integer
    Protected FromEmail As String
    Dim ContactEmail, ContactEmailToSend, WaitingAirlineResponse, pFrom As String
    Protected SeriesId, DepId As Integer
    Protected defUserId, defDepId As String
    Public statusMail As String = "successfully"
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents fileupload1 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents content As System.Web.UI.HtmlControls.HtmlTextArea
    Protected WithEvents sContent As System.Web.UI.HtmlControls.HtmlTextArea
    Protected SelSearchUsers, SelSearchDep As System.Web.UI.HtmlControls.HtmlSelect
    Protected plhForm, plhResultSuccess, plhResultError As PlaceHolder


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
        WorkerId = Request.Cookies("bizpegasus")("UserId")

        If IsNumeric(WorkerId) Then
            FromEmail = func.GetUserEmail(WorkerId)
            DepId = func.GetDepId(WorkerId)
        End If



        BaseUrl = Application("VirDir") & "/" 'Request.ApplicationPath & "/"

        topIncludeU = CType(Page.FindControl("topInclude"), topInclude)
        topIncludeU = CType(Page.FindControl("topInclude"), topInclude)
        topIncludeU.numOfTab = 127
        topIncludeU.numOfLink = 1
        pwidth = "600"
        If Not Page.IsPostBack Then
            'GetData()
            getDep()
            GetUsers()
            'GetBCCUsers()
            'GetContacts()


        Else

            If Request.Form.Count > 0 Then
                SendMail()
            End If

        End If
    End Sub
    Public Sub SendMail()
        ''''''   Dim userId = Request.Form("SelSearchUsers")
        ''''''Dim rptD = Request.Form("rptD")
        Dim CCMail As String
        Dim ToMail As String
        Dim BCCMail As String
        Dim pSubject As String
        Dim content As String

        CCMail = Request.Form("pCC")
        ToMail = Request.Form("SelSearchUsers")
        BCCMail = Request.Form("pBCC")
        pFrom = Request.Form("pFrom")
        pSubject = Request.Form("pSubject")
        content = Request.Form("sContent")

     
        'Response.Write("From=" & pFrom & "<BR>BCCMail=" & BCCMail)
        'Response.Write("CCMail=" & CCMail & "<BR>to=" & ToMail)
        'Response.Write("file=" & fileupload1.PostedFile.ContentLength)
        '   Response.End()



        '''''''fileName  upload
        Dim path As String = Server.MapPath("~/Download/SendEmail/")
        Dim fileName As String = ""
        If Not (fileupload1.PostedFile Is Nothing) And (fileupload1.PostedFile.ContentLength > 0) Then 'Check to make sure we actually have a file to upload
            Dim fileOK As Boolean = False
            Dim fileExtension As String
            fileExtension = System.IO.Path.GetExtension(fileupload1.PostedFile.FileName).ToLower()

            Dim allowedExtensions As String() = _
                {".jpg", ".jpeg", ".png", ".gif", "pdf", "docx", "pdf"}
            For i As Integer = 0 To allowedExtensions.Length - 1
                If fileExtension = allowedExtensions(i) Then
                    fileOK = True
                End If
            Next

            fileOK = True

            If fileOK Then
                fileupload1.PostedFile.SaveAs(path & _
                      fileupload1.PostedFile.FileName)
                fileName = fileupload1.PostedFile.FileName
            End If
        End If

        Try

            objMail.From = pFrom
            objMail.Cc = CCMail
            objMail.Bcc = BCCMail
            objMail.To = ToMail
            objMail.To = "erez.kosher@pegasusisrael.co.il" '"faina@cyberserve.co.il"
            objMail.Subject = pSubject
            objMail.Body = func.breaks(content)
            objMail.BodyFormat = MailFormat.Html
            
            objMail.BodyEncoding = System.Text.Encoding.UTF8
            If Len(fileName) > 0 Then
                objMail.Attachments.Add(New System.Web.Mail.MailAttachment(path & fileName))
            End If
            '    SmtpMail.SmtpServer = "Josh-vm"
            SmtpMail.Send(objMail)


            plhForm.Visible = False
            plhResultSuccess.Visible = True

        Catch ex As Exception
            statusMail = ex.Message.ToString()
            Response.Write(statusMail)
            objMail = Nothing
            plhForm.Visible = False
            plhResultError.Visible = True

        End Try
        objMail = Nothing
     

    End Sub

    Public Sub getDep()
        Dim squery As String
        squery = "SELECT departmentId,departmentName FROM Departments ORDER BY PriorityLevel desc ,departmentName"
        ' and Department_Id=" & DepId & "
        '   Response.Write(squery)

        Dim cmdSelect As New SqlClient.SqlCommand(squery, con)

        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtDep)
        'sSRSCODE
        con.Close()
        primKeydtDep(0) = dtDep.Columns("departmentId")
        dtDep.PrimaryKey = primKeydtDep
        SelSearchDep.Items.Clear()
   

        For i As Integer = 0 To dtDep.Rows.Count - 1
            Dim list As New ListItem(dtDep.Rows(i)("departmentName"), dtDep.Rows(i)("departmentId"))
            '  If Trim(Request.Cookies("bizpegasus")("IsSendFromTL")) = 1 Then
            If defDepId = dtDep.Rows(i)("departmentId") Or DepIds = "" Then
                list.Selected = True
            End If

            '  End If
            SelSearchDep.Items.Add(list)
        Next

    End Sub

    Public Sub GetUsers()
        'If Request.Form("sUsers") > 0 Then
        '    defUserId = Request.Form("sUsers")
        'Else
        '    defUserId = 977
        'End If
        Dim squery As String

        squery = "select  User_Id,Email,(FIRSTNAME + Char(32) + LASTNAME) as [USER_NAME] from Users where  len(email)>0 and ACTIVE=1 order by USER_NAME"
        ' and Department_Id=" & DepId & "
        '   Response.Write(squery)

        Dim cmdSelect As New SqlClient.SqlCommand(squery, con)

        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtUsers)
        'sSRSCODE
        con.Close()
        primKeydtUsers(0) = dtUsers.Columns("Email")
        dtUsers.PrimaryKey = primKeydtUsers
        SelSearchUsers.Items.Clear()
        'If Trim(Request.Cookies("bizpegasus")("IsSendFromTL")) = 1 Then
        '    Dim list1 As New ListItem("All", "0")
        '    SelSearchUsers.Items.Add(list1)
        'End If

        For i As Integer = 0 To dtUsers.Rows.Count - 1
            Dim list As New ListItem(dtUsers.Rows(i)("USER_NAME"), dtUsers.Rows(i)("Email"))
            '  If Trim(Request.Cookies("bizpegasus")("IsSendFromTL")) = 1 Then
            If defUserId = dtUsers.Rows(i)("Email") Then
                list.Selected = True
            End If

            '  End If
            SelSearchUsers.Items.Add(list)
        Next



    End Sub

End Class
