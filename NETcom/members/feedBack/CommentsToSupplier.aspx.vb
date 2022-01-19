Public Class CommentsToSupplier
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Protected DepartureId As Integer
    Dim cmdSelect As New SqlClient.SqlCommand

    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected conBIns As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim PegasusSiteDBName = ConfigurationSettings.AppSettings.Item("PegasusSiteDBName")

    Protected WithEvents fselect As HtmlSelect
    Protected dateP, UserName As HtmlInputText
    Protected Userid As String
    Protected supplier_Id, supplierName, supplierUserName, supplierGUID As String
    Protected DepatureCode, supplierEmail As String
    Protected EmailContent As String

    Protected ContentText As HtmlTextArea
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected rptCustomers As Repeater
    Protected deleteId As Integer

    ''  Protected WithEvents Button1 As HtmlInputButton



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
        '''    Button1.Value = "טען הערות משובים לאזור האישי של הספק/ספקים"
        If IsNumeric(Request.QueryString("deleteId")) Then
            deleteId = CInt(Request.QueryString("deleteId"))
            Dim cmdDelete As New SqlClient.SqlCommand("DELETE FROM FeedbackToSsupplier WHERE Id = @deleteId", conB)
            cmdDelete.CommandType = CommandType.Text
            cmdDelete.Parameters.Add("@deleteId", SqlDbType.Int).Value = deleteId
            conB.Open()
            '  Try
            cmdDelete.ExecuteNonQuery()
            conB.Close()
            '  Catch ex As Exception
            'Response.Write(Err.Description)
            '   Finally

            '   End Try
        End If
        dateP.Value = Now()
        Userid = Trim(Request.Cookies("bizpegasus")("UserId"))
        UserName.Value = func.GetUserNameEng(Userid)
        If IsNumeric(Request.QueryString("sdep")) Then
            DepartureId = CInt(Request.QueryString("sdep"))

            If DepartureId > 0 Then
                Dim sqlstr As String
                sqlstr = "SELECT Departure_Code, isGuaranteed  FROM Tours_Departures where Departure_Id=" & DepartureId
                Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
                Dim myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                Dim myReader = myCommand.ExecuteReader(CommandBehavior.SingleRow)
                While myReader.Read()
                    DepatureCode = myReader("Departure_Code")

                End While
                myConnection.Close()
            End If



        End If
        GetSuppliers()
        GetData()
        If Page.IsPostBack Then
            SaveData()
            GetData()

        End If
    End Sub
    Function GetSuppliers()
        If DepartureId > 0 Then
            Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id,S.supplier_Name,Departure_Id,Vouchers_Status ,Country_Name " & _
               " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
               " left join " & PegasusSiteDBName & ".dbo.Countries on " & PegasusSiteDBName & ".dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), conB)
            ''  Response.Write(cmdSelect.CommandText)
            cmdSelect.CommandType = CommandType.Text
            conB.Open()
            ''    Dim list1 As New ListItem("הכל", "0")
            fselect.DataTextField = "supplier_Name"
            fselect.DataValueField = "supplier_Id"
            fselect.DataSource = cmdSelect.ExecuteReader()
            fselect.DataBind()
            Dim list1 As New ListItem("שלח לכל הספקים", "0")
            fselect.Items.Insert(0, list1)

            conB.Close()
        End If
    End Function
    Public Sub GetData()
        Dim cmdSelect As New SqlClient.SqlCommand("select Id,Date,IsNull(supplier_Name,'כל הספקים') as supplier_Name ,Content_Text,isNull(FIRSTNAME_ENG + ' ' + LASTNAME_ENG,'Automatic System') as Username from   FeedbackToSsupplier FS left join Users on FS.User_Id=Users.User_Id left join Suppliers on FS.supplier_Id=Suppliers.supplier_Id  where  Departure_Id = @DepartureId order by Date,supplier_Name ", conB)
        cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)

        conB.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptCustomers.DataSource = dr
            rptCustomers.DataBind()

        End If
        conB.Close()

    End Sub


    Public Sub SaveData()
        Dim fileOK As Boolean = False

        Dim UserId = Trim(Request.Cookies("bizpegasus")("UserId"))

        If Request.Form("fselect") > 0 Then   'save only one supplier_Id
            supplier_Id = Request.Form("fselect")
            supplierName = func.GetSelectSupplierName(supplier_Id)
            supplierUserName = func.GetSelectSupplierUserName(supplier_Id)
            supplierEmail = func.GetSelectSupplierUserEmail(supplier_Id)
            supplierGUID = func.GetSelectSupplierGUID(supplier_Id)
            ''Insert ''
            cmdSelect = New SqlClient.SqlCommand("Ins_FeedbackListToSuppliers", conB)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@supplier_Id", SqlDbType.Int).Value = CInt(Request.Form("fselect"))
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
            cmdSelect.Parameters.Add("@ContentText", SqlDbType.NText).Value = Request.Form("ContentText")

            cmdSelect.Parameters.Add("@date", SqlDbType.DateTime).Value = Now()



            conB.Open()
            cmdSelect.ExecuteNonQuery()
            conB.Close()
            fileOK = True

            '''----
            ''---  Send Mail

            If Trim(supplierEmail) <> "" Then
                EmailContent = "Dear " & supplierUserName & ".<BR>Please enter Your Pegasus account to view the customers feedbacks regarding group " & DepatureCode & " has been sent to your personal account"
                EmailContent = EmailContent & "<BR>" & "Click here to  <A href=https://www.pegasusisrael.co.il/suppliers>login</a>"
                EmailContent = EmailContent & "<BR>" & "Best Regards " & "<BR>" & "Pegasus operation team"

                Dim Msg As New System.Web.Mail.MailMessage
                Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                Msg.BodyEncoding = System.Text.Encoding.UTF8
                Msg.From = "info@pegasusisrael.co.il"
                ' Msg.Subject = DepatureCode & " Feedbacks"
                Msg.Subject = "New Feedbacks has been uploaded to your personal zone – Group " & DepatureCode  ' DepatureCode & " Feedbacks"


                Msg.To = supplierEmail
                Msg.Body = EmailContent.ToString()
                System.Web.Mail.SmtpMail.Send(Msg)
                Msg = Nothing
            End If
            'Response.Write("t=" & EmailContent)
            'Response.End()
            ''---


        Else   'save to all suppliers
            Dim myReader As System.Data.SqlClient.SqlDataReader
            Dim cmdSelectIns As New SqlClient.SqlCommand
            Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id ,supplier_Name  from VouchersToSuppliers VS left join Suppliers S on VS.supplier_Id= S.supplier_Id where Departure_Id =  " & CInt(DepartureId), conB)
            cmdSelect.CommandType = CommandType.Text
            conB.Open()
            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("supplier_Id")) Then
                    ''Insert ''
                    If Len(supplierName) > 0 Then
                        supplierName = supplierName & ", " & myReader("supplier_Name")
                    Else
                        supplierName = myReader("supplier_Name")
                    End If
                    ''   Response.Write(Request.Form("Guaranteed"))
                    ''   Response.End()
                    cmdSelectIns = New SqlClient.SqlCommand("Ins_FeedbackListToSuppliers", conBIns)
                    cmdSelectIns.CommandType = CommandType.StoredProcedure
                    cmdSelectIns.Parameters.Add("@supplier_Id", SqlDbType.Int).Value = CInt(myReader("supplier_Id"))
                    cmdSelectIns.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                    cmdSelectIns.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
                    cmdSelectIns.Parameters.Add("@ContentText", SqlDbType.NVarChar, 1000).Value = Request.Form("ContentText")

                    cmdSelectIns.Parameters.Add("@date", SqlDbType.DateTime).Value = Now()

                    conBIns.Open()
                    cmdSelectIns.ExecuteNonQuery()
                    conBIns.Close()
                    cmdSelectIns = Nothing
                    fileOK = True
                End If
                supplierUserName = func.GetSelectSupplierUserName(supplier_Id)

                supplierEmail = func.GetSelectSupplierUserEmail(supplier_Id)
                If Trim(supplierEmail) <> "" Then
                    EmailContent = "Dear " & supplierUserName & ".<BR>A new Rooming list regarding group " & DepatureCode & " has been sent to your personal account"
                    EmailContent = EmailContent & "<BR>" & "Click here to  <A href=https://www.pegasusisrael.co.il/suppliers>login</a>"
                    EmailContent = EmailContent & "<BR>" & "Best Regards " & "<BR>" & "Pegasus operation team" & supplierEmail

                    Dim Msg As New System.Web.Mail.MailMessage
                    Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                    Msg.BodyEncoding = System.Text.Encoding.UTF8
                    Msg.From = "info@pegasusisrael.co.il"
                    Msg.Subject = "New Feedbacks has been uploaded to your personal zone – Group " & DepatureCode  ' DepatureCode & " Feedbacks"

                    Msg.To = supplierEmail
                    Msg.Body = EmailContent.ToString()
                    System.Web.Mail.SmtpMail.Send(Msg)

                    Msg = Nothing


                End If

            End While
            conB.Close()



        End If

        If fileOK Then

            Dim cScript As String

            cScript = "<script language='javascript'>alert('הוספנו הערות לטבלת המשובים לספקים הבאים:" & supplierName & "');document.location.href='CommentsToSupplier.aspx?sdep=" & DepartureId & "' ; </script>"
            RegisterStartupScript("ReloadScrpt", cScript)

        End If


    End Sub


   
End Class
