Public Class Send_FeedbackByEmailToSuppliers
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected DepartureId, supplierId As Integer
    Protected supplierEmail, DepartureCode, supplierName, supplierUserName, EmailContent As String
    Public func As New bizpower.cfunc

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
        Dim cmdSelect As String = "select distinct FF.Departure_Id,Departure_Code from FeedBack_Form FF " & _
        " left join Tours_Departures TD on FF.Departure_Id=TD.Departure_Id where  DateDiff(dd, Feedback_Date, GetDate()) = 1"

        '==1 ----left join Tours_Departures TD  on TD.Departure_Id=FF.Departure_Id
        'where (isCanceled=0 and datediff(dd,Departure_Date_End,GetDate())>=1)

        Dim dt As New DataTable
        conPegasus.Open()

        Dim dta = New SqlClient.SqlDataAdapter(cmdSelect, conPegasus)
        dt.BeginLoadData()
        dta.Fill(dt)
        dt.EndLoadData()
        conPegasus.Close()
        Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        For i As Integer = 0 To dt.DefaultView.Count - 1
            DepartureId = dt.DefaultView(i).Item("Departure_Id")
            DepartureCode = dt.DefaultView(i).Item("Departure_Code")
            supplierEmail = ""
            Dim srtSuppl = New SqlClient.SqlCommand("select VouchersToSuppliers.supplier_Id,Departure_Id,Vouchers_Status," & _
            " supplier_Email1 , supplier_Email2, supplier_Email3, supplier_Email4  from VouchersToSuppliers left join Suppliers  " & _
            " on Suppliers.supplier_Id=VouchersToSuppliers.supplier_Id where Departure_Id=@DepartureId", conU)
            srtSuppl.CommandType = CommandType.Text
            srtSuppl.Parameters.Add("@DepartureId", SqlDbType.Int).Value = DepartureId
            conU.Open()
            Dim dr_m As SqlClient.SqlDataReader = srtSuppl.ExecuteReader()

            While dr_m.Read
                If Not dr_m("supplier_Id") Is DBNull.Value Then
                    supplierId = dr_m("supplier_Id")
                End If
                If Not dr_m("supplier_Email1") Is DBNull.Value Then
                    If Trim(dr_m("supplier_Email1")) <> "" Then
                        supplierEmail = dr_m("supplier_Email1")
                    End If
                End If

                If Not dr_m("supplier_Email2") Is DBNull.Value Then
                    If supplierEmail = "" Then
                        If Trim(dr_m("supplier_Email2")) <> "" Then
                            supplierEmail = dr_m("supplier_Email2")
                        End If
                    Else
                        If Trim(dr_m("supplier_Email2")) <> "" Then
                            supplierEmail = supplierEmail & "," & dr_m("supplier_Email2")
                        End If
                    End If
                End If


                If Not dr_m("supplier_Email3") Is DBNull.Value Then
                    If supplierEmail = "" Then
                        supplierEmail = dr_m("supplier_Email3")
                    Else
                        If Trim(dr_m("supplier_Email3")) <> "" Then
                            supplierEmail = supplierEmail & "," & dr_m("supplier_Email3")
                        End If
                        End If
                End If
                If Not dr_m("supplier_Email4") Is DBNull.Value Then
                    If supplierEmail = "" Then
                        supplierEmail = dr_m("supplier_Email4")
                    Else
                        If Trim(dr_m("supplier_Email4")) <> "" Then
                            supplierEmail = supplierEmail & "," & dr_m("supplier_Email4")
                        End If
                        End If
                End If

            End While
            dr_m.Close()
            srtSuppl.Dispose()
            conU.Close()
            If Trim(supplierEmail) <> "" Then
                '   Response.Write("supplierEmail=" & supplierEmail)
                supplierName = func.GetSelectSupplierName(supplierId)
                supplierUserName = func.GetSelectSupplierUserName(supplierId)

                EmailContent = "Dear " & supplierUserName & ".<BR>Please enter Your Pegasus account to view the customers feedbacks regarding group  " & DepartureCode
                EmailContent = EmailContent & "<BR>" & "Click here to  <A href=https://www.pegasusisrael.co.il/suppliers>login</a>"
                EmailContent = EmailContent & "<BR>" & "Best Regards " & "<BR>" & "Pegasus operation team" '& supplierEmail

                Dim Msg As New System.Web.Mail.MailMessage
                Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                Msg.BodyEncoding = System.Text.Encoding.UTF8
                Msg.From = "info@pegasusisrael.co.il"
                Msg.Subject = DepartureCode & " Feedbacks"
                Msg.To = supplierEmail
                ' Msg.Bcc = "erez.kosher@pegasusisrael.co.il"
                'Msg.To = "erez.kosher@pegasusisrael.co.il" ' "faina@cyberserve.co.il"
                '    Msg.Cc = "faina@cyberserve.co.il,erez.kosher@pegasusisrael.co.il"
                '      Msg.To = "faina@cyberserve.co.il,furfaina@gmail.com"
                Msg.Body = EmailContent.ToString()

                Try
                    System.Web.Mail.SmtpMail.SmtpServer = ConfigurationSettings.AppSettings.Item("smtp_server")
                    System.Web.Mail.SmtpMail.Send(Msg)

                    '    Response.Write("ok")
                Catch ex As Exception

                End Try
                Msg = Nothing
            End If

        Next







    End Sub


End Class
