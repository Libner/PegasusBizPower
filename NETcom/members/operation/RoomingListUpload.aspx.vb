Public Class RoomingListUpload
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
    Protected Userid, supplier_Id, supplierName, supplierUserName, supplierGUID, texCheck As String
    Protected RID As Integer
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
        If IsNumeric(Request.QueryString("dID")) Then
            DepartureId = CInt(Request.QueryString("dID"))
        End If
        RID = 0

        dateP.Value = Now()
        Userid = Trim(Request.Cookies("bizpegasus")("UserId"))
        UserName.Value = func.GetUserNameEng(Userid)

        If DepartureId > 0 Then
            Dim sqlstr As String
            sqlstr = "SELECT Departure_Code, isGuaranteed  FROM Tours_Departures where Departure_Id=" & DepartureId
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
            Dim myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            Dim myReader = myCommand.ExecuteReader(CommandBehavior.SingleRow)
            While myReader.Read()
                DepatureCode = myReader("Departure_Code")
                isGuaranteed = myReader("isGuaranteed")
            End While
            myConnection.Close()
            GetSuppliers()
            CheckSupppliers()
            GetLogdata()
        End If
        If Page.IsPostBack Then
            SaveData()
        End If
    End Sub
    Function CheckSupppliers()
        Dim myReader As System.Data.SqlClient.SqlDataReader
        Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id,S.supplier_Name " & _
               " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
               " left join pegasus.dbo.Countries on pegasus.dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId) & _
               " and (len(supplier_Email1)=0  and len(supplier_Email2)=0  and len(supplier_Email3)=0  and len(supplier_Email4)=0)", conB)
        cmdSelect.CommandType = CommandType.Text

        conB.Open()
        myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        While myReader.Read()
            If texCheck = "" Then
                texCheck = myReader("supplier_Name")
            Else
                texCheck = texCheck & "," & myReader("supplier_Name")
            End If
        End While

        conB.Close()

        '' texCheck = "לא קיים מייל מעודכן במערכת לספק " & "<BR>" & texCheck
        ''texCheck = "<BR>" & texCheck & "<BR>" & "להוספת כתובת מייל תיקנית יש לגשת למסך ניהול הספקים"
        ''  checkSup.Style.Add("visibility", "block")
        If texCheck <> "" Then
            checkSup.Visible = True
        End If
        ''checkSup.


    End Function
    Function GetSuppliers()
        If DepartureId > 0 Then
            'Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id,S.supplier_Name,Departure_Id,Vouchers_Status ,Country_Name " & _
            '   " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
            '   " left join pegasus_test.dbo.Countries on pegasus_test.dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), conB)
            Dim cmdSelect As New SqlClient.SqlCommand("select distinct VS.supplier_Id,S.supplier_Name,Departure_Id,Country_Name " & _
                    " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
                    " left join pegasus.dbo.Countries on pegasus.dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), conB)

            ''  Response.Write(cmdSelect.CommandText)
            cmdSelect.CommandType = CommandType.Text
            conB.Open()

            ''    Dim list1 As New ListItem("הכל", "0")
            fselect.DataTextField = "supplier_Name"
            fselect.DataValueField = "supplier_Id"
            Dim s = cmdSelect.ExecuteReader()

            fselect.DataSource = s
            fselect.DataBind()


            Dim list1 As New ListItem("שלח לכל הספקים", "0")
            fselect.Items.Insert(0, list1)
            conB.Close()
            '  Response.Write(fselect.Items.Count)
            If fselect.Items.Count > 1 Then
                divSelect.Visible = True
                divfselectError.Visible = False
            Else
                divfselectError.Visible = True
                divSelect.Visible = False
            End If

        End If
    End Function
    Function GetLogdata()
        If DepartureId > 0 Then
            Dim cmdSelect As New SqlClient.SqlCommand("select Date, Guaranteed, FIRSTNAME_ENG + ' ' + LASTNAME_ENG as Username  " & _
                            " from RoomingList_Guaranteed_Log RL left join Users on  RL.User_id=Users.User_Id  where Departure_Id = " & CInt(DepartureId) & " order by Date desc ", conB)

            cmdSelect.CommandType = CommandType.Text
            conB.Open()


            GLog.DataSource = cmdSelect.ExecuteReader()
            GLog.DataBind()
            conB.Close()
            If GLog.Items.Count > 0 Then
                GLog.Visible = True
            End If

        End If
    End Function
    Public Sub SaveData()

        Dim cmd As New SqlClient.SqlCommand
        Dim UserId = Trim(Request.Cookies("bizpegasus")("UserId"))
        Dim path As String = Server.MapPath("~/Download/RoomingList/")
        Dim fileOK As Boolean = False

        Dim fileName As String = ""
        If Not (fileupload1.PostedFile Is Nothing) And (fileupload1.PostedFile.ContentLength > 0) Then 'Check to make sure we actually have a file to upload
            Dim fileExtension As String
            fileExtension = System.IO.Path.GetExtension(fileupload1.PostedFile.FileName).ToLower()
            ' Response.Write("fileExtension=" & fileExtension & "<BR>")
            Dim allowedExtensions As String() = _
                {".jpg", ".jpeg", ".png", ".gif", ".pdf", ".docx", ".doc", ".xls", ".xlsx", ".txt"}
            For i As Integer = 0 To allowedExtensions.Length - 1
                ' Response.Write(allowedExtensions(i) & "<BR>")
                If fileExtension = allowedExtensions(i) Then
                    fileOK = True
                End If
            Next
            '   Response.Write("fileOK=" & fileOK)
            If fileOK Then
                Dim originName = Replace(System.IO.Path.GetFileName(fileupload1.PostedFile.FileName), System.IO.Path.GetExtension(fileupload1.PostedFile.FileName), "")
                fileName = originName & "_" & DepartureId & "_" & Day(Now()) & Month(Now()) & Year(Now) & "_" & Hour(Now) & "_" & Minute(Now()) & "_" & Second(Now()) & System.IO.Path.GetExtension(fileupload1.PostedFile.FileName)
                fileupload1.PostedFile.SaveAs(path & _
                     fileName)
            Else
                RegisterStartupScript("scriptName_FreeText", "<script type=""text/javascript"">alert('סוג קובץ לא תקין')</script>")
            End If
        End If
        Dim pGuaranteed As Boolean
        If Request.Form("Guaranteed") = "on" Then
            pGuaranteed = True
        Else
            pGuaranteed = False
        End If

        If Len(Request.Form("fselect")) > 0 Then   'save only one supplier_Id
            If Len(Request.Form("fselect")) > 0 And Request.Form("fselect") <> "0" Then  'save only one supplier_Id
                supplier_Id = Request.Form("fselect")
            End If
            If Len(Request.Form("fselect")) > 0 And Request.Form("fselect") = "0" Then
                supplier_Id = func.GetSuppliersByDepId(CInt(Request("dID")))
            End If
            supplierName = func.GetSelectSupplierName(supplier_Id)
            supplierUserName = func.GetSelectSupplierUserName(supplier_Id)
            supplierEmail = func.GetSelectSupplierUserEmail(supplier_Id)
            'supplierGUID = func.GetSelectSupplierGUID(supplier_Id)
            ''Insert ''
            Dim arrSupplierId As Array

            arrSupplierId = Split(supplier_Id, ",")


            'Response.Write(UBound(arrSupplierId))
            'Response.End()
            Dim j As Integer
            For j = 0 To UBound(arrSupplierId)
                cmdSelect = New SqlClient.SqlCommand("Ins_RoomingListToSuppliers", conB)
                cmdSelect.CommandType = CommandType.StoredProcedure
                ' cmdSelect.Parameters.Add("@RID", SqlDbType.Int).Direction = ParameterDirection.InputOutput
                'cmdSelect.Parameters("@RID").Value = 0
                cmdSelect.Parameters.Add("@RID", CInt(RID)).Direction = ParameterDirection.InputOutput
                cmdSelect.Parameters.Add("@supplier_Id", SqlDbType.Int).Value = CInt(arrSupplierId(j)) 'CInt(Request.Form("fselect"))
                cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
                cmdSelect.Parameters.Add("@FileName", SqlDbType.NVarChar, 200).Value = fileName
                cmdSelect.Parameters.Add("@date", SqlDbType.DateTime).Value = Now()
                cmdSelect.Parameters.Add("@Guaranteed", SqlDbType.Bit).Value = pGuaranteed
                conB.Open()
                cmdSelect.ExecuteNonQuery()
                conB.Close()

                ID = cmdSelect.Parameters("@RID").Value
                cmdSelect.Dispose()

            Next
            '    Response.End()

            '''----
            ''---  Send Mail
            'supplierEmail = "adina@cyberserve.co.il"
            'If Trim(supplierEmail) = "" Then
            '    conBIns.Open()
            '    cmd = New System.Data.SqlClient.SqlCommand("update  RoomingList set Status_MailError=2 where id=@RID", conBIns)
            '    cmd.CommandType = CommandType.Text
            '    cmd.Parameters.Add("@RID", SqlDbType.Int).Value = CInt(ID)
            '    cmd.ExecuteNonQuery()
            '    conBIns.Close()

            'End If
            If Trim(supplierEmail) <> "" Then
                EmailContent = "Dear " & supplierUserName & ".<BR>A new Rooming list regarding group " & DepatureCode & " has been sent to your personal account"
                EmailContent = EmailContent & "<BR>" & "Click here to  <A href=https://www.pegasusisrael.co.il/suppliers>login</a>"
                EmailContent = EmailContent & "<BR>" & "Best Regards " & "<BR>" & "Pegasus operation team" & supplierEmail
                Dim Msg As New System.Web.Mail.MailMessage
                Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                Msg.BodyEncoding = System.Text.Encoding.UTF8
                Msg.From = "info@pegasusisrael.co.il"
                Msg.Subject = "New Rooming List of group" & DepatureCode & " is waiting for you in your Pegasus personal account"
                Msg.To = supplierEmail
                '  Msg.To = "faina@cyberserve.co.il"
                Msg.Body = EmailContent.ToString()

                Try
                    System.Web.Mail.SmtpMail.SmtpServer = ConfigurationSettings.AppSettings.Item("smtp_server")
                    System.Web.Mail.SmtpMail.Send(Msg)
                Catch ex As Exception
                    conBIns.Open()
                    cmd = New System.Data.SqlClient.SqlCommand("update  RoomingList set Status_MailError=1 where id=@RID", conBIns)
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add("@RID", SqlDbType.Int).Value = CInt(ID)
                    cmd.ExecuteNonQuery()
                    conBIns.Close()
                End Try
                Msg = Nothing
            End If

        Else   'save to all suppliers
                ''''''Dim myReader As System.Data.SqlClient.SqlDataReader
                ''''''Dim cmdSelectIns As New SqlClient.SqlCommand
                ''''''Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id ,supplier_Name  from VouchersToSuppliers VS left join Suppliers S on VS.supplier_Id= S.supplier_Id where Departure_Id =  " & CInt(DepartureId), conB)
                ''''''cmdSelect.CommandType = CommandType.Text
                ''''''conB.Open()
                ''''''myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                ''''''While myReader.Read()
                ''''''    If Not IsDBNull(myReader("supplier_Id")) Then
                ''''''        ''Insert ''
                ''''''        If Len(supplierName) > 0 Then
                ''''''            supplierName = supplierName & ", " & myReader("supplier_Name")
                ''''''        Else
                ''''''            supplierName = myReader("supplier_Name")
                ''''''        End If
                ''''''           cmdSelectIns = New SqlClient.SqlCommand("Ins_RoomingListToSuppliers", conBIns)
                ''''''        cmdSelectIns.CommandType = CommandType.StoredProcedure
                ''''''        cmdSelectIns.Parameters.Add("@RID", CInt(RID)).Direction = ParameterDirection.InputOutput
                ''''''        cmdSelectIns.Parameters.Add("@supplier_Id", SqlDbType.Int).Value = CInt(myReader("supplier_Id"))
                ''''''        cmdSelectIns.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                ''''''        cmdSelectIns.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
                ''''''        cmdSelectIns.Parameters.Add("@FileName", SqlDbType.NVarChar, 200).Value = fileName
                ''''''        cmdSelectIns.Parameters.Add("@date", SqlDbType.DateTime).Value = Now()
                ''''''        cmdSelectIns.Parameters.Add("@Guaranteed", SqlDbType.Bit).Value = pGuaranteed


                ''''''        conBIns.Open()
                ''''''        cmdSelectIns.ExecuteNonQuery()
                ''''''        conBIns.Close()
                ''''''        RID = cmdSelectIns.Parameters("@RID").Value
                ''''''        cmdSelectIns = Nothing
                ''''''    End If
                ''''''    supplierUserName = func.GetSelectSupplierUserName(supplier_Id)

                ''''''    supplierEmail = func.GetSelectSupplierUserEmail(supplier_Id)
                ''''''     If Trim(supplierEmail) = "" Then
                ''''''        conBIns.Open()
                ''''''        cmd = New System.Data.SqlClient.SqlCommand("update  RoomingList set Status_MailError=2 where id=@RID", conBIns)
                ''''''        cmd.CommandType = CommandType.Text
                ''''''        cmd.Parameters.Add("@RID", SqlDbType.Int).Value = CInt(ID)
                ''''''        cmd.ExecuteNonQuery()
                ''''''        conBIns.Close()

                ''''''    End If
                ''''''    supplierGUID = func.GetSelectSupplierGUID(supplier_Id)
                ''''''    If Trim(supplierEmail) <> "" Then
                ''''''        EmailContent = "Dear " & supplierUserName & ".<BR>A new Rooming list regarding group " & DepatureCode & " has been sent to your personal account"
                ''''''        EmailContent = EmailContent & "<BR>" & "Click here to  <A href=https://www.pegasusisrael.co.il/suppliers>login</a>"
                ''''''        EmailContent = EmailContent & "<BR>" & "Best Regards " & "<BR>" & "Pegasus operation team" '& supplierEmail
                ''''''        Dim Msg As New System.Web.Mail.MailMessage
                ''''''        Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                ''''''        Msg.BodyEncoding = System.Text.Encoding.UTF8
                ''''''        Msg.From = "info@pegasusisrael.co.il"
                ''''''        Msg.Subject = "New Rooming List of group" & DepatureCode & " is waiting for you in your Pegasus personal account"
                ''''''        Msg.To = supplierEmail
                ''''''        ' Msg.To = "faina@cyberserve.co.il"
                ''''''        Msg.Body = EmailContent.ToString()
                ''''''        Try
                ''''''            System.Web.Mail.SmtpMail.SmtpServer = ConfigurationSettings.AppSettings.Item("smtp_server")
                ''''''            System.Web.Mail.SmtpMail.Send(Msg)
                ''''''        Catch ex As Exception
                ''''''            conBIns.Open()
                ''''''            cmd = New System.Data.SqlClient.SqlCommand("update  RoomingList set Status_MailError=1 where id=@RID", conBIns)
                ''''''            cmd.CommandType = CommandType.Text
                ''''''            cmd.Parameters.Add("@RID", SqlDbType.Int).Value = CInt(ID)
                ''''''            cmd.ExecuteNonQuery()
                ''''''            conBIns.Close()
                ''''''        End Try
                ''''''        Msg = Nothing
                ''''''    End If

                ''''''End While
                ''''''conB.Close()



        End If

        If fileOK Then
            Dim cScript As String
            cScript = "<script language='javascript'>alert('הקובץ הועלה בהצלחה לספקים הבאים:" & supplierName & "');self.close(); </script>"
            RegisterStartupScript("ReloadScrpt", cScript)
        End If
    End Sub

End Class
