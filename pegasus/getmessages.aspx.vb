Imports System.Net
Imports System.Xml
Imports System.IO
Imports System.Web.Services.Protocols

Public Class getmessages
    Inherits System.Web.UI.Page
    Protected fromDate, toDate As Date

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
        Server.ScriptTimeout = 6000
        Session.LCID = 1033

        If Trim(Request.QueryString("fromDate")) = "" Then
            fromDate = Now.AddDays(0).AddHours(-1)
        Else
            fromDate = CDate(Request.QueryString("fromDate"))
        End If
          If Trim(Request.QueryString("toDate")) = "" Then
            toDate = Now.AddDays(0).AddHours(0)
        Else
            toDate = CDate(Request.QueryString("toDate"))
        End If
        '  Response.Write(fromDate)
        Dim resultStr As String
        Dim isErr As Boolean
        isErr = False
        runWS()
        Try
            'update entrance count
        Catch exs As SoapException
            resultStr += "Fault Code Namespace: " & exs.Code.Namespace
            resultStr += "<br>Fault Code Name: " & exs.Code.Name
            resultStr += "<br>SOAP Actor that threw Exception: " & exs.Actor
            resultStr += "<br>Error Message: " & exs.Message
            resultStr += "<br>Detail.OuterXml: " & HttpUtility.HtmlEncode(exs.Detail.OuterXml)
            isErr = True
        Catch exi As System.IO.IOException
            resultStr += "exi.InnerException=" & exi.InnerException.Message
            isErr = True
        Catch ex As Exception
            resultStr += " Err.Description: " & Err.Description
            isErr = True
        End Try

        If isErr = True Then
            'Send an email to programmer when site wide error occurs
            'Response.Write(resultStr)
            'Response.End()
            Dim Mailer As Mail.MailMessage
            Dim smtp As Mail.SmtpMail
            Mailer = New Mail.MailMessage

            Mailer.From = "furfaina@gmail.com"
            Mailer.To = "faina@cyberserve.co.il"
            Mailer.Bcc = "elad@kishurit.co.il"

            Mailer.BodyEncoding = System.Text.Encoding.GetEncoding("windows-1255")
            Mailer.BodyFormat = Mail.MailFormat.Text
            'If isErr = True Then
            Mailer.Subject = "Pegasus Bizpower: getmessages - ERROR"
            Mailer.Body = "Pegasus Bizpower: getmessages - ERROR" & vbCrLf & resultStr
            'Else
            '    Mailer.Subject = "Pegasus Bizpower: getmessages - NO ERROR"
            '    Mailer.Body = ""
            ' End If
            smtp.Send(Mailer)
            smtp = Nothing

        End If
    End Sub

    Private Function runWS()

        Dim kishurit_ws As New kishurit.Service

        'Create the network credentials and assign them to the service credentials
        Dim netCredential As New NetworkCredential("4383", "4444")
        Dim uri As New Uri(kishurit_ws.Url)
        Dim credentials As ICredentials = netCredential.GetCredential(uri, "Basic")
        kishurit_ws.Credentials = credentials

        'Be sure to set PreAuthenticate to true or else  authentication will not be sent.
        kishurit_ws.PreAuthenticate = True

        Dim strXML As String = kishurit_ws.GetMessagesByDateText("4383", "4444", fromDate, toDate)
        Dim XMLreader As StringReader = New System.IO.StringReader(strXML)

        Dim ds As DataSet = New DataSet
        ds.ReadXml(XMLreader)

        Dim dt As New DataTable
        dt = ds.Tables(0)
        Dim rowCount As Integer = dt.Rows.Count - 1

        Dim SqlConn As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
        Dim cmdDelete As New SqlClient.SqlCommand("DELETE FROM pegasus_messages;", SqlConn)
        SqlConn.Open()
        cmdDelete.ExecuteNonQuery()
        cmdDelete.Dispose()
        SqlConn.Close()
        ' Response.Write("rowCount=" & rowCount)
        If False Then ' rowCount = 0 Then
            Dim Mailer As Mail.MailMessage
            Dim smtp As Mail.SmtpMail
            Mailer = New Mail.MailMessage

            Mailer.From = "furfaina@gmail.com"
            Mailer.To = "faina@cyberserve.co.il"
            Mailer.Bcc = "elad@kishurit.co.il;dev@kishurit.co.il"
            Mailer.BodyEncoding = System.Text.Encoding.GetEncoding("windows-1255")
            Mailer.BodyFormat = Mail.MailFormat.Text
            'If isErr = True Then
            Mailer.Subject = "Pegasus Bizpower: getmessages - ERROR"
            Mailer.Body = "Pegasus Bizpower: getmessages - xml rows numbers-" & vbCrLf & rowCount
            'Else
            '    Mailer.Subject = "Pegasus Bizpower: getmessages - NO ERROR"
            '    Mailer.Body = ""
            ' End If
            smtp.Send(Mailer)
            smtp = Nothing
        End If
        For rr As Integer = 0 To rowCount
            SqlConn = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
            Dim cmdUpdate As New SqlClient.SqlCommand("messages_insert_message", SqlConn)
            cmdUpdate.CommandType = CommandType.StoredProcedure

            cmdUpdate.Parameters.Add("@MessageCode", SqlDbType.Char, 25).Value = CheckForNull(dt.Rows(rr)("MessageCode"))
            cmdUpdate.Parameters.Add("@MessageDate", SqlDbType.Char, 25).Value = CheckForNull(dt.Rows(rr)("MessageDate"))
            cmdUpdate.Parameters.Add("@MessageTime", SqlDbType.Char, 25).Value = CheckForNull(dt.Rows(rr)("MessageTime"))
            If dt.Columns.IndexOf("for") >= 0 Then
                cmdUpdate.Parameters.Add("@for", SqlDbType.VarChar, 50).Value = CheckForNull(dt.Rows(rr)("for"))
            Else
                cmdUpdate.Parameters.Add("@for", SqlDbType.VarChar, 50).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("fn") >= 0 Then
                cmdUpdate.Parameters.Add("@fn", SqlDbType.VarChar, 50).Value = CheckForNull(dt.Rows(rr)("fn"))
            Else
                cmdUpdate.Parameters.Add("@fn", SqlDbType.VarChar, 50).Value = DBNull.Value
            End If


            If dt.Columns.IndexOf("sn") >= 0 Then
                cmdUpdate.Parameters.Add("@sn", SqlDbType.VarChar, 50).Value = CheckForNull(dt.Rows(rr)("sn"))
            Else
                cmdUpdate.Parameters.Add("@sn", SqlDbType.VarChar, 50).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("pn") >= 0 Then
                Dim phone As String = ""

                If Not IsDBNull(dt.Rows(rr)("pn")) Then
                    phone = Trim(dt.Rows(rr)("pn"))
                    If phone.IndexOf("-00000") > 0 Then
                        phone = "000000000"
                    End If
                    phone = Replace(phone, " ", "")
                    phone = Replace(phone, "_", "")
                    phone = Replace(phone, "-", "")
                    phone = Replace(phone, ".", "")
                    phone = Replace(phone, Chr(39), "")
                    phone = Replace(phone, Chr(34), "")
                End If
                cmdUpdate.Parameters.Add("@pn", SqlDbType.VarChar, 50).Value = phone
            Else
                cmdUpdate.Parameters.Add("@pn", SqlDbType.VarChar, 50).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("cp") >= 0 Then
                Dim cellphone As String = ""

                If Not IsDBNull(dt.Rows(rr)("cp")) Then
                    cellphone = Trim(dt.Rows(rr)("cp"))

                    If cellphone.IndexOf("-00000") > 0 Then
                        cellphone = "0000000000"
                    End If

                    cellphone = Replace(cellphone, " ", "")
                    cellphone = Replace(cellphone, "_", "")
                    cellphone = Replace(cellphone, "-", "")
                    cellphone = Replace(cellphone, ".", "")
                    cellphone = Replace(cellphone, Chr(39), "")
                    cellphone = Replace(cellphone, Chr(34), "")

                    cmdUpdate.Parameters.Add("@cp", SqlDbType.VarChar, 50).Value = cellphone
                Else
                    cmdUpdate.Parameters.Add("@cp", SqlDbType.VarChar, 50).Value = DBNull.Value
                End If
            Else
                cmdUpdate.Parameters.Add("@cp", SqlDbType.VarChar, 50).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("msg") >= 0 Then
                cmdUpdate.Parameters.Add("@msg", SqlDbType.VarChar, 200).Value = CheckForNull(dt.Rows(rr)("msg"))
            Else
                cmdUpdate.Parameters.Add("@msg", SqlDbType.VarChar, 200).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("a1000") >= 0 Then
                cmdUpdate.Parameters.Add("@a1000", SqlDbType.VarChar, 1000).Value = CheckForNull(dt.Rows(rr)("a1000"))
            Else
                cmdUpdate.Parameters.Add("@a1000", SqlDbType.VarChar, 1000).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("fxn") >= 0 Then
                cmdUpdate.Parameters.Add("@fxn", SqlDbType.VarChar, 50).Value = CheckForNull(dt.Rows(rr)("fxn"))
            Else
                cmdUpdate.Parameters.Add("@fxn", SqlDbType.VarChar, 50).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("st") >= 0 Then
                cmdUpdate.Parameters.Add("@st", SqlDbType.VarChar, 250).Value = CheckForNull(dt.Rows(rr)("st"))
            Else
                cmdUpdate.Parameters.Add("@st", SqlDbType.VarChar, 250).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("hn") >= 0 Then
                cmdUpdate.Parameters.Add("@hn", SqlDbType.VarChar, 250).Value = CheckForNull(dt.Rows(rr)("hn"))
            Else
                cmdUpdate.Parameters.Add("@hn", SqlDbType.VarChar, 250).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("cty") >= 0 Then
                cmdUpdate.Parameters.Add("@cty", SqlDbType.VarChar, 250).Value = CheckForNull(dt.Rows(rr)("cty"))
            Else
                cmdUpdate.Parameters.Add("@cty", SqlDbType.VarChar, 250).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("zip") >= 0 Then
                cmdUpdate.Parameters.Add("@zip", SqlDbType.VarChar, 250).Value = CheckForNull(dt.Rows(rr)("zip"))
            Else
                cmdUpdate.Parameters.Add("@zip", SqlDbType.VarChar, 250).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("destrp") >= 0 Then
                cmdUpdate.Parameters.Add("@destrp", SqlDbType.VarChar, 250).Value = CheckForNull(dt.Rows(rr)("destrp"))
            Else
                cmdUpdate.Parameters.Add("@destrp", SqlDbType.VarChar, 250).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("howreach") >= 0 Then
                cmdUpdate.Parameters.Add("@howreach", SqlDbType.VarChar, 250).Value = CheckForNull(dt.Rows(rr)("howreach"))
            Else
                cmdUpdate.Parameters.Add("@howreach", SqlDbType.VarChar, 250).Value = DBNull.Value
            End If

            If dt.Columns.IndexOf("emlad1") >= 0 Then
                cmdUpdate.Parameters.Add("@emlad1", SqlDbType.VarChar, 250).Value = CheckForNull(dt.Rows(rr)("emlad1"))
            Else
                cmdUpdate.Parameters.Add("@emlad1", SqlDbType.VarChar, 250).Value = DBNull.Value
            End If
            If dt.Columns.IndexOf("ctyp") >= 0 Then
                cmdUpdate.Parameters.Add("@ctyp", SqlDbType.VarChar, 25).Value = CheckForNull(dt.Rows(rr)("ctyp"))
            Else
                cmdUpdate.Parameters.Add("@ctyp", SqlDbType.VarChar, 25).Value = DBNull.Value
            End If

            SqlConn.Open()
            cmdUpdate.ExecuteNonQuery()
            cmdUpdate.Dispose()
            SqlConn.Close()
        Next

        kishurit_ws = Nothing

        Dim cmdInsert As New SqlClient.SqlCommand("messages_insert_appeals", SqlConn)
        cmdInsert.CommandType = CommandType.StoredProcedure
        SqlConn.Open()
        cmdInsert.ExecuteNonQuery()
        cmdInsert.Dispose()
        SqlConn.Close()

    End Function

    Public Function CheckForNull(ByVal strValue As Object) As Object
        If Not strValue Is DBNull.Value Then
            Return Trim(strValue)
        Else
            Return DBNull.Value
        End If
    End Function

End Class