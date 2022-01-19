Imports System.Net
Imports System.Xml
Imports System.IO
Imports System.Web.Services.Protocols
Public Class getKishuritHandly
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Public funcB As New bizpower.cfunc
    Public fName, DepartureCode, fTitle, fValue As String
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents btnSubmit As UI.WebControls.Button
    Protected WithEvents ltMess As Literal
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
        ltMess.Text = ""

        If Not Page.IsPostBack Then

            fValue = Now()
        Else

            Session.LCID = 1037
            If Trim(Request("fromDate")) = "" Then
                fromDate = Now.AddDays(0).AddHours(-1)
            Else
                fromDate = CDate(Request("fromDate"))
            End If
            If Trim(Request("toDate")) = "" Then
                toDate = Now.AddDays(0).AddHours(0)
            Else
                toDate = CDate(Request("toDate"))
            End If
            'Response.Write(toDate.Month & "/" & toDate.Day & "/" & toDate.Year)
            'Response.Write(toDate.Month & "/" & toDate.Day & "/" & toDate.Year)
            'Response.End()
            Response.Redirect("getmessages.aspx?fromDate=" & fromDate.Month & "/" & fromDate.Day & "/" & fromDate.Year & " 00:01&toDate=" & toDate.Month & "/" & toDate.Day & "/" & toDate.Year & " 23:59&handly=1")
            ''  Response.Write(fromDate)
            'Dim resultStr As String
            'Dim isErr As Boolean
            'isErr = False
            'Try
            '    'update entrance count

            '    runWS()
            'Catch exs As SoapException
            '    resultStr += "Fault Code Namespace: " & exs.Code.Namespace
            '    resultStr += "<br>Fault Code Name: " & exs.Code.Name
            '    resultStr += "<br>SOAP Actor that threw Exception: " & exs.Actor
            '    resultStr += "<br>Error Message: " & exs.Message
            '    resultStr += "<br>Detail.OuterXml: " & HttpUtility.HtmlEncode(exs.Detail.OuterXml)
            '    isErr = True
            'Catch exi As System.IO.IOException
            '    resultStr += "exi.InnerException=" & exi.InnerException.Message
            '    isErr = True
            'Catch ex As Exception
            '    resultStr += " Err.Description: " & Err.Description
            '    isErr = True
            'End Try

            ''Response.Write(resultStr)
            'If isErr = True Then
            '    'Send an email to programmer when site wide error occurs
            '    'Response.Write(resultStr)
            '    'Response.End()
            '    Dim Mailer As Mail.MailMessage
            '    Dim smtp As Mail.SmtpMail
            '    Mailer = New Mail.MailMessage

            '    Mailer.From = "mila@cyberserve.co.il"
            '    Mailer.To = "mila@cyberserve.co.il"
            '    'Mailer.Bcc = "elad@kishurit.co.il"

            '    Mailer.BodyEncoding = System.Text.Encoding.GetEncoding("windows-1255")
            '    Mailer.BodyFormat = Mail.MailFormat.Text
            '    'If isErr = True Then
            '    Mailer.Subject = "Pegasus Bizpower: getmessages - ERROR"
            '    Mailer.Body = "Pegasus Bizpower: getmessages - ERROR" & vbCrLf & resultStr
            '    'Else
            '    '    Mailer.Subject = "Pegasus Bizpower: getmessages - NO ERROR"
            '    '    Mailer.Body = ""
            '    ' End If
            '    smtp.Send(Mailer)
            '    smtp = Nothing
            'Else

            '    Session.LCID = 1037
            '    'update entrance count
            '    Dim dateM As String
            '    Dim P16719Kishurit_Sales, P16719Kishurit_Service, P16724ContactUs_Sales, P16724ContactUs_Service, P16504 As String
            '    Dim pYear, pMonth As String

            '    Dim startDate, endDate As Date
            '    '  pYear = Request.QueryString("y")
            '    '  pMonth = Request.QueryString("p")
            '    startDate = fromDate
            '    endDate = toDate
            '    While DateDiff(DateInterval.Month, startDate, endDate) >= 0
            '        Try
            '            pYear = startDate.Year
            '            pMonth = startDate.Month

            '            'Response.Write(startDate & "<bR>")
            '            'Response.Write(endDate & "<bR>")
            '            Dim sql As String

            '            'Put user code to initialize the page here
            '            Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            '            Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            '            If IsNumeric(pYear) Then
            '                sql = "SET DATEFORMAT dmy;SELECT  DateKey from DimDate  where year(DateKey)=" & pYear.ToString & " and month(DateKey)=" & pMonth.ToString
            '            Else
            '                sql = "SET DATEFORMAT dmy;SELECT  DateKey from DimDate  where (DateDiff(dd, DateKey, GetDate()) = 1)"
            '            End If
            '            Dim cmdSelect As New SqlClient.SqlCommand(sql, con)
            '            'Response.Write(sql & "<bR>")
            '            cmdSelect.CommandType = CommandType.Text
            '            con.Open()
            '            Dim dr_m As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
            '            Dim urlF
            '            While dr_m.Read
            '                dateM = dr_m("DateKey")
            '                'Response.Write("d=" & dateM & "<bR>")
            '                ' Response.End()
            '                'appeal_CallStatus= 1 'פניית שירות
            '                'else פניית מכירה
            '                'Dim questionsid, statusId



            '                '     questionsid = 16719
            '                '  statusId = 1

            '                Dim sqls As String

            '                ' Try
            '                P16719Kishurit_Sales = funcB.GetSalesCount(dateM, "16719", "0")  'הודעות קישורית
            '                P16719Kishurit_Service = funcB.GetSalesCount(dateM, "16719", "1")
            '                Dim Upd As New SqlClient.SqlCommand("SET DATEFORMAT dmy;UPDATE DimDate SET  P16719Kishurit_Sales=@P16719Kishurit_Sales,P16719Kishurit_Service=@P16719Kishurit_Service where DateKey=@dateM", conU)
            '                Upd.Parameters.Add("@P16719Kishurit_Sales", SqlDbType.Int).Value = P16719Kishurit_Sales
            '                Upd.Parameters.Add("@P16719Kishurit_Service", SqlDbType.Int).Value = P16719Kishurit_Service
            '                Upd.Parameters.Add("@dateM", SqlDbType.VarChar).Value = dateM
            '                conU.Open()
            '                Upd.CommandType = CommandType.Text
            '                Upd.ExecuteNonQuery()
            '                conU.Close()

            '            End While
            '            dr_m.Close()
            '            cmdSelect.Dispose()
            '            con.Close()


            '        Catch exs As SoapException
            '            resultStr += "Fault Code Namespace: " & exs.Code.Namespace
            '            resultStr += "<br>Fault Code Name: " & exs.Code.Name
            '            resultStr += "<br>SOAP Actor that threw Exception: " & exs.Actor
            '            resultStr += "<br>Error Message: " & exs.Message
            '            resultStr += "<br>Detail.OuterXml: " & HttpUtility.HtmlEncode(exs.Detail.OuterXml)
            '            isErr = True
            '        Catch exi As System.IO.IOException
            '            resultStr += "exi.InnerException=" & exi.InnerException.Message
            '            isErr = True
            '        Catch ex As Exception
            '            resultStr += " Err.Description: " & Err.Description
            '            isErr = True
            '        End Try

            '        Response.Write(resultStr)
            '        startDate = DateAdd(DateInterval.Month, 1, startDate)
            '    End While
            'End If
        End If
    End Sub


    Private Function runWS()

        Session.LCID = 1033
        Dim kishurit_ws As New kishurit.Service

        Dim numRows As Integer
        numRows = 0
        'Create the network credentials and assign them to the service credentials
        Dim netCredential As New NetworkCredential("4383", "4444")
        Dim uri As New Uri(kishurit_ws.Url)
        Dim credentials As ICredentials = netCredential.GetCredential(uri, "Basic")
        kishurit_ws.Credentials = credentials

        'Be sure to set PreAuthenticate to true or else  authentication will not be sent.
        kishurit_ws.PreAuthenticate = True

        Dim strXML As String = kishurit_ws.GetMessagesByDateText("4383", "4444", fromDate, toDate)
        If strXML <> "" Then
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
                numRows = numRows + cmdUpdate.ExecuteNonQuery()
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
        End If
        ltMess.Text = numRows & "קישוריות "
    End Function

    Public Function CheckForNull(ByVal strValue As Object) As Object
        If Not strValue Is DBNull.Value Then
            Return Trim(strValue)
        Else
            Return DBNull.Value
        End If
    End Function

End Class
