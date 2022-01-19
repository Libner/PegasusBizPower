Public Class SendCodebySMS_schedule_withSend
    Inherits System.Web.UI.Page
    Public dr_m As SqlClient.SqlDataReader
    Public UserID As Integer
    Public MOBILE, sms_phone As String
    Public orgName As String
    Public VerificationCode As String
    Dim random As New Random
    Protected datesend, LastDateSend As String
    Protected SMS_content As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim con1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))


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
        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        Dim cmdSelect As New SqlClient.SqlCommand("Select top 1 QDays,QDaysUpdate,LastDateSend from DateSendCode", con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        dr_m = cmdSelect.ExecuteReader()
        While dr_m.Read
            If Not dr_m("QDays") Is DBNull.Value Then
                datesend = dr_m("QDays")
            End If
            If Not dr_m("LastDateSend") Is DBNull.Value Then
                LastDateSend = dr_m("LastDateSend")
            End If
        End While
        dr_m.Close()
        cmdSelect.Dispose()
        con.Close()




        'Response.Write("datesend=" & datesend & "<BR>")
        'Response.Write("LastDateSend=" & LastDateSend & "<BR>")
        'Response.Write("Now=" & Now() & "<BR>" & DateDiff("d", Now(), LastDateSend))
        If DateDiff("d", Now(), LastDateSend) <= datesend Then
            ' Response.Write("Yes")
            'Response.End()
            Dim VerificationCode As Integer
            sms_phone = "036374000"
            Dim cmdSelectUsers As New SqlClient.SqlCommand("SELECT  USER_ID,MOBILE_PRIVATE as MOBILE,FIRSTNAME,LASTNAME FROM   USERS " & _
            "  where ACTIVE=1  and (MOBILE_PRIVATE<>'' or MOBILE_PRIVATE<>null) and User_Bloked=0", con)
            'and (user_id=1160 or user_id=1011 or user_id=1295)

            cmdSelect.CommandType = CommandType.Text
            con.Open()
            dr_m = cmdSelect.ExecuteReader()
            While dr_m.Read

                If Not dr_m("USER_ID") Is DBNull.Value Then
                    UserID = dr_m("USER_ID")
                End If
                If Not dr_m("MOBILE") Is DBNull.Value Then
                    MOBILE = dr_m("MOBILE")
                    '  MOBILE = "0507740302"
                End If

                If Trim(MOBILE) <> "" Then

                    VerificationCode = random.Next(100000)
                    '   Response.Write(VerificationCode)
                    ' Response.End()
                    If Left(MOBILE, 3) = "050" Or Left(MOBILE, 3) = "052" Or Left(MOBILE, 3) = "053" Or Left(MOBILE, 3) = "054" Or Left(MOBILE, 3) = "055" Or Left(MOBILE, 3) = "058" Or Left(MOBILE, 3) = "077" Then
                        Dim str
                        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON;update USERS SET SendVerificationCode=GetDate(),VerificationCode_Confirmed=0,VerificationCode=" & VerificationCode & " WHERE USER_ID=" & UserID, con1)
                        cmdInsert.CommandType = CommandType.Text
                        con1.Open()
                        cmdInsert.ExecuteNonQuery()
                        con1.Close()

                        SMS_content = " קוד האימות שנשלח כעת מאתר פגסוס " & ":" & VerificationCode
                        Dim sendUrl, strResponse, getUrl As String
                        sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php"

                        Dim xmlhttp As Object
                        xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
                        xmlhttp.open("POST", sendUrl, False)
                        ' ---block send sms---
                        xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                        xmlhttp.send("uid=2575&un=pegasus&msglong=" & Server.UrlEncode(SMS_content) & "&charset=iso-8859-8" & _
                     "&from=" & sms_phone & "&post=2&list=" & Server.UrlEncode(MOBILE) & _
                       "&desc=" & Server.UrlEncode(orgName))
                        strResponse = xmlhttp.responseText
                        xmlhttp = Nothing
                        '   ---block send sms---

                        'If InStr(strResponse, "ERROR") > 0 Then
                        '    'error = Mid(strResponse, 6)	
                        '    smsStatusId = 2
                        'ElseIf InStr(strResponse, "OK") > 0 Then
                        '    smsStatusId = 1

                        'End If
                        '                Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
                        '" company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,departure_id,date_send)" & _
                        '" VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & MOBILE & "','" & sms_phone & "','" & SMS_content & "'," & UserID & "," & DepartureId & ",GetDate())", con)
                        '                cmdInsert.CommandType = CommandType.Text
                        '                con.Open()
                        '                cmdInsert.ExecuteNonQuery()
                        '                con.Close()
                        'Dim cmdSelectUp As New SqlClient.SqlCommand("SendSMSFeedBack", con)
                        'cmdSelectUp.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                        'cmdSelectUp.Parameters.Add("@companyID", SqlDbType.Int).Value = CInt(companyID)
                        'cmdSelectUp.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
                        'cmdSelectUp.Parameters.Add("@smsStatusId", SqlDbType.Int).Value = CInt(smsStatusId)
                        'cmdSelectUp.Parameters.Add("@MOBILE", SqlDbType.VarChar, 50).Value = MOBILE
                        'cmdSelectUp.Parameters.Add("@sms_phone", SqlDbType.VarChar, 50).Value = sms_phone
                        'cmdSelectUp.Parameters.Add("@FeedBack_Url", SqlDbType.VarChar, 500).Value = urlF 'FeedBack_Url
                        'cmdSelectUp.Parameters.Add("@FeedBack_content", SqlDbType.VarChar, 500).Value = FeedBack_content
                        'cmdSelectUp.Parameters.Add("@UserID", SqlDbType.Int).Value = CInt(UserID)
                        'cmdSelectUp.CommandType = CommandType.StoredProcedure
                        'con.Open()
                        'cmdSelectUp.ExecuteNonQuery()
                        'con.Close()


                    End If


                    'Response.Write(MOBILE & ":" & smsStatusId & ":" & Len(MOBILE) & "<BR>")
                    '   Response.End()
                    'send ....
                End If
            End While
            dr_m.Close()
            cmdSelect.Dispose()
            con.Close()
            Dim cmdUpd As New SqlClient.SqlCommand("Update DateSendCode set LastDateSend=GetDate()", con)
            cmdUpd.CommandType = CommandType.Text

            con.Open()
            '  Try
            cmdUpd.ExecuteNonQuery()
            con.Close()
        End If
        Dim cScript As String
        cScript = "<script language='javascript'>window.opener.location.reload();self.close(); </script>"
        RegisterStartupScript("ReloadScrpt", cScript)
        ' Response.Write("send- end ")

    End Sub

End Class
