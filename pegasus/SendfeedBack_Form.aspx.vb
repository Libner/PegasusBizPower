Public Class SendfeedBack_Form
    Inherits System.Web.UI.Page
    Public dr_m As SqlClient.SqlDataReader
    Public ContactId, MemberId, smsStatusId, companyID, UserID As Integer
    Public Cellular, sms_phone, SMStype_id, orgName, DateBegin, TourTitle, GuideName, FeedBack_Url, FeedBack_content As String
    Protected DepartureId As Integer = 0
    Protected SMS_content As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))

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
        companyID = 12681
        UserID = 1028    'מנהל אתר אינטרנט
        SMStype_id = 8   'משוב טיול
        sms_phone = "036374000"
        ' Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT  Tours_Departures.Departure_Id,Departure_Code,Tours_Departures.Tour_Id,Departure_Date,Departure_Date_End ,Member_Id,Contact_Id,Cellular FROM   dbo.Tours_Departures " & _
        '" left join Members on Members.Departure_Id=Tours_Departures.Departure_Id" & _
        '" where ((Departure_Date>='01-05-2017' and Departure_Date<=GetDate()) and  Departure_Date_End<=GetDate())and (Cellular<>'' or Cellular<>null) ", conPegasus)
        Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT  Tours_Departures.Departure_Id,Departure_Code,Tours_Departures.Tour_Id,Departure_Date,Departure_Date_End ,Member_Id,Contact_Id,Cellular FROM   dbo.Tours_Departures " & _
  " left join Members on Members.Departure_Id=Tours_Departures.Departure_Id where isCanceled=0 and datediff(dd,Departure_Date_End,GetDate())=1 and (Cellular<>'' or Cellular<>null) and  bizpower_pegasus.dbo.GetIsBitulim(Tours_Departures.Departure_Id,Contact_Id)=0", conPegasus)

        cmdSelect.CommandType = CommandType.Text
        conPegasus.Open()
        dr_m = cmdSelect.ExecuteReader()
        While dr_m.Read

            If Not dr_m("Member_Id") Is DBNull.Value Then
                MemberId = dr_m("Member_Id")
            End If
            If Not dr_m("Cellular") Is DBNull.Value Then
                Cellular = dr_m("Cellular")
                ' Cellular = "0507740302"
            End If
            If Not dr_m("Contact_Id") Is DBNull.Value Then
                ContactId = dr_m("Contact_Id")
            End If

            If Not dr_m("Departure_Id") Is DBNull.Value Then
                DepartureId = dr_m("Departure_Id")
            End If

            If Trim(Cellular) <> "" Then
                If Len(Cellular) < 10 Or Len(Cellular) > 10 Then
                    smsStatusId = 3 ' מספר שגוי
                    '        Dim t = "SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
                    '" company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                    '" VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & Cellular & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate())," & DepartureId
                    '    Response.Write(t)
                    ' Response.End()
                    Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
             " company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
             " VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & Cellular & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", con)

                    cmdInsert.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert.ExecuteNonQuery()
                    con.Close()

                ElseIf Left(Cellular, 3) <> "050" And Left(Cellular, 3) <> "052" And Left(Cellular, 3) <> "053" And Left(Cellular, 3) <> "054" And Left(Cellular, 3) <> "055" And Left(Cellular, 3) <> "058" And Left(Cellular, 3) <> "077" Then
                    smsStatusId = 3
                    Dim str
                    Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
        " company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
        " VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & Cellular & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", con)
                    cmdInsert.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert.ExecuteNonQuery()
                    con.Close()

                Else


                    FeedBack_Url = "https://www.pegasusisrael.co.il/feedback/faq.aspx?DepartureId=" & DepartureId & "&ContactId=" & ContactId
                    FeedBack_Url = "url=" & Server.UrlEncode(FeedBack_Url)
                    Dim xmlhttpShort As Object
                    Dim sendUrlShort, strResponseShort, urlF As String
                    sendUrlShort = "http://tinyurl.com/api-create.php"
                    'sendUrlShort = "http://tiny-url.info/api/v1/create"
                    xmlhttpShort = CreateObject("MSXML2.ServerXMLHTTP")
                    xmlhttpShort.open("POST", sendUrlShort, False)

                    xmlhttpShort.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                    xmlhttpShort.send(FeedBack_Url & "&charset=iso-8859-8")
                    strResponseShort = xmlhttpShort.responseText
                    urlF = strResponseShort

                    xmlhttpShort = Nothing
                    FeedBack_content = " ברוך שובך ארצה. נשמח לקבלת משוב על הטיול בלחיצה על הקישור :"
                    FeedBack_content = FeedBack_content & urlF
                    FeedBack_content = FeedBack_content & " תודה, פגסוס "
                    SMS_content = FeedBack_content

                    smsStatusId = 0
                    Dim sendUrl, strResponse, getUrl As String
                    sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php"

                    Dim xmlhttp As Object
                    xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
                    xmlhttp.open("POST", sendUrl, False)
                    ' ---block send sms---
                    xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                    xmlhttp.send("uid=2575&un=pegasus&msglong=" & Server.UrlEncode(SMS_content) & "&charset=iso-8859-8" & _
                 "&from=" & sms_phone & "&post=2&list=" & Server.UrlEncode(Cellular) & _
                   "&desc=" & Server.UrlEncode(orgName))
                    strResponse = xmlhttp.responseText
                    xmlhttp = Nothing
                    '   ---block send sms---

                    If InStr(strResponse, "ERROR") > 0 Then
                        'error = Mid(strResponse, 6)	
                        smsStatusId = 2
                    ElseIf InStr(strResponse, "OK") > 0 Then
                        smsStatusId = 1

                    End If
                    Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
" company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,departure_id,date_send)" & _
" VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & Cellular & "','" & sms_phone & "','" & SMS_content & "'," & UserID & "," & DepartureId & ",GetDate())", con)
                    cmdInsert.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert.ExecuteNonQuery()
                    con.Close()
                    Dim cmdSelectUp As New SqlClient.SqlCommand("SendSMSFeedBack", con)
                    cmdSelectUp.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                    cmdSelectUp.Parameters.Add("@companyID", SqlDbType.Int).Value = CInt(companyID)
                    cmdSelectUp.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
                    cmdSelectUp.Parameters.Add("@smsStatusId", SqlDbType.Int).Value = CInt(smsStatusId)
                    cmdSelectUp.Parameters.Add("@Cellular", SqlDbType.VarChar, 50).Value = Cellular
                    cmdSelectUp.Parameters.Add("@sms_phone", SqlDbType.VarChar, 50).Value = sms_phone
                    cmdSelectUp.Parameters.Add("@FeedBack_Url", SqlDbType.VarChar, 500).Value = urlF 'FeedBack_Url
                    cmdSelectUp.Parameters.Add("@FeedBack_content", SqlDbType.VarChar, 500).Value = FeedBack_content
                    cmdSelectUp.Parameters.Add("@UserID", SqlDbType.Int).Value = CInt(UserID)
                    cmdSelectUp.CommandType = CommandType.StoredProcedure
                    con.Open()
                    cmdSelectUp.ExecuteNonQuery()
                    con.Close()


                End If


                'Response.Write(Cellular & ":" & smsStatusId & ":" & Len(Cellular) & "<BR>")
                '   Response.End()
                'send ....
            End If
        End While
        dr_m.Close()
        cmdSelect.Dispose()
        conPegasus.Close()

        '   Response.Write("send- end ")
       
    End Sub

End Class
