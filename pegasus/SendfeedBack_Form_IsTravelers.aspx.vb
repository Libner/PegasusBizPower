Public Class SendfeedBack_Form_IsTravelers
    Inherits System.Web.UI.Page
    Public dr_m As SqlClient.SqlDataReader
    Public TravelerId, smsStatusId, UserID, ContactId As Integer
    Public Phone, sms_phone, SMStype_id, orgName, DateBegin, TourTitle, GuideName, FeedBack_Url, FeedBack_content As String
    Protected DepartureId As Integer = 0
    Protected SMS_content, TravelerNAME As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")


    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim conPegasusU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))


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
        UserID = 1028    'מנהל אתר אינטרנט
        SMStype_id = 8   'משוב טיול
        SMStype_id = 88   ' משוב טיול למטייל
        sms_phone = "036374000"
        'Dim sql = "SET DATEFORMAT dmy;SELECT  Tours_Departures.Departure_Id,Departure_Code,Tours_Departures.Tour_Id,Departure_Date,Departure_Date_End ," & _
        '     " Traveler_Id,TT.Phone,TT.Contact_Id FROM   dbo.Tours_Departures left join  " & BizpowerDBName & ".dbo.Tour_Travelers TT  on TT.Departure_Id=Tours_Departures.Departure_Id  where(IsTravelersTour = 1 And isCanceled = 0)" & _
        '     " and datediff(dd,Departure_Date_End,GetDate())=1 and (TT.Phone<>'' or TT.Phone<>null) and  " & BizpowerDBName & ".dbo.GetIsBitulim(Tours_Departures.Departure_Id,Traveler_Id)=0 "
        'Response.Write(sql)
        'Response.End()
        'Dim s = "SET DATEFORMAT dmy;SELECT  Tours_Departures.Departure_Id,Departure_Code,Tours_Departures.Tour_Id,Departure_Date,Departure_Date_End ," & _
        '  " Traveler_Id,TT.Phone,TT.Contact_Id FROM   dbo.Tours_Departures left join  " & BizpowerDBName & ".dbo.Tour_Travelers TT  on TT.Departure_Id=Tours_Departures.Departure_Id  where(IsTravelersTour = 1 And isCanceled = 0)" & _
        '   " and datediff(dd,Departure_Date_End,GetDate())=1 and (TT.Phone<>'' or TT.Phone<>null) and  " & BizpowerDBName & ".dbo.GetIsBitulim(Tours_Departures.Departure_Id,Traveler_Id)=0 "
        'Response.Write(s)
        '   Response.End()
        'datediff(dd,Departure_Date_End,GetDate())>-3
        Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT  T.FirstName+' '+T.LastName as _Traveler_NAME ,Tours_Departures.Departure_Id,Departure_Code,Tours_Departures.Tour_Id,Departure_Date,Departure_Date_End ," & _
          " TT.Traveler_Id,TT.Phone,TT.Contact_Id FROM   dbo.Tours_Departures left join  " & BizpowerDBName & ".dbo.Tour_Travelers TT  on TT.Departure_Id=Tours_Departures.Departure_Id  " & _
          " left join  " & BizpowerDBName & ".dbo.Travelers T  on T.Traveler_Id=TT.Traveler_Id " & _
          " where(IsTravelersTour = 1 And isCanceled = 0)" & _
          " and datediff(dd,Departure_Date_End,GetDate())=1 and (TT.Phone<>'' or TT.Phone<>null) and  " & BizpowerDBName & ".dbo.GetIsBitulim(Tours_Departures.Departure_Id,TT.Contact_Id)=0 ", conPegasus)
        cmdSelect.CommandType = CommandType.Text
        conPegasus.Open()
        dr_m = cmdSelect.ExecuteReader()
        While dr_m.Read
            If Not dr_m("Phone") Is DBNull.Value Then
                Phone = dr_m("Phone")
                ' Phone = "0507740302"
            End If
            If Not dr_m("Traveler_Id") Is DBNull.Value Then
                TravelerId = dr_m("Traveler_Id")
            End If
            If Not dr_m("Departure_Id") Is DBNull.Value Then
                DepartureId = dr_m("Departure_Id")
            End If
            If Not dr_m("Contact_Id") Is DBNull.Value Then
                ContactId = dr_m("Contact_Id")
            End If
            If Not dr_m("_Traveler_NAME") Is DBNull.Value Then
                TravelerNAME = dr_m("_Traveler_NAME")
            End If

            If Trim(Phone) <> "" Then
                If Len(Phone) < 10 Or Len(Phone) > 10 Then
                    smsStatusId = 3 ' מספר שגוי
                    Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Travelers (" & _
                    " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,Traveler_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                    " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", con)
                    cmdInsert.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert.ExecuteNonQuery()
                    con.Close()

                    Dim cmdInsert2 As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Contact (" & _
                  " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                  " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", con)
                    cmdInsert2.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert2.ExecuteNonQuery()
                    con.Close()

                ElseIf Left(Phone, 3) <> "050" And Left(Phone, 3) <> "052" And Left(Phone, 3) <> "053" And Left(Phone, 3) <> "054" And Left(Phone, 3) <> "055" And Left(Phone, 3) <> "058" And Left(Phone, 3) <> "077" Then
                    smsStatusId = 3
                    Dim str
                    Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Travelers (" & _
                    " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,Traveler_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                   " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", con)
                    cmdInsert.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert.ExecuteNonQuery()
                    con.Close()
                    Dim cmdInsert2 As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Contact (" & _
                    " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                    " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", con)
                    cmdInsert2.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert2.ExecuteNonQuery()
                    con.Close()
                Else


                    FeedBack_Url = ConfigurationSettings.AppSettings.Item("PegasusFeedbackUrl") & "/feedback/faqIsTraveler.aspx?DepartureId=" & DepartureId & "&TravelerId=" & TravelerId & "&ContactId=" & ContactId
                    FeedBack_Url = "url=" & Server.UrlEncode(FeedBack_Url)
                    Dim xmlhttpShort As Object
                    Dim sendUrlShort, strResponseShort, urlF As String
                    sendUrlShort = "http://tinyurl.com/api-create.php"
                    xmlhttpShort = CreateObject("MSXML2.ServerXMLHTTP")
                    xmlhttpShort.open("POST", sendUrlShort, False)

                    xmlhttpShort.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                    xmlhttpShort.send(FeedBack_Url & "&charset=iso-8859-8")
                    strResponseShort = xmlhttpShort.responseText
                    urlF = strResponseShort

                    xmlhttpShort = Nothing
                    FeedBack_content = TravelerName & ","
                    FeedBack_content = FeedBack_content & Chr(13)
                    FeedBack_content = FeedBack_content & " ברוך שובך ארצה. נשמח לקבלת משוב על הטיול בלחיצה על הקישור :"
                    FeedBack_content = FeedBack_content & urlF
                    FeedBack_content = FeedBack_content & " תודה, פגסוס "
                    SMS_content = FeedBack_content
                    '  Response.Write(FeedBack_content)
                    ' Response.End()

                    smsStatusId = 0
                    Dim sendUrl, strResponse, getUrl As String
                    sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php"

                    Dim xmlhttp As Object
                    xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
                    xmlhttp.open("POST", sendUrl, False)
                    ' ---block send sms---
                    xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                    xmlhttp.send("uid=2575&un=pegasus&msglong=" & Server.UrlEncode(SMS_content) & "&charset=iso-8859-8" & _
                   "&from=" & sms_phone & "&post=2&list=" & Server.UrlEncode(Phone) & _
                   "&desc=" & Server.UrlEncode(orgName))
                    strResponse = xmlhttp.responseText
                    xmlhttp = Nothing

                    '   ---block send sms---

                    If InStr(strResponse, "ERROR") > 0 Then
                        smsStatusId = 2
                    ElseIf InStr(strResponse, "OK") > 0 Then
                        smsStatusId = 1

                    End If
                    Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Travelers (" & _
                    " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,Traveler_CELL,sms_phone,sms_Content,User_id,departure_id,date_send)" & _
                    " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & "," & DepartureId & ",GetDate())", con)
                    cmdInsert.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert.ExecuteNonQuery()
                    con.Close()
                    Dim cmdInsert2 As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Contact (" & _
                      " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                      " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", con)
                    cmdInsert2.CommandType = CommandType.Text
                    con.Open()
                    cmdInsert2.ExecuteNonQuery()
                    con.Close()
                    Dim cmdSelectUp As New SqlClient.SqlCommand("SendSMSFeedBack_IsTravelers", con)
                    cmdSelectUp.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                    cmdSelectUp.Parameters.Add("@TravelerId", SqlDbType.Int).Value = CInt(TravelerId)
                    cmdSelectUp.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
                    cmdSelectUp.Parameters.Add("@smsStatusId", SqlDbType.Int).Value = CInt(smsStatusId)
                    cmdSelectUp.Parameters.Add("@Cellular", SqlDbType.VarChar, 50).Value = Phone
                    cmdSelectUp.Parameters.Add("@sms_phone", SqlDbType.VarChar, 50).Value = sms_phone
                    cmdSelectUp.Parameters.Add("@FeedBack_Url", SqlDbType.VarChar, 500).Value = urlF 'FeedBack_Url
                    cmdSelectUp.Parameters.Add("@FeedBack_content", SqlDbType.VarChar, 500).Value = FeedBack_content
                    cmdSelectUp.Parameters.Add("@UserID", SqlDbType.Int).Value = CInt(UserID)
                    cmdSelectUp.CommandType = CommandType.StoredProcedure
                    con.Open()
                    cmdSelectUp.ExecuteNonQuery()
                    con.Close()


                End If

            End If
            conPegasusU.Open()
            Dim cmdUpd As New SqlClient.SqlCommand("update Tours_Departures set IsTravelersTour=1 where Departure_Id=" & DepartureId, conPegasusU)
            cmdUpd.ExecuteNonQuery()
            conPegasusU.Close()

        End While
        dr_m.Close()
        cmdSelect.Dispose()
        conPegasus.Close()
           '   Response.Write("send- end ")

    End Sub

End Class
