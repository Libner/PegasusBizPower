Imports System.Data.SqlClient
Public Class gradesByDepId
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Public DepartureId As Integer
    Protected DepartureCode, GuideName, TourName, CountMembers, sortQuery, sortQuery1 As String
    Protected CountFeedBack As Integer
    Protected allTourGrade As Decimal
    Protected WithEvents rptCustomers, rptCustomers2, rptCustomers3 As Repeater
    Public qrystring As String
    Public BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    Public PegasusSiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")
    Protected IsTravelersTour As String


    Protected CountPeopleFeedBack As Integer

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
        If IsNumeric(Request.QueryString("depId")) Then
            DepartureId = CInt(Request.QueryString("depId"))
        Else
            DepartureId = 0
        End If
        sortQuery = ""
        CountPeopleFeedBack = func.NumFeedBackPeople(CInt(DepartureId))
        Dim r As Integer
        qrystring = Request.ServerVariables("QUERY_STRING")
        r = qrystring.IndexOf("sort")
        ' Response.Write("r=" & r)
        '  Response.End()
        If r > 0 Then
            sortQuery1 = Replace(Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)), "&", ",")
            sortQuery1 = Replace(sortQuery1, "=", " ")
            sortQuery1 = Replace(sortQuery1, "sort_6", "TourGrade")
            sortQuery1 = Replace(sortQuery1, "sort_5", "NextMinGrade")

            sortQuery1 = Replace(sortQuery1, "sort_4", "FaqGradeMinValue")
            sortQuery1 = Replace(sortQuery1, "sort_3", "FAQ14")
            sortQuery1 = Replace(sortQuery1, "sort_2", "FAQ13")
            sortQuery1 = Replace(sortQuery1, "sort_1", "CONTACT_NAME")

            sortQuery1 = sortQuery1.Substring(1)

            sortQuery = sortQuery1
            ' Response.Write("sortQuery=" & sortQuery)
            '  Response.End()
        Else
       

            sortQuery = " CONTACT_NAME desc"
        End If
        'Response.Write("DepartureId=" & DepartureId)
        If DepartureId > 0 Then
            Dim cmdSelect = New SqlCommand("SELECT Departure_Code,Tour_Name=case when isnull(Tour_PageTitle,'')='' then Tour_Name else Tour_PageTitle end, dbo.GetDepartureGuide(Departure_Id) as Guide_Name,CountMembers,IsNull(IsTravelersTour,0) as IsTravelersTour  FROM dbo.Tours_Departures TD left join Tours T on TD.Tour_Id=T.Tour_Id  WHERE (Departure_Id = @DepartureId)", conPegasus)
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            conPegasus.Open()
            Dim dr_m As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
            If dr_m.Read() Then
                If Not dr_m("Departure_Code") Is DBNull.Value Then
                    DepartureCode = dr_m("Departure_Code")
                End If
                If Not dr_m("Tour_Name") Is DBNull.Value Then
                    TourName = dr_m("Tour_Name")
                End If

                If Not dr_m("Guide_Name") Is DBNull.Value Then
                    GuideName = dr_m("Guide_Name")
                End If
                If Not dr_m("CountMembers") Is DBNull.Value Then
                    CountMembers = dr_m("CountMembers")
                End If
                If Not dr_m("IsTravelersTour") Is DBNull.Value Then
                    IsTravelersTour = dr_m("IsTravelersTour")
                End If
                '  Response.Write("IsTravelersTour=" & IsTravelersTour)
                '  Response.End()

            End If
            conPegasus.Close()

            cmdSelect = New SqlCommand("select count(Id) from FeedBack_Form  WHERE FeedBack_Status=2 and  Departure_Id = @DepartureId", conPegasus)
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            conPegasus.Open()
            If Not IsDBNull(cmdSelect.ExecuteScalar()) Then
                CountFeedBack = cmdSelect.ExecuteScalar()
            Else
                'CountFeedBack = 1
                CountFeedBack = 0
            End If
            conPegasus.Close()


            '           cmdSelect = New SqlCommand("SELECT  M.Member_Id, M.Contact_Id,dbo.getFeedBackFaq(M.Contact_Id,@DepartureId,14) as FAQ14,dbo.getFeedBackFaq(M.Contact_Id,@DepartureId,13) as FAQ13,C.CONTACT_NAME,M.Cellular,Company_Id FROM Members M left join " & BizpowerDBName &".dbo.CONTACTS C on M.Contact_id=C.Contact_Id   WHERE (Departure_Id = @DepartureId)", conPegasus)
            '''' cmdSelect = New SqlCommand("select F.Id, F.Contact_Id,C.Company_Id,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,14) as FAQ14,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,13) as FAQ13,C.CONTACT_NAME,dbo.[getFeedBackFaqGradeMinValue](F.Contact_Id,@DepartureId) as FaqGradeMinValue from FeedBack_Form F " & _
            '' "    left join " & BizpowerDBName &".dbo.CONTACTS C on F.Contact_id=C.Contact_Id   where FeedBack_Status=2 and Departure_Id = @DepartureId", conPegasus)

            ' Response.Write("@sortQuery=" & sortQuery)
            ' Response.End()

            ''  cmdSelect = New SqlCommand("select F.Id, F.Contact_Id,C.Company_Id,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,14) as FAQ14,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,13) as FAQ13,C.CONTACT_NAME,dbo.[getFeedBackFaqGradeMinValue](F.Contact_Id,@DepartureId) as FaqGradeMinValue,dbo.[getFeedback_GetNextMinGrade](F.Contact_Id,@DepartureId) as NextMinGrade,F.TourGrade,SourceDetection from FeedBack_Form F " & _
            ''    "    left join " & BizpowerDBName &".dbo.CONTACTS C on F.Contact_id=C.Contact_Id   where FeedBack_Status=2 and Departure_Id = @DepartureId and  " & BizpowerDBName &".dbo.GetIsBitulim(Departure_Id,C.Contact_Id)=0 order by " & sortQuery, conPegasus)
            If IsTravelersTour = "1" Then
                'cmdSelect = New SqlCommand("select F.Id, F.Traveler_id,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,14) as FAQ14,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,13) as FAQ13,C.(LastName+ ' '+FirstName) as  Traveler_Name,dbo.[getFeedBackFaqGradeMinValue](F.Contact_Id,@DepartureId) as FaqGradeMinValue,dbo.[getFeedback_GetNextMinGradebyCatId](F.Contact_Id,@DepartureId,dbo.[getFeedBackFaqGradeMinValueId](F.Contact_Id,@DepartureId)) as NextMinGrade,F.TourGrade,SourceDetection from FeedBack_Form F " & _
                '                     "    left join " & BizpowerDBName & ".dbo.Travelers C on F.Traveler_id=C.Traveler_id   where FeedBack_Status=2 and Departure_Id = @DepartureId and  " & BizpowerDBName & ".dbo.GetIsBitulim(Departure_Id,C.Contact_Id)=0 order by " & sortQuery, conPegasus)
                'Dim s = "select F.Id, F.Traveler_id,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,14) as FAQ14,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,13) as FAQ13,C.(LastName+ ' '+FirstName) as  Traveler_Name,dbo.[getFeedBackFaqGradeMinValue](F.Contact_Id,@DepartureId) as FaqGradeMinValue,dbo.[getFeedback_GetNextMinGradebyCatId](F.Contact_Id,@DepartureId,dbo.[getFeedBackFaqGradeMinValueId](F.Contact_Id,@DepartureId)) as NextMinGrade,F.TourGrade,SourceDetection from FeedBack_Form F " & _
                '                                     "    left join " & BizpowerDBName & ".dbo.Travelers C on F.Traveler_id=C.Traveler_id   where FeedBack_Status=2 and Departure_Id = @DepartureId and  " & BizpowerDBName & ".dbo.GetIsBitulim(Departure_Id,C.Contact_Id)=0 "

                'Response.Write(s)
                'Response.End()

                cmdSelect = New SqlCommand("select  F.Traveler_id,F.Contact_Id,12681 as Company_Id,dbo.getFeedBackFaq_isTraveler(F.Contact_Id,F.Traveler_Id,@DepartureId,14) as FAQ14,dbo.getFeedBackFaq_isTraveler(F.Contact_Id,F.Traveler_Id,@DepartureId,13) as FAQ13,(C.LastName+ ' '+C.FirstName) as CONTACT_NAME, " & _
                " dbo.[getFeedBackFaqGradeMinValue](F.Contact_Id,@DepartureId) as FaqGradeMinValue, dbo.[getFeedback_GetNextMinGradebyCatId](F.Contact_Id,@DepartureId,dbo.[getFeedBackFaqGradeMinValueId](F.Contact_Id,@DepartureId)) as NextMinGrade, " & _
                " F.TourGrade,SourceDetection from FeedBack_Form F left join " & BizpowerDBName & ".dbo.Travelers C on F.Traveler_id=C.Traveler_id " & _
                " left join " & BizpowerDBName & ".dbo.Contacts on " & BizpowerDBName & ".dbo.Contacts.Contact_Id=F.Contact_Id  where(FeedBack_Status = 2) and Departure_Id = @DepartureId and " & BizpowerDBName & ".dbo.GetIsBitulim(@DepartureId,F.Contact_Id)=0 order by " & sortQuery, conPegasus)
            Else

                cmdSelect = New SqlCommand("select F.Id,'' as Traveler_id, F.Contact_Id,C.Company_Id,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,14) as FAQ14,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,13) as FAQ13,C.CONTACT_NAME,dbo.[getFeedBackFaqGradeMinValue](F.Contact_Id,@DepartureId) as FaqGradeMinValue,dbo.[getFeedback_GetNextMinGradebyCatId](F.Contact_Id,@DepartureId,dbo.[getFeedBackFaqGradeMinValueId](F.Contact_Id,@DepartureId)) as NextMinGrade,F.TourGrade,SourceDetection from FeedBack_Form F " & _
                                     "    left join " & BizpowerDBName & ".dbo.CONTACTS C on F.Contact_id=C.Contact_Id   where FeedBack_Status=2 and Departure_Id = @DepartureId and  " & BizpowerDBName & ".dbo.GetIsBitulim(Departure_Id,C.Contact_Id)=0 order by " & sortQuery, conPegasus)
            End If

            'Response.Write("<br>rptCustomers:" & cmdSelect.commandtext)
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)

            conPegasus.Open()
            Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.HasRows Then
                '  Response.Write(PageSize.SelectedValue)
                '  Response.End()
                rptCustomers.DataSource = dr
                rptCustomers.DataBind()

            End If
            conPegasus.Close()
            '----------

            'cmdSelect = New SqlCommand("select S.company_id,S.contact_id,C.CONTACT_NAME from " & BizpowerDBName &".dbo.sms_FeedBackTo_contact S " & _
            '" left join " & BizpowerDBName &".dbo.CONTACTS C on S.Contact_id=C.Contact_Id where  NOT EXISTS  (Select * from pegasus.dbo.FeedBack_Form where S.contact_id=pegasus.dbo.FeedBack_Form.contact_id) " & _
            '" and  Departure_Id = @DepartureId", conPegasus)



            '  cmdSelect = New SqlCommand("select  OpenDevice,M.Member_Id, M.Contact_Id,BC.CONTACT_NAME,date_send,IsNull(SendCount,0) as SendCount,BC.Company_Id from Members M " & _
            '  " left join  " & BizpowerDBName &".dbo.Contacts BC on BC.Contact_Id=M.Contact_Id " & _
            '  " left join " & BizpowerDBName &".dbo.sms_FeedBackTo_contact C on C.Contact_Id=M.Contact_Id and M.departure_id=C.departure_id " & _
            '  " where  (NOT EXISTS  (Select * from pegasus.dbo.FeedBack_Form where M.contact_id=pegasus.dbo.FeedBack_Form.contact_id and Departure_Id = @DepartureId) or pegasus.dbo.getFeedBackFaqMax(C.Contact_Id,C.departure_id) is null) " & _
            '  " and  M.Departure_Id = @DepartureId and   " & BizpowerDBName &".dbo.GetIsBitulim(M.Departure_Id,M.Contact_Id)=0", conPegasus)
            Dim SQLResult As String

            Dim cmdSelectSms = New SqlCommand("select count(*) as SendSMSCount from sms_FeedBackTo_contact where Departure_Id = @DepartureId", con)
            'cmdSelectSms.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            'con.Open()
            'Dim SendSMSCount As Integer = cmdSelectSms.ExecuteScalar()
            'con.Close()
            ''If SendSMSCount > 0 Then
            If IsTravelersTour = 1 Then
                'SQLResult = "select  OpenDevice,M.Member_Id, M.Contact_Id,BC.CONTACT_NAME,date_send,IsNull(SendCount,0) as SendCount,BC.Company_Id from Members M " & _
                '                      " left join  " & BizpowerDBName & ".dbo.Contacts BC on BC.Contact_Id=M.Contact_Id " & _
                '                      " left join " & BizpowerDBName & ".dbo.sms_FeedBackTo_contact C on C.Contact_Id=M.Contact_Id and M.departure_id=C.departure_id " & _
                '                      " where  (NOT EXISTS  (Select * from " & PegasusSiteDBName & ".dbo.FeedBack_Form where M.contact_id=" & PegasusSiteDBName & ".dbo.FeedBack_Form.contact_id and Departure_Id = @DepartureId) or " & PegasusSiteDBName & ".dbo.getFeedBackFaqMax(C.Contact_Id,C.departure_id) is null) " & _
                '                      " and  M.Departure_Id = @DepartureId and    " & BizpowerDBName & ".dbo.GetIsBitulim(M.Departure_Id,M.Contact_Id)=0"

                SQLResult = "select  12681 as Company_Id,OpenDevice,TT.Contact_Id,TT.Traveler_Id,T.FirstName+' '+T.LastName as CONTACT_NAME," & _
                            " date_send,IsNull(SendCount,0) as SendCount from  " & BizpowerDBName & ".dbo.Tour_Travelers TT" & _
                            " left join " & BizpowerDBName & ".dbo.Travelers T on T.Traveler_id=TT.Traveler_Id" & _
                            " left join " & BizpowerDBName & ".dbo.sms_FeedBackTo_Travelers C on C.Contact_Id=TT.Contact_Id and C.departure_id=TT.departure_id " & _
                            " and C.Traveler_id=TT.Traveler_id  where  (NOT EXISTS  (Select * from " & PegasusSiteDBName & ".dbo.FeedBack_Form F " & _
                            " where TT.contact_id=F.contact_id  and  TT.Traveler_id=F.Traveler_id  and Departure_Id = @DepartureId) or " & PegasusSiteDBName & ".dbo.getFeedBackFaqMax_IsTraveler(C.Contact_Id,C.Traveler_Id,C.departure_id) is null) " & _
                            " and  TT.Departure_Id = @DepartureId  and  " & BizpowerDBName & ".dbo.GetIsBitulim(TT.Departure_Id,TT.Contact_Id)=0"
                '    Response.Write(SQLResult)
                '  Response.End()

            Else
                SQLResult = "select  '' as Traveler_Id,OpenDevice,M.Member_Id, M.Contact_Id,BC.CONTACT_NAME,date_send,IsNull(SendCount,0) as SendCount,BC.Company_Id from Members M " & _
                              " left join  " & BizpowerDBName & ".dbo.Contacts BC on BC.Contact_Id=M.Contact_Id " & _
                              " left join " & BizpowerDBName & ".dbo.sms_FeedBackTo_contact C on C.Contact_Id=M.Contact_Id and M.departure_id=C.departure_id " & _
                              " where  (NOT EXISTS  (Select * from " & PegasusSiteDBName & ".dbo.FeedBack_Form where M.contact_id=" & PegasusSiteDBName & ".dbo.FeedBack_Form.contact_id and Departure_Id = @DepartureId) or " & PegasusSiteDBName & ".dbo.getFeedBackFaqMax(C.Contact_Id,C.departure_id) is null) " & _
                              " and  M.Departure_Id = @DepartureId and    " & BizpowerDBName & ".dbo.GetIsBitulim(M.Departure_Id,M.Contact_Id)=0"




            End If
            'Response.Write("<br>rptCustomers2:" & SQLResult)
            cmdSelect = New SqlCommand(SQLResult, conPegasus)

            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            conPegasus.Open()
            Dim dr2 As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

            If dr2.HasRows Then
                '  Response.Write(PageSize.SelectedValue)
                '  Response.End()
                rptCustomers2.DataSource = dr2
                rptCustomers2.DataBind()
            End If
            conPegasus.Close()


            If IsTravelersTour = 1 Then
                cmdSelect = New SqlCommand("select  F.Traveler_id,12681 as Company_Id,OpenDevice, F.Contact_Id,(C.LastName+ ' '+C.FirstName) as CONTACT_NAME,SendCount, " & _
                 " dbo.getFeedBackFaqMax_IsTraveler(F.Contact_Id,F.Traveler_Id,F.departure_id) as LastFAQ,SourceDetection from FeedBack_Form F 	left join " & BizpowerDBName & ".dbo.Travelers C on F.Traveler_id=C.Traveler_id " & _
                 " left join " & BizpowerDBName & ".dbo.CONTACTS  on F.Contact_id=CONTACTS.Contact_Id  left join " & BizpowerDBName & ".dbo.sms_FeedBackTo_Travelers S on S.Traveler_id=F.Traveler_id and S.departure_id=F.departure_id " & _
                 " where FeedBack_Status = 1 And F.Departure_Id = @DepartureId and " & BizpowerDBName & ".dbo.GetIsBitulim(F.Departure_Id,F.Contact_Id)=0 " & _
                 "  and  dbo.getFeedBackFaqMax_IsTraveler(F.Contact_Id,F.Traveler_Id,F.departure_id)>0", conPegasus)
              
            Else
                cmdSelect = New SqlCommand("select OpenDevice,'' as Traveler_id,  F.Contact_Id,C.Company_Id,C.CONTACT_NAME,SendCount," & PegasusSiteDBName & ".dbo.getFeedBackFaqMax(F.Contact_Id,F.departure_id) as LastFAQ,SourceDetection from " & PegasusSiteDBName & ".dbo.FeedBack_Form F  " & _
                                       "  left join " & BizpowerDBName & ".dbo.CONTACTS C on F.Contact_id=C.Contact_Id " & _
                                      "  left join " & BizpowerDBName & ".dbo.sms_FeedBackTo_contact S on S.Contact_Id=F.Contact_Id and S.departure_id=F.departure_id " & _
                                      "     where FeedBack_Status = 1 And F.Departure_Id = @DepartureId and " & BizpowerDBName & ".dbo.GetIsBitulim(F.Departure_Id,F.Contact_Id)=0 and " & PegasusSiteDBName & ".dbo.getFeedBackFaqMax(F.Contact_Id,F.departure_id)>0", conPegasus)



            End If
            'Response.Write("<br>rptCustomers3:" & cmdSelect.commandtext)
            conPegasus.Open()
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            Dim dr3 As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

            If dr3.HasRows Then
                '  Response.Write(PageSize.SelectedValue)
                '  Response.End()
                rptCustomers3.DataSource = dr3
                rptCustomers3.DataBind()
            End If
            conPegasus.Close()


        End If

    End Sub
    Private Sub rptCustomers3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptCustomers3.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim cmdSelect As New SqlCommand
            Dim ContactId, TravelerId As Integer
            Dim SendCount As Integer
            Dim con1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
              Dim lblCountSendAll As Label
             lblCountSendAll = e.Item.FindControl("lblCountSendAll")
            Dim MailSend As String
            ContactId = e.Item.DataItem("Contact_Id")

            If Not e.Item.DataItem("SendCount") Is DBNull.Value Then
                SendCount = e.Item.DataItem("SendCount")
            End If
            ' Response.Write("SendCount=" & SendCount)

            If IsTravelersTour = 1 Then
                TravelerId = e.Item.DataItem("Traveler_Id")
                cmdSelect = New SqlCommand("Select count(*) as C from Email_to_Traveler where Traveler_id=@TravelerId and departure_id=@DepartureId", con1)
            Else
                cmdSelect = New SqlCommand("Select count(*) as C from Email_to_contact where contact_id=@ContactId and departure_id=@DepartureId", con1)

            End If

            Dim CMail As Integer

            cmdSelect.CommandType = CommandType.Text
            cmdSelect.Parameters.Add("@ContactId", SqlDbType.Int).Value = ContactId
            cmdSelect.Parameters.Add("@TravelerId", SqlDbType.Int).Value = TravelerId
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = DepartureId
            con1.Open()
            CMail = cmdSelect.ExecuteScalar
            If CMail > 0 Then
                SendCount = SendCount + CMail
            Else
            End If
            lblCountSendAll.Text = SendCount
            con1.Close()


        End If
    End Sub
    Private Sub rptCustomers2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptCustomers2.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim ContactId, TravelerId As Integer
            Dim SendCount As String
            Dim con1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            Dim lblSendmail As Label
            Dim lblCountSendAll As Label
            lblSendmail = e.Item.FindControl("lblSendmail")
            lblCountSendAll = e.Item.FindControl("lblCountSendAll")
            Dim MailSend As String
            ContactId = e.Item.DataItem("Contact_Id")
            SendCount = e.Item.DataItem("SendCount")
            Dim cmdSelect = New SqlCommand
            If IsTravelersTour = 1 Then
                TravelerId = e.Item.DataItem("Traveler_Id")
                cmdSelect = New SqlCommand("Select count(*) as C from Email_to_Traveler where Traveler_id=@TravelerId and departure_id=@DepartureId", con1)
            Else
                cmdSelect = New SqlCommand("Select count(*) as C from Email_to_contact where contact_id=@ContactId and departure_id=@DepartureId", con1)

            End If

             Dim CMail As Integer

            cmdSelect.CommandType = CommandType.Text
            cmdSelect.Parameters.Add("@ContactId", SqlDbType.Int).Value = ContactId
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = DepartureId
            cmdSelect.Parameters.Add("@TravelerId", SqlDbType.Int).Value = TravelerId
            con1.Open()
            CMail = cmdSelect.ExecuteScalar
            If CMail > 0 Then
                lblSendmail.Text = "כן"
                If IsNumeric(SendCount) Then
                    SendCount = SendCount + CMail
                End If
            Else
                lblSendmail.Text = "לא"
            End If
            lblCountSendAll.Text = SendCount
            con1.Close()


            '''    conPegasus1.Open()
        End If

    End Sub
    Private Sub rptCustomers_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptCustomers.ItemDataBound
        '''Dim conPegasus1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))

        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then


            '''    Dim lblTourGrade As Label
            '''    lblTourGrade = e.Item.FindControl("lblTourGrade")

            '''    Dim ContactId, FormId, ProcW, FQ13, FQ14 As Integer
            '''    Dim avg2, avg3, avg4, avgP, FP13, FP14, TourGrade As Decimal
            Dim TourGrade As Decimal
            TourGrade = e.Item.DataItem("TourGrade")
            '''    ProcW = 100
            '''    Dim cmdSelect As SqlCommand
            '''    FormId = e.Item.DataItem("Id")
            '''    ' ContactId = e.Item.DataItem("Contact_Id")
            '''    '    Response.Write("FormId=" & FormId & "-")
            '''    cmdSelect = New SqlCommand("getFeedBackCat_Percent", conPegasus1)
            '''    cmdSelect.CommandType = CommandType.StoredProcedure
            '''    cmdSelect.Parameters.Add("@FormId", SqlDbType.Int).Value = FormId
            '''    conPegasus1.Open()
            '''    Dim dr_page As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            '''    While dr_page.Read()
            '''        Select Case dr_page("Category_Id")
            '''            Case "2"
            '''                avg2 = avg2 + dr_page("perc")
            '''            Case "3"
            '''                avg3 = avg3 + dr_page("perc")
            '''            Case "4"
            '''                avg4 = avg4 + dr_page("perc")
            '''            Case "5"
            '''                If dr_page("FaqIndex") = 13 Then
            '''                    FQ13 = dr_page("perc")
            '''                    ' Response.Write("FQ13=" & FQ13 & ":" & ContactId)
            '''                End If
            '''                If dr_page("FaqIndex") = 14 Then
            '''                    FQ14 = dr_page("perc")
            '''                    '  Response.Write("FQ14=" & FQ14 & ":" & ContactId)

            '''                End If

            '''        End Select
            '''    End While
            '''    avg2 = Math.Round(avg2 / 3, 0)
            '''    avg3 = Math.Round(avg3 / 3, 0)
            '''    avg4 = Math.Round(avg4 / 4, 0)
            '''    '   Response.Write("=" & avg2 & ":" & avg3 & ":" & avg4 & "<BR>")
            '''    conPegasus1.Close()
            '''    If avg2 < 80 Then
            '''        ProcW = ProcW - 15
            '''        avgP = (avg2 * 15) / 100
            '''    End If
            '''    If avg3 < 80 Then
            '''        ProcW = ProcW - 15
            '''        avgP = avgP + (avg3 * 15) / 100

            '''    End If
            '''    If avg4 < 80 Then
            '''        ProcW = ProcW - 15
            '''        avgP = avgP + (avg4 * 15) / 100
            '''    End If
            '''    If ProcW = 100 Then
            '''        FP13 = (70 * (FQ13 / 100))
            '''        FP14 = (30 * (FQ14 / 100))
            '''    Else
            '''        FP13 = (70 * ProcW * (FQ13 / 100)) / 100
            '''        FP14 = (30 * ProcW * (FQ14 / 100)) / 100
            '''    End If

            '''    TourGrade = FP13 + FP14 + avgP
            '''    ' Response.Write("FormId=" & FormId & "<BR>")
            '''    '  Response.Write("avgP=" & avgP & "<BR>")
            '''    '    Response.Write("FP13=" & FP13 & ":" & FP14 & "<br>")
            '''    '  Response.Write("TourGrade=" & TourGrade & "<BR>------<BR>")
            '''    ' Response.Write("allTourGrade=" & allTourGrade & "<BR>------<BR>")

            '''    'Response.Write("ProcW=" & ProcW & "<BR>------<BR>")

            '''    lblTourGrade.Text = TourGrade & "%"
            '''    allTourGrade = allTourGrade + TourGrade
            allTourGrade = allTourGrade + TourGrade

        End If


    End Sub
End Class
