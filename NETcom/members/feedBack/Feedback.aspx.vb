Imports System.Data.SqlClient

Public Class Feedback
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Public DepartureId, status, ContactId, FrmId As Integer
    Protected DepartureCode, GuideName, TourName, CONTACTNAME, FAQ14, FAQ13, FaqGradeMinValue, LastFAQ, SendCount, OpenDevice As String
    Protected plhStatus2 As PlaceHolder
    Protected WithEvents rptCat As Repeater
    Protected WithEvents rptFaq As Repeater
    Protected lblValue, lblTitle, lblFAQID As Label
    Protected TourGrade, GuideGrade As Decimal
    Public BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    Public PegasusSiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")






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
        If IsNumeric(Request.QueryString("ContactId")) Then
            ContactId = CInt(Request.QueryString("ContactId"))
        Else
            ContactId = 0
        End If

        If IsNumeric(Request.QueryString("status")) Then
            status = CInt(Request.QueryString("status"))
        Else
            status = 0
        End If

        If DepartureId <> 0 And ContactId <> 0 Then


            Dim cmdSelect = New SqlCommand("SELECT Departure_Code, dbo.GetDepartureGuide(Departure_Id) as Guide_Name,Tour_Name=case when isnull(Tour_PageTitle,'')='' then Tour_Name else Tour_PageTitle end   FROM dbo.Tours_Departures TD " & _
             " left JOIN Tours ON TD.Tour_Id = Tours.Tour_Id WHERE (Departure_Id = @DepartureId)", conPegasus)
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            conPegasus.Open()
            Dim dr_m As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
            If dr_m.Read() Then
                If Not dr_m("Departure_Code") Is DBNull.Value Then
                    DepartureCode = dr_m("Departure_Code")
                End If
                If Not dr_m("Guide_Name") Is DBNull.Value Then
                    GuideName = dr_m("Guide_Name")
                End If
                If Not dr_m("Tour_Name") Is DBNull.Value Then
                    TourName = dr_m("Tour_Name")
                End If
            End If
            conPegasus.Close()
            cmdSelect = New SqlCommand("select TourGrade,GuideGrade,OpenDevice,F.Contact_Id,C.Company_Id,C.CONTACT_NAME,SendCount,dbo.getFeedBackFaqMax(F.Contact_Id,F.departure_id) as LastFAQ from dbo.FeedBack_Form F  " & _
                          " left join " & BizpowerDBName & ".dbo.CONTACTS C on F.Contact_id=C.Contact_Id " & _
                         "  left join " & BizpowerDBName & ".dbo.sms_FeedBackTo_contact S on S.Contact_Id=F.Contact_Id and F.departure_id=F.departure_id " & _
                         "  where (FeedBack_Status = 1 or FeedBack_Status = 2)  And F.Departure_Id = @DepartureId and F.Contact_Id =@ContactId", conPegasus)
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            cmdSelect.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
            'Response.Write(conPegasus.ConnectionString & ":" & cmdSelect.CommandText)
            'Response.End()
            conPegasus.Open()
            dr_m = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
            If dr_m.Read() Then
                If Not dr_m("LastFAQ") Is DBNull.Value Then
                    LastFAQ = dr_m("LastFAQ")
                End If
                If Not dr_m("SendCount") Is DBNull.Value Then
                    SendCount = dr_m("SendCount")
                End If
                If Not dr_m("OpenDevice") Is DBNull.Value Then
                    OpenDevice = dr_m("OpenDevice")
                End If
                If Not dr_m("TourGrade") Is DBNull.Value Then
                    TourGrade = dr_m("TourGrade")
                End If
                If Not dr_m("GuideGrade") Is DBNull.Value Then
                    GuideGrade = dr_m("GuideGrade")
                End If
                If Not dr_m("CONTACT_NAME") Is DBNull.Value Then
                    CONTACTNAME = dr_m("CONTACT_NAME")
                End If


            End If


            conPegasus.Close()


            FrmId = GetForm()
            BindCategory()
            '     Dim cmdQ = New SqlCommand("select * from FeedBack_FormValue where FrmId=@FrmId", conPegasus)

        End If

        ' If status = 2 Then
        '     plhStatus2.Visible = True
        '     Dim cmdSelect2 = New SqlCommand("select F.Contact_Id,C.Company_Id,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,14) as FAQ14,dbo.getFeedBackFaq(F.Contact_Id,@DepartureId,13) as FAQ13,C.CONTACT_NAME,dbo.[getFeedBackFaqGradeMinValue](F.Contact_Id,@DepartureId) as FaqGradeMinValue from FeedBack_Form F " & _
        '"    left join " & BizpowerDBName & ".dbo.CONTACTS C on F.Contact_id=C.Contact_Id   where FeedBack_Status=2 and Departure_Id = @DepartureId and F.Contact_Id=@ContactId", conPegasus)

        '     cmdSelect2.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
        '     cmdSelect2.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
        '     conPegasus.Open()
        '     Dim dr As SqlDataReader = cmdSelect2.ExecuteReader(CommandBehavior.CloseConnection)
        '     If dr.Read Then
        '         FAQ14 = dr("FAQ14")
        '         FAQ13 = dr("FAQ13")
        '         FaqGradeMinValue = dr("FaqGradeMinValue")
        '     End If

        ' End If

    End Sub
    Public Sub BindCategory()
        Dim cmdSelect = New System.Data.SqlClient.SqlCommand("SELECT category_Id,Category_Name +'  ' + Category_Name2 as CategoryName from Feedback_Category order by Category_Id", conPegasus)
        conPegasus.Open()

        rptCat.DataSource = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        rptCat.DataBind()
        conPegasus.Close()

    End Sub
    Public Function GetForm() As Integer
        Dim myCommand As System.Data.SqlClient.SqlCommand
        Dim cmdSelecF = New SqlCommand("SELECT Id from  FeedBack_Form where Departure_Id=@DepartureId and Contact_Id=@ContactId", conPegasus)
        cmdSelecF.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
        cmdSelecF.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
        conPegasus.Open()
        FrmId = cmdSelecF.ExecuteScalar
        conPegasus.Close()
        GetForm = FrmId
    End Function

    Private Sub rptCat_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptCat.ItemDataBound
        '  Response.Write("f=" & FrmId)
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim Cat = e.Item.DataItem("Category_Id")
            rptFaq = e.Item.FindControl("rptFaq")
            Dim conF As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))
            '   If Cat < 6 Then

            Dim cmdSelFaq = New SqlClient.SqlCommand("SELECT FAQ_ID,FAQ_Title,Faq_Type,radioF from Feedback_FAQ left join FeedBack_FormValue on FeedBack_FormValue.FaqIndex=Feedback_FAQ.FAQ_ID " & _
            " where Category_Id=@Cat  and frmId=@frmId order  by FAQ_ID", conF)
            cmdSelFaq.CommandType = CommandType.Text
            cmdSelFaq.Parameters.Add("@Cat", SqlDbType.Int).Value = Cat
            cmdSelFaq.Parameters.Add("@frmId", SqlDbType.Int).Value = FrmId
            conF.Open()
            rptFaq.DataSource = cmdSelFaq.ExecuteReader()
            rptFaq.DataBind()
            conF.Close()
            'End If

        End If
    End Sub

    Private Sub rptFaq_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptFaq.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim conV As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))

            Dim selValue As String = ""
            Dim r = e.Item.DataItem("radioF")
            Dim type = e.Item.DataItem("Faq_Type")
            Dim faqid = e.Item.DataItem("FAQ_ID")
            lblValue = e.Item.FindControl("lblValue")
            lblTitle = e.Item.FindControl("lblTitle")
            lblTitle.Text = e.Item.DataItem("FAQ_Title")
            lblFAQID = e.Item.FindControl("lblFAQID")
            lblFAQID.Text = e.Item.DataItem("FAQ_ID")
            If type = 1 Then
                Select Case r
                    Case "0"
                        lblValue.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)
                        lblTitle.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)
                        lblFAQID.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)
                        lblValue.Text = "מועטה ביותר"
                    Case "1"
                        lblValue.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)
                        lblTitle.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)
                        lblFAQID.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)

                        lblValue.Text = "מועטה"
                    Case "2"
                        lblValue.Text = "בינוני"
                        lblValue.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)
                        lblTitle.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)
                        lblFAQID.ForeColor = System.Drawing.Color.FromArgb(186, 0, 0)

                    Case "3"
                        lblValue.Text = "רבה"
                    Case "4"
                        lblValue.Text = "רבה ביותר"
                End Select
            End If
            If type = 3 Then
                lblValue.Text = r
            End If

            If type = 2 Then
                If faqid = 15 Or faqid = 16 Or faqid = 17 Then
                    If Len(r) > 0 Then
                        Dim cmdSelect = New SqlCommand("select * from Feedback_FAQ_Select where [FAQ_ID]=@faqid and FAQ_Value_Number in (" & r & ")", conV)
                        cmdSelect.Parameters.Add("@faqid", SqlDbType.Int).Value = CInt(faqid)
                        ' cmdSelect.Parameters.Add("@r", SqlDbType.VarChar, 150).Value = r
                        conV.Open()
                        Dim dr_m As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

                        Do While dr_m.Read()
                            If selValue = "" Then
                                selValue = dr_m("FAQ_Value")
                            Else
                                selValue = selValue & " / " & dr_m("FAQ_Value")
                            End If

                        Loop
                        lblValue.Text = selValue

                        conV.Close()
                    End If
                End If
                If faqid = 20 And r <> "" Then
                    '  Response.Write("r=" & r)
                    '  Response.End()

                    Dim cmdSelect = New SqlCommand("select  Country_Id as FAQ_Value_Number,Country_Name as FAQ_Value   FROM Countries  where Country_Id in (" & r & " )", con)
                    cmdSelect.Parameters.Add("@faqid", SqlDbType.Int).Value = CInt(faqid)
                    ' cmdSelect.Parameters.Add("@r", SqlDbType.VarChar, 150).Value = r
                    con.Open()
                    Dim dr_m As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

                    Do While dr_m.Read()
                        If selValue = "" Then
                            selValue = dr_m("FAQ_Value")
                        Else
                            selValue = selValue & "," & dr_m("FAQ_Value")
                        End If

                    Loop
                    lblValue.Text = selValue

                    con.Close()
                End If

            End If





        End If
    End Sub
End Class
