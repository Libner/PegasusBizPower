Imports System.Data.SqlClient
Public Class AddTravelerToTours
    Inherits System.Web.UI.Page
    Protected BirthdayDate, IdNumber, LastName, FirstName, PasportNumber, MobileNumber, tourcode As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected rptData As Repeater
    Protected sBirthdayDate, sIdNumber, sLastName, sFirstName, sPasportNumber, sMobileNumber As HtmlControls.HtmlInputText
    Public PegasusDBName As String = ConfigurationSettings.AppSettings("PegasusDBName")
    Protected WithEvents btnSearch As System.Web.UI.WebControls.LinkButton
    Protected sortQuery, sortQuery1 As String
    Protected appDocket As String
    Protected appid, departureId As Integer

    Public func As New bizpower.cfunc
    Protected qrystring As String

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
        If IsNumeric(Request.Form("appDocket")) Then
            appDocket = Request.Form("appDocket")
        ElseIf IsNumeric(Request.QueryString("appDocket")) Then
            appDocket = Request.QueryString("appDocket")
        End If
        appid = Request("appid")
        Dim cmdSel As New SqlClient.SqlCommand("select Departure_Id from Appeals where appeal_Id=@appId", con)
        cmdSel.Parameters.Add("@appid", SqlDbType.Int).Value = appid
        con.Open()
        Try
            departureId = func.fixNumeric(cmdSel.ExecuteScalar)
            con.Close()
        Catch ex As Exception
            departureId = 0
        Finally
            con.Close()
        End Try

        Dim r As Integer
        sortQuery = ""
        qrystring = Request.ServerVariables("QUERY_STRING")
        r = qrystring.IndexOf("sort")
        If r > 0 Then
            '  Response.Write("<BR>gg=" & Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)) & "<BR>")
            sortQuery1 = Replace(Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)), "&", ",")
            sortQuery1 = Replace(sortQuery1, "=", " ")
            sortQuery1 = Replace(sortQuery1, "sort_1", "IdNumber")
            sortQuery1 = Replace(sortQuery1, "sort_2", "LastName")
            sortQuery1 = Replace(sortQuery1, "sort_3", "FirstName")
            sortQuery1 = Replace(sortQuery1, "sort_4", "BirthDate")
            sortQuery1 = Replace(sortQuery1, "sort_5", "PassportNum")
            sortQuery1 = Replace(sortQuery1, "sort_6", "Phone")


            '  sortQuery1 = Replace(sortQuery1, "sort_8", "TIME_LIMIT")

            sortQuery1 = sortQuery1.Substring(1)

            sortQuery = sortQuery1
            '  Response.Write("<BR>" & sortQuery1 & "<BR>")
            '         Response.End() '   sortQuery = sortQuery1
        End If
        bindList()
    End Sub

    Sub bindList()




        '    Response.Write("tourcode=" & tourcode)

        If IsDate(Request.Form("sBirthdayDate")) Then
            BirthdayDate = Request.Form("sBirthdayDate")
            sBirthdayDate.Value = BirthdayDate
        ElseIf IsDate(Request.QueryString("sBirthdayDate")) Then
            BirthdayDate = Request.QueryString("sBirthdayDate")
            sBirthdayDate.Value = BirthdayDate
        Else
            BirthdayDate = ""
        End If

        If Len(Request.Form("sIdNumber")) > 0 Then
            IdNumber = Request.Form("sIdNumber")
            sIdNumber.Value = IdNumber
        ElseIf Len(Request.QueryString("sIdNumber")) > 0 Then
            IdNumber = Request.QueryString("sIdNumber")
            sIdNumber.Value = IdNumber
        Else
            IdNumber = ""
        End If

        If Len(Request.Form("sLastName")) > 0 Then
            LastName = Request.Form("sLastName")
            sLastName.Value = LastName
        ElseIf Len(Request.QueryString("sLastName")) > 0 Then
            LastName = Request.QueryString("sLastName")
            sLastName.Value = LastName
        Else
            LastName = ""
        End If


        If Len(Request.Form("sFirstName")) > 0 Then
            FirstName = Request.Form("sFirstName")
            sFirstName.Value = FirstName
        ElseIf Len(Request.QueryString("sFirstName")) > 0 Then
            LastName = Request.QueryString("sFirstName")
            sFirstName.Value = FirstName
        Else
            FirstName = ""
        End If

        If Len(Request.Form("sPasportNumber")) > 0 Then
            PasportNumber = Request.Form("sPasportNumber")
            sPasportNumber.Value = PasportNumber
        ElseIf Len(Request.QueryString("sPasportNumber")) > 0 Then
            LastName = Request.QueryString("sPasportNumber")
            sPasportNumber.Value = PasportNumber
        Else
            PasportNumber = ""
        End If

        If Len(Request.Form("sMobileNumber")) > 0 Then
            MobileNumber = Request.Form("sMobileNumber")
            sMobileNumber.Value = MobileNumber
        ElseIf Len(Request.QueryString("sMobileNumber")) > 0 Then
            MobileNumber = Request.QueryString("sMobileNumber")
            sMobileNumber.Value = MobileNumber
        Else
            MobileNumber = ""
        End If
        '  Response.Write("MobileNumber=" & MobileNumber)

        ''------------  [GetTravelersData]
        'Response.Write("@depId=" & departureId)
        'Response.End()
        '   Dim cmdSelect As New SqlClient.SqlCommand("select * from Travelers", con)
        'Modified by Mila 27/04/20 - without docket
        'Dim cmdSelect As New SqlClient.SqlCommand("GetTravelersByParam", con)
        Dim cmdSelect As New SqlClient.SqlCommand("GetTravelersByParam", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@PageSize", DbType.Int32).Value = 100 ' Trim(PageSize.SelectedValue)
        cmdSelect.Parameters.Add("@PageNumber", DbType.Int32).Value = 1 ' CInt(CurrentIndex) + 1
        cmdSelect.Parameters.Add("@BirthdayDate", SqlDbType.VarChar, 30).Value = BirthdayDate
        cmdSelect.Parameters.Add("@LastName", SqlDbType.VarChar, 50).Value = LastName
        cmdSelect.Parameters.Add("@FirstName", SqlDbType.VarChar, 50).Value = FirstName
        cmdSelect.Parameters.Add("@MobileNumber", SqlDbType.VarChar, 50).Value = MobileNumber
        cmdSelect.Parameters.Add("@PassportNum", SqlDbType.VarChar, 50).Value = PasportNumber
        cmdSelect.Parameters.Add("@IdNumber", SqlDbType.VarChar, 50).Value = IdNumber
        cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery
        cmdSelect.Parameters.Add("@appDocket", SqlDbType.BigInt).Value = appDocket
        'Modified by Mila 27/04/20 - only travelers not joined to departure
        cmdSelect.Parameters.Add("@depId", SqlDbType.Int).Value = departureId
        cmdSelect.Parameters.Add("@CountRecords", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@CountRecords").Direction = ParameterDirection.Output
        con.Open()


        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptData.DataSource = dr
            rptData.DataBind()
            rptData.Visible = True
         Else
            '  pnlSearchMess.Visible = True : 
            rptData.Visible = False
        End If

        dr.Close()
        cmdSelect.Dispose()
        con.Close()

    End Sub
    Sub btnSearch_onClick(ByVal s As Object, ByVal e As EventArgs) Handles btnSearch.Click


        bindList()
    End Sub

End Class
