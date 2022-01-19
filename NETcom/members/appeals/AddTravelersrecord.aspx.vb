Public Class AddTravelersrecord
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Protected strTravelerid, appDocket As String
    Protected appid As Integer
    Protected arrayT As Array
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected ContactId, TourId, DepartureId As Integer



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
        Dim i As Integer = 0
        strTravelerid = Request("traveler_id")
        appDocket = Trim(Request("appDocket"))
        appid = func.fixNumeric(Request("appid"))
        cmdSelect = New SqlClient.SqlCommand("Select Contact_Id,Tour_Id,Departure_Id from appeals where appeal_id=@appid", con)
        cmdSelect.Parameters.Add("@appid", SqlDbType.Int).Value = CInt(appid)
        cmdSelect.CommandType = CommandType.Text
         con.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
        While dr.Read()
            If Not IsDBNull(dr("Contact_Id")) Then
                ContactId = dr("Contact_Id")
            End If
            If Not IsDBNull(dr("Tour_Id")) Then
                TourId = dr("Tour_Id")
            End If
            If Not IsDBNull(dr("Departure_Id")) Then
                DepartureId = dr("Departure_Id")
            End If
        End While
        con.Close()



        arrayT = Split(strTravelerid, ",")
       
        For i = 0 To UBound(arrayT)
            con.Open()
            Try
             '===modified by Mila 27/04/2020 (update number of people in appeal according to new number of travelers after all insertings instead + one)
            'cmdSelect = New System.Data.SqlClient.SqlCommand("Insert into  Tour_Travelers (Traveler_Id,Tour_Id,Departure_Id,Docket_pfileNum,Registration_Date,Contact_Id) " & _
            '" values(@TravelerId,@TourId,@DepartureId,@appDocket,GetDate(),@ContactId);update FORM_VALUE set FIELD_VALUE=convert(int,FIELD_VALUE)+1  where FIELD_ID=40660 and APPEAL_ID=@appid and IsNumeric(FIELD_VALUE)=1;", con)
            cmdSelect = New System.Data.SqlClient.SqlCommand("Insert into  Tour_Travelers (Traveler_Id,Tour_Id,Departure_Id,Docket_pfileNum,Registration_Date,Contact_Id,Appeal_Id) " & _
            " values(@TravelerId,@TourId,@DepartureId,@appDocket,GetDate(),@ContactId,@appID);", con)
            cmdSelect.CommandType = CommandType.Text
            cmdSelect.Parameters.Add("@TravelerId", SqlDbType.Int).Value = CInt(arrayT(i))
            cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = TourId
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = DepartureId
            cmdSelect.Parameters.Add("@appDocket", SqlDbType.BigInt).Value = appDocket
            cmdSelect.Parameters.Add("@ContactId", SqlDbType.Int).Value = ContactId
            cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appid)
            cmdSelect.ExecuteNonQuery()
            con.Close()
            Catch ex As Exception
                ' Response.Write("BR>" & Err.Description)
            Finally
                con.Close()
            End Try

        Next

        '===Added by Mila 27/04/2020 
        'number of passangers - hatum rishum
        If appid > 0 Then
            cmdSelect = New System.Data.SqlClient.SqlCommand("select count(Traveler_Id) from Tour_Travelers " & _
            " where Departure_Id=@DepartureId and Docket_pfileNum=@appDocket", con)
            cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = DepartureId
            cmdSelect.Parameters.Add("@appDocket", SqlDbType.BigInt).Value = appDocket
            Dim numTravelers As Integer = 0
            con.Open()
            Try
                numTravelers = func.fixNumeric(cmdSelect.ExecuteScalar)
            Catch ex As Exception
            Finally
                con.Close()

            End Try

            Dim sqlupdNm As New SqlClient.SqlCommand("update FORM_VALUE set Field_Value =@Field_Value where Appeal_Id=@appID and Field_Id=@Field_Id", con)
            sqlupdNm.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appid)
            sqlupdNm.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40660)
            sqlupdNm.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = numTravelers 'number of passangers
            con.Open()
            Try
                sqlupdNm.ExecuteNonQuery()
                con.Close()
            Catch ex As Exception

            Finally
                con.Close()
            End Try


        End If
        Response.Write("פעולה בוצה בהצלחה") ' & appDocket & ":--" & appid)
       
    End Sub

End Class
