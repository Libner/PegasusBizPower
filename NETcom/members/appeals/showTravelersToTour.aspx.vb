Imports System.Data.SqlClient
Public Class showTravelersToTour
    Inherits System.Web.UI.Page
    Protected BirthdayDate, IdNumber, LastName, FirstName, PasportNumber, MobileNumber, tourcode As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected WithEvents rptData As Repeater
    Protected sBirthdayDate, sIdNumber, sLastName, sFirstName, sPasportNumber, sMobileNumber As HtmlControls.HtmlInputText
    Public PegasusDBName As String = ConfigurationSettings.AppSettings("PegasusDBName")
    Protected WithEvents btnSearch As System.Web.UI.WebControls.LinkButton
    Protected sortQuery, sortQuery1 As String
    Protected appDocket As String
    Protected appid, departureId As Integer

    Public func As New bizpower.cfunc
    Protected qrystring As String
    Public actRoomJoinNumber As Integer = 0
    Protected DocketfileNum, RoomJoinNum, Departure_Code As String

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

        cmdSel = New SqlClient.SqlCommand("select Departure_Code from Tours_Departures where departure_Id=@departureId", conPegasus)
        cmdSel.Parameters.Add("@departureId", SqlDbType.Int).Value = departureId
        conPegasus.Open()
        Try
            Departure_Code = func.dbNullFix(cmdSel.ExecuteScalar)
            conPegasus.Close()
        Catch ex As Exception
            Departure_Code = ""
        Finally
            conPegasus.Close()
        End Try

        bindList()
    End Sub

    Sub bindList()

        

        ''------------  [GetTravelersData]

        '   Dim cmdSelect As New SqlClient.SqlCommand("select * from Travelers", con)

        ' Response.Write(hSelTOURCODE = "& hSelTOURCODE")

        Dim cmdSelect As New SqlClient.SqlCommand("GetTravelersRoomsDataByDepartureCode", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@PageSize", DbType.Int32).Value = 50
        cmdSelect.Parameters.Add("@PageNumber", DbType.Int32).Value = 1
        cmdSelect.Parameters.Add("@BirthdayDate", SqlDbType.VarChar, 30).Value = ""
        cmdSelect.Parameters.Add("@LastName", SqlDbType.VarChar, 50).Value = ""
        cmdSelect.Parameters.Add("@FirstName", SqlDbType.VarChar, 50).Value = ""
        cmdSelect.Parameters.Add("@MobileNumber", SqlDbType.VarChar, 50).Value = ""
        cmdSelect.Parameters.Add("@PassportNum", SqlDbType.VarChar, 50).Value = ""
        cmdSelect.Parameters.Add("@IdNumber", SqlDbType.VarChar, 50).Value = ""
        cmdSelect.Parameters.Add("@DocketfileNum", SqlDbType.VarChar, 50).Value = ""
        cmdSelect.Parameters.Add("@RoomJoinNum", SqlDbType.VarChar, 50).Value = ""
        cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery
        cmdSelect.Parameters.Add("@CountRecords", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters.Add("@DepartureId", System.Data.SqlDbType.Int).Value = departureId
        ' SelTOURCODE.SelectedItem.Value
        cmdSelect.Parameters("@CountRecords").Direction = ParameterDirection.Output
        con.Open()

        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptData.DataSource = dr
            rptData.DataBind()
            rptData.Visible = True
            'Response.Write("totalRows=" & totalRows)





        Else
            '  pnlSearchMess.Visible = True : 
            rptData.Visible = False

        End If

        dr.Close()
        cmdSelect.Dispose()
        con.Close()



    End Sub
   
    Sub rptData_onBound(ByVal s As Object, ByVal e As RepeaterItemEventArgs) Handles rptData.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim ltRoomTypeDesc As Literal
            ltRoomTypeDesc = e.Item.FindControl("ltRoomTypeDesc")
            Dim phSeparateRoomsLine As PlaceHolder
            phSeparateRoomsLine = e.Item.FindControl("phSeparateRoomsLine")

            Dim currentRoomJoinNumber As Integer
            currentRoomJoinNumber = func.fixNumeric(e.Item.DataItem("joinRoomNum"))
            If currentRoomJoinNumber <> actRoomJoinNumber Then
                actRoomJoinNumber = currentRoomJoinNumber
                If e.Item.ItemIndex > 0 Then
                    phSeparateRoomsLine.Visible = True
                End If
            End If

            'room type
            If func.fixNumeric(e.Item.DataItem("paxPerRoom")) = 0 And func.dbNullBool(e.Item.DataItem("reservedAsRequestShare")) = True Then
                ltRoomTypeDesc.Text = "SHARING REQUEST"
            End If

            If func.fixNumeric(e.Item.DataItem("paxPerRoom")) = 1 Then
                ltRoomTypeDesc.Text = "SINGLE"
            End If
            If func.dbNullBool(e.Item.DataItem("doubleForSingle")) = True Then
                ltRoomTypeDesc.Text = "DOUBLE FOR SINGLE"
            End If
            If func.fixNumeric(e.Item.DataItem("paxPerRoom")) = 2 Then
                If func.dbNullBool(e.Item.DataItem("isTwin")) = True Then
                    ltRoomTypeDesc.Text = "TWIN"
                Else
                    ltRoomTypeDesc.Text = "DOUBLE"
                End If
            End If
            If func.fixNumeric(e.Item.DataItem("paxPerRoom")) = 3 Then
                ltRoomTypeDesc.Text = "TRIPLE"
            End If
            If func.fixNumeric(e.Item.DataItem("paxPerRoom")) = 4 Then
                ltRoomTypeDesc.Text = "QUADRUPLE"
            End If
            If func.fixNumeric(e.Item.DataItem("paxPerRoom")) = 5 Then
                ltRoomTypeDesc.Text = "QUINTUPLE"
            End If

            If func.dbNullBool(e.Item.DataItem("childShare")) = True Then
                ltRoomTypeDesc.Text = "CHILD"
            End If
        End If
    End Sub
End Class
