Public Class modify_records
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Public FinancialEuro_val, FinancialDolar_val, DepartureTimeBrief_val, TimeMeetingAfterTrip_val, Departure_TimeGroupMeeting_val As String
    Public GilboaHotel_val, isSendLetterGroupMeeting_val, isSendToGuide_val, isSendTicket_val, Voucher_Simultaneous_val, Departure_Costing_val, Voucher_Group_val, Departure_GuideTelphone_val, Vouchers_Provider_val As String

    Protected WithEvents btnSubmit As UI.WebControls.Button

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

        If IsNumeric(Request.Form("row_id")) Then
            DepartureId = CInt(Request.Form("row_id"))
        End If
        FinancialEuro_val = Trim(Request.Form("FinancialEuro_val"))
        FinancialDolar_val = Trim(Request.Form("FinancialDolar_val"))
        DepartureTimeBrief_val = Trim(Request.Form("DepartureTimeBrief_val"))
        TimeMeetingAfterTrip_val = Trim(Request.Form("TimeMeetingAfterTrip_val"))
        Departure_TimeGroupMeeting_val = Trim(Request.Form("Departure_TimeGroupMeeting_val"))
        GilboaHotel_val = Trim(Request.Form("GilboaHotel_val"))

        isSendLetterGroupMeeting_val = Trim(Request.Form("isSendLetterGroupMeeting_val"))
        isSendToGuide_val = Trim(Request.Form("isSendToGuide_val"))
        isSendTicket_val = Trim(Request.Form("isSendTicket_val"))

        Voucher_Simultaneous_val = Trim(Request.Form("Voucher_Simultaneous_val"))
        Departure_Costing_val = Trim(Request.Form("Departure_Costing_val"))
        Voucher_Group_val = Trim(Request.Form("Voucher_Group_val"))
        Departure_GuideTelphone_val = Trim(Request.Form("Departure_GuideTelphone_val"))
        Vouchers_Provider_val = Trim(Request.Form("Vouchers_Provider_val"))
        cmdSelect = New SqlClient.SqlCommand("update Tours_Departures set Financial_Euro=@Financial_Euro,Financial_Dolar=@Financial_Dolar,Departure_TimeBrief=@DepartureTimeBrief,TimeMeetingAfterTrip=@TimeMeetingAfterTrip,Departure_TimeGroupMeeting=@Departure_TimeGroupMeeting" & _
        " ,GilboaHotel=@GilboaHotel,isSendLetterGroupMeeting=@isSendLetterGroupMeeting, isSendToGuide=@isSendToGuide, isSendTicket=@isSendTicket,Voucher_Simultaneous= @Voucher_Simultaneous,Departure_Costing=@Departure_Costing,Voucher_Group=@Voucher_Group,Departure_GuideTelphone=@Departure_GuideTelphone,Vouchers_Provider=@Vouchers_Provider " & _
      " WHERE (Departure_Id = @Departure_Id)", conPegasus)
        cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
        cmdSelect.Parameters.Add("@Financial_Euro", SqlDbType.Text).Value = FinancialEuro_val
        cmdSelect.Parameters.Add("@Financial_Dolar", SqlDbType.Text).Value = FinancialDolar_val
        cmdSelect.Parameters.Add("@DepartureTimeBrief", SqlDbType.VarChar, 5).Value = IIf(Trim(DepartureTimeBrief_val) = "", DBNull.Value, Trim(DepartureTimeBrief_val))

        cmdSelect.Parameters.Add("@TimeMeetingAfterTrip", SqlDbType.VarChar, 5).Value = IIf(Trim(TimeMeetingAfterTrip_val) = "", DBNull.Value, Trim(TimeMeetingAfterTrip_val))
        cmdSelect.Parameters.Add("@Departure_TimeGroupMeeting", SqlDbType.VarChar, 5).Value = IIf(Trim(Departure_TimeGroupMeeting_val) = "", DBNull.Value, Trim(Departure_TimeGroupMeeting_val))
        cmdSelect.Parameters.Add("@GilboaHotel", SqlDbType.VarChar, 2).Value = Trim(GilboaHotel_val)

        cmdSelect.Parameters.Add("@isSendLetterGroupMeeting", SqlDbType.Bit).Value = IIf(Trim(isSendLetterGroupMeeting_val) = "λο", "True", "False")
        cmdSelect.Parameters.Add("@isSendToGuide", SqlDbType.Bit).Value = IIf(Trim(isSendToGuide_val) = "λο", "True", "False")
        cmdSelect.Parameters.Add("@isSendTicket", SqlDbType.Bit).Value = IIf(Trim(isSendTicket_val) = "λο", "True", "False")


        cmdSelect.Parameters.Add("@Voucher_Simultaneous", SqlDbType.VarChar, 2).Value = Trim(Voucher_Simultaneous_val)
        cmdSelect.Parameters.Add("@Departure_Costing", SqlDbType.VarChar, 2).Value = Trim(Departure_Costing_val)
        cmdSelect.Parameters.Add("@Voucher_Group", SqlDbType.VarChar, 2).Value = Trim(Voucher_Group_val)
        cmdSelect.Parameters.Add("@Departure_GuideTelphone", SqlDbType.Text).Value = Trim(Departure_GuideTelphone_val)
        cmdSelect.Parameters.Add("@Vouchers_Provider", SqlDbType.VarChar, 5).Value = Trim(Vouchers_Provider_val)

        conPegasus.Open()
        cmdSelect.ExecuteNonQuery()
        conPegasus.Close()
        Response.Write("upd=" & DepartureId & ":" & isSendLetterGroupMeeting_val)
    End Sub

End Class
