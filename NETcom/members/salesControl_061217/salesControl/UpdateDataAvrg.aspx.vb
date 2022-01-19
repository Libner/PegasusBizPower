Imports System.Data.SqlClient
Public Class UpdateDataAvrg
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmd As New SqlClient.SqlCommand
    Public parName, Uid, parValue As String




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
        parName = Request.Form("name")
        ' parName = "Avrg_Call"
        Uid = Request.Form("Uid") '"12_2017"
        Dim pM, pY As String
        Dim strV As Array

        strV = Uid.Split("_")
        If strV.Length > 0 Then
            pM = strV(0)
            pY = strV(1)
        End If



        parValue = Request.Form("value")
        'parName = "call_Sales"
        'Uid = "2017-09-01"
        'parValue = 10


        con.Open()
        'Dim sql = "Update DimDate set " & parName & "=" & parValue & " where DateKey=" & Uid
        'Response.Write(sql)

            cmd = New System.Data.SqlClient.SqlCommand("UpdateDataAvrg_insert_Monthly" & parName, con)
            cmd.Parameters.Add("@parValue", SqlDbType.Float).Value = parValue
            cmd.Parameters.Add("@pM", SqlDbType.Int).Value = CInt(pM)
            cmd.Parameters.Add("@pY", SqlDbType.Int).Value = CInt(pY)

            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteNonQuery()
            con.Close()
        ' Response.Write(selVName & "_" & DayTypeColor & "_" & DayFontColor)
    End Sub

End Class
