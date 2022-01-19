Imports System.Data.SqlClient
Public Class UpdateDataMonthly
    Inherits System.Web.UI.Page

    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmd As New SqlClient.SqlCommand
    Public parName, Uid, parValue, dep As String


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
        Uid = Request.Form("Uid") '"12_2017" month -dep - year
        parValue = Request.Form("value")
        ' dep = Request.Form("dep")
        Dim pM, pY As String
        Dim strV As Array
        strV = Uid.Split("_")
        If strV.Length > 0 Then
            pM = strV(0)
            dep = strV(1)
            pY = strV(2)
        End If

        con.Open()
        'Dim sql = "Update DimDate set " & parName & "=" & parValue & " where DateKey=" & Uid
        'Response.Write(sql)
      
        cmd = New System.Data.SqlClient.SqlCommand("update_Departments_Data_Monthly", con)

        cmd.Parameters.Add("@parName", SqlDbType.VarChar).Value = parName
        cmd.Parameters.Add("@parValue", SqlDbType.VarChar).Value = parValue
        cmd.Parameters.Add("@pM", SqlDbType.Char, 2).Value = CInt(pM)
        cmd.Parameters.Add("@pY", SqlDbType.Char, 4).Value = CInt(pY)
        cmd.Parameters.Add("@dep", SqlDbType.Int).Value = CInt(dep)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        con.Close()

    End Sub


End Class
