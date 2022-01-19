Imports System.Data.SqlClient
Public Class utilitiesUpdateHolyday
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public func As New bizpower.cfunc


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
        con.Open()
        Dim cmdselectSubR = New SqlClient.SqlCommand("Select  Date,DayName,DayTypeColor,DayTypeName From HOLIDAYS", con)

        Dim rsSubR As SqlClient.SqlDataReader = cmdselectSubR.ExecuteReader()
        Do While rsSubR.Read()
            Dim PDate = rsSubR("Date")
            Dim DayName = rsSubR("DayName")
            Dim DayTypeColor = rsSubR("DayTypeColor")
            Dim DayTypeName = rsSubR("DayTypeName")
            Dim con1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
            Dim cmdUpdate As New SqlClient.SqlCommand("Update DimDate Set HolidayName=@DayName,DayTypeName=@DayTypeName,DayTypeColor=@DayTypeColor,IsHoliday=1  Where DateKey=@PDate", con1)
                    cmdUpdate.CommandType = CommandType.Text
            cmdUpdate.Parameters.Add("@PDate", SqlDbType.DateTime).Value = PDate
            cmdUpdate.Parameters.Add("@DayName", SqlDbType.NVarChar, 100).Value = DayName
            cmdUpdate.Parameters.Add("@DayTypeColor", SqlDbType.NVarChar, 100).Value = DayTypeColor
            cmdUpdate.Parameters.Add("@DayTypeName", SqlDbType.NVarChar, 100).Value = DayTypeName
          
            con1.Open()
                    cmdUpdate.ExecuteNonQuery()
                    con1.Close()
                 
        Loop
        rsSubR.Close()
        con.Close()
    End Sub

End Class
