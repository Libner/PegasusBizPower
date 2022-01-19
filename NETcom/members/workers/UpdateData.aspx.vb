Imports System.Data.SqlClient

Public Class UpdateData
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
        Dim PName As String

        parName = Request("name")
        Uid = Request("Uid")
        parValue = Request("value")
        con.Open()

        Select Case parName
            Case "OrderPrice"
                cmd = New System.Data.SqlClient.SqlCommand("Update Users set Order_Price=" & Trim(parValue) & " where USER_ID=" & Uid, con)
                '      '  PName = "Order_Price"
            Case "minOrder"
                cmd = New System.Data.SqlClient.SqlCommand("Update Users set Month_Min_Order=" & Trim(parValue) & " where USER_ID=" & Uid, con)

                'PName = "Month_Min_Order"
            Case "email"
                cmd = New System.Data.SqlClient.SqlCommand("Update Users set EMAIL='" & Trim(parValue) & "' where USER_ID=" & Uid, con)

                PName = "EMAIL"

        End Select



        ' Dim s = "Update Users set " & parName & "=@parValue where User_Id=" & Uid
        '    Response.Write(s)
        '    Response.End()
        ' cmd = New System.Data.SqlClient.SqlCommand("Update Users set " & PName & "=" & Trim(parValue) & " where USER_ID=" & Uid, con)

        'cmd.Parameters.Add("@parName", SqlDbType.VarChar).Value = parName
        'cmd.Parameters.Add("@parValue", SqlDbType.VarChar).Value = Trim(parValue)
        'cmd.Parameters.Add("@Uid", SqlDbType.Int).Value = CInt(Uid)
        cmd.CommandType = CommandType.Text
        cmd.ExecuteNonQuery()

        Dim exc As Integer
        Try
            '  exc = cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try


        con.Close()
        Response.Write(exc)
        'Response.End()
        ' Response.Write(selVName & "_" & DayTypeColor & "_" & DayFontColor)
    End Sub

End Class
