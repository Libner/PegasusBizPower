Imports System.Data.SqlClient
Public Class UpdateSelect
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmd As New SqlClient.SqlCommand
    Protected UId, selV As Integer
    Protected selVName As String
    Protected func As New include.funcs
    Dim myReader As System.Data.SqlClient.SqlDataReader


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
        UId = Request.Form("id")
        selV = Request.Form("selV")
        con.Open()
        cmd = New System.Data.SqlClient.SqlCommand("Update Users set Department_Id=@selV where USER_ID=@UId", con)
        cmd.Parameters.Add("@selV", SqlDbType.Int).Value = CInt(selV)
        cmd.Parameters.Add("@UId", SqlDbType.Int).Value = CInt(UId)
           cmd.CommandType = CommandType.Text
        ' cmd.Parameters.Add("@chkValue", SqlDbType.Int).Value = CInt(chk.Value())
        'cmd.Parameters.Add("@ProfId", SqlDbType.Int).Value = CInt(chk.Attributes("value"))
        cmd.ExecuteNonQuery()
        con.Close()
        Response.Write(GetDepName(selV))
    End Sub
    Function GetDepName(ByVal ValueStr As String) As String
        Dim result As String
        Dim ss As String
        ss = ValueStr
        '  Response.Write("ss=" & ss)
        '  Response.End()
        Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        If ss <> "" Then

            Dim cmdSelect As New SqlClient.SqlCommand("SELECT departmentName from Departments  where  departmentId=" & ss, myConnection)
            cmdSelect.CommandType = CommandType.Text
            myConnection.Open()

            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("departmentName")) Then
                    If result <> "" Then
                        result = result & ", " & myReader("departmentName")
                    Else
                        result = myReader("departmentName")
                    End If


                End If

            End While
            'If result = "" Then
            '    result = "הכל"
            'End If
            myConnection.close()


            Return result
        End If
    End Function


End Class
