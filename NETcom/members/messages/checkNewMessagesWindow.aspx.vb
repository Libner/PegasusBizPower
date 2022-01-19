Public Class checkNewMessagesWindow
    Inherits System.Web.UI.Page
    Protected strReturnHtml, strReturnalt, messId As String
    Protected replFlag As Boolean
    Protected UserId As Integer
    Dim con As New SqlClient.SqlConnection
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Protected windowF As Integer = 0
    Protected count_messages As Integer = 0

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
        If IsNumeric(Request.Cookies("bizpegasus")("UserId")) Then
            ' Response.Write(Request.Cookies("bizpegasus")("UserId"))
            '  Response.End()
            con = New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            cmdSelect = New SqlClient.SqlCommand("GetNewInBox", con)
            cmdSelect.CommandType = CommandType.StoredProcedure
            ' cmdSelect.Parameters.Add("@cookieCode", Request.Cookies("bizpegasus")("UserId"))
            cmdSelect.Parameters.Add("@cookieCode", Request.Cookies("bizpegasus")("UserId"))

            con.Open()
            Dim dt As SqlClient.SqlDataReader

            dt = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
            ' If dt.HasRows() Then
            If dt.Read() Then
                If Not dt(0) Is DBNull.Value Then
                    messId = CStr(Trim(dt(0)))
                End If
                If Not dt(1) Is DBNull.Value Then
                    replFlag = CStr(Trim(dt(1)))
                End If
                con.Close()
                Dim cmdExecute As SqlClient.SqlCommand
                Dim sqlstr As String

                '13.11.13 update flag U
                'sqlstr = "Update messages Set MessageUP_Flag = 1 Where message_id = " & messId
                'con.Open()
                'cmdExecute = New SqlClient.SqlCommand(sqlstr, con)
                'cmdExecute.ExecuteNonQuery()

                'cmdExecute.Dispose()
                'con.Close()

                'Response.Write("messText=" & messText)
                'Response.End()
                '   strReturnalt = "<script language='javascript'>" & vbCrLf
                'If replFlag = False Then
                '    strReturnalt += "window.parent.open('ViewMessage.asp?messageId=" & messId & "', ""Wind"" ,""scrollbars=1,toolbar=0,top=50,left=100,width=660,height=220,align=center,resizable=0"");" & vbCrLf
                'Else
                '    strReturnalt += "window.parent.open('CloseMessage.asp?messageId=" & messId & "', ""Wind"" ,""scrollbars=1,toolbar=0,top=50,left=100,width=600,height=430,align=center,resizable=0"");" & vbCrLf

                'End If

                '  strReturnalt += "</script>" & vbCrLf

                If replFlag = False Then
                    strReturnHtml = "http://bluto/bizpower_pegasus/netcom/members/messages/ViewMessageWindow.asp?messageId=" & messId
                Else
                    strReturnHtml = "http://bluto/bizpower_pegasus/netcom/members/messages/CloseMessageWindow.asp?messageId=" & messId

                End If
                windowF = 1
                '    RegisterStartupScript("AlertScript", strReturnalt)
                ' strReturnHtml = "yes"

            End If
            End If
    End Sub

End Class
