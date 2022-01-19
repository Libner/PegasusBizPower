Public Class editTextAdmin
    Inherits System.Web.UI.Page
    Public numOftab, elemId, templateSource, templateId, pwidth, perSize, templateTitle As String
    Protected WithEvents SqlConn As System.Data.SqlClient.SqlConnection
    Protected WithEvents catCMD As System.Data.SqlClient.SqlCommand
    Dim myReader As System.Data.SqlClient.SqlDataReader
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents title1 As System.Web.UI.HtmlControls.HtmlTextArea
    Protected WithEvents templateId1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents templateTitle1 As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents PHAlert As System.Web.UI.WebControls.PlaceHolder

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here     
        numOftab = 3
        templateId = Request("templateId")
        templateId1.Value = templateId

        SqlConn = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        If Not IsPostBack Then
            pwidth = "645"
            If templateId <> "" Then

                catCMD = New System.Data.SqlClient.SqlCommand("Select template_Title,template_Source From templates Where template_Id=" & templateId, SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

                If myReader.Read() Then

                    If Not myReader("template_Source") Is DBNull.Value Then
                        templateSource = Trim(myReader("template_Source"))
                    Else
                        templateSource = ""
                    End If

                    If Not myReader("template_Title") Is DBNull.Value Then
                        templateTitle = Trim(myReader("template_Title"))
                    Else
                        templateTitle = ""
                    End If

                End If

                myReader.Close()
                SqlConn.Close()
                title1.Style.Item("WIDTH") = "645px"
                title1.Style.Item("HEIGHT") = "360px"
                title1.Value = templateSource

                templateTitle1.Value = templateTitle

            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Trim(templateTitle1.Value) = "" Then
            Dim stralert As String
            stralert = "<script language='javascript'>" & vbCrLf
            stralert += "alert('חובה למלא שם של דף מעוצב');" & vbCrLf
            stralert += "</script>" & vbCrLf
            PHAlert.Controls.Add(New LiteralControl(stralert))
        ElseIf Trim(title1.Value) = "" Then
            Dim stralert As String
            stralert = "<script language='javascript'>" & vbCrLf
            stralert += "alert('חובה למלא את השדה טקסט');" & vbCrLf
            stralert += "</script>" & vbCrLf
            PHAlert.Controls.Add(New LiteralControl(stralert))
        Else
            catCMD = New System.Data.SqlClient.SqlCommand("updateTemplateContent", SqlConn)
            catCMD.CommandType = CommandType.StoredProcedure
            If Trim(templateId) = "" Or IsNumeric(templateId) = False Then
                templateId = 0
            End If

            catCMD.Parameters.Add("@templateId", templateId)
            catCMD.Parameters.Add("@templateTitle", templateTitle1.Value)
            catCMD.Parameters.Add("@templateSource", title1.InnerText)

            SqlConn.Open()
            catCMD.ExecuteNonQuery()
            SqlConn.Close()

            If templateId = 0 Then
                catCMD = New System.Data.SqlClient.SqlCommand("Select Max(Template_ID) From Templates", SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

                If myReader.Read() Then

                    If Not myReader(0) Is DBNull.Value Then
                        templateId = Trim(myReader(0))
                    Else
                        templateId = "0"
                    End If

                End If

                myReader.Close()
                SqlConn.Close()

            End If
            
            Response.Redirect("editPageAdmin.aspx?templateId=" & templateId)

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Response.Redirect("default.asp")
    End Sub

End Class
