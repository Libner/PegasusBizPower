Public Class editTextE
    Inherits System.Web.UI.Page
    Public catida, elemId, elText, pageId, pwidth, ptitle, siteId, siteName, maincat, preview As String
    Protected WithEvents SqlConn As System.Data.SqlClient.SqlConnection
    Protected WithEvents catCMD As System.Data.SqlClient.SqlCommand
    Dim myReader As System.Data.SqlClient.SqlDataReader
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents content As System.Web.UI.HtmlControls.HtmlTextArea
    Protected WithEvents PHAlert As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents Button3 As System.Web.UI.WebControls.Button
    Protected WithEvents previewpage As System.Web.UI.HtmlControls.HtmlInputHidden

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
        pageId = Request.QueryString("pageId")
        catida = Request.QueryString("catId")
        maincat = Request.QueryString("maincat")
        siteId = Request.QueryString("siteId")
        ' preview = Request.QueryString("preview")
        SqlConn = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        If Not IsPostBack Then
            pwidth = "100%"
            If maincat <> "" Then
                catCMD = New System.Data.SqlClient.SqlCommand("SELECT Site_Name FROM Sites WHERE Site_Id IN (SELECT Site_Id FROM Main_Categories WHERE Main_Category_ID=" & maincat & ")", SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)
                If myReader.Read() Then
                    If Not myReader("Site_Name") Is DBNull.Value Then
                        siteName = Trim(myReader("Site_Name"))
                    End If
                End If
            End If

            myReader.Close()
            SqlConn.Close()

            If pageId <> "" Then

                catCMD = New System.Data.SqlClient.SqlCommand("Select Page_Title,Page_Width,Page_content From pagesTav Where Page_Id=" & pageId, SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

                If myReader.Read() Then
                    If Not myReader("Page_Title") Is DBNull.Value Then
                        ptitle = Trim(myReader("Page_Title"))
                    End If

                    If Not myReader("Page_Width") Is DBNull.Value Then
                        pwidth = Trim(myReader("Page_Width"))
                        If InStr(LCase(pwidth), "px") > 0 Or InStr(pwidth, "%") = 0 Then
                            pwidth = CStr(CInt(Replace(LCase(pwidth), "px", "")) + 35) & "px"
                        End If
                    End If

                    If Not myReader("Page_content") Is DBNull.Value Then
                        elText = Trim(myReader("Page_content"))
                    Else
                        'elText = "<P style='DIRECTION: rtl' align=right>&nbsp;</P>"
                        If siteId <> 2 Then
                            elText = "<P dir=rtl>&nbsp;</P>"
                        Else
                            elText = "<P>&nbsp;</P>"
                        End If

                    End If
                End If

                myReader.Close()
                SqlConn.Close()
                content.Style.Item("WIDTH") = pwidth
                content.Style.Item("HEIGHT") = "360px"
                content.Value = elText
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        UpdatePage()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button1_Click(sender, e)
        previewpage.Value = "yes"
        'Response.Redirect("editPage.aspx?maincat=" & maincat & "&catId=" & catida & "&PageId=" & pageId & "&subcat=" & Request.QueryString("subcat") & "&innerparent=" & Request.QueryString("innerparent") & "&parentID=" & Request.QueryString("parentID") & "&preview=1", True)
    End Sub

    Private Sub UpdatePage()
        If Trim(content.Value) = "" Then
            Dim stralert As String
            stralert = "<script language='javascript'>" & vbCrLf
            stralert += "alert('חובה למלא את השדה טקסט');" & vbCrLf
            stralert += "</script>" & vbCrLf
            PHAlert.Controls.Add(New LiteralControl(stralert))
        Else
            catCMD = New System.Data.SqlClient.SqlCommand("updatePageTavContent", SqlConn)
            catCMD.CommandType = CommandType.StoredProcedure
            catCMD.Parameters.Add("@pageId", pageId)
            catCMD.Parameters.Add("@content_text", content.InnerText)

            SqlConn.Open()
            catCMD.ExecuteNonQuery()
            SqlConn.Close()
        End If
    End Sub
End Class
