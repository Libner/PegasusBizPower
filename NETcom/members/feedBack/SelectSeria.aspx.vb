Public Class SelectSeria1
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dtSerias As New DataTable
    Dim primKeydtSerias(0) As Data.DataColumn
    Dim myReader As System.Data.SqlClient.SqlDataReader
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Public Departure_GuideTelphone As String
    Protected WithEvents btnSubmit As UI.WebControls.Button
    Protected fName, fValue, DepartureCode, fTitle As String
    Protected WithEvents fselect As HtmlSelect
    Dim ss, ssName, val_SeriasSelect, val_User As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

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
        'If IsNothing(Request.Cookies(ConfigurationSettings.AppSettings("AdminCookieName"))) OrElse _
        'Not IsNumeric(Request.Cookies(ConfigurationSettings.AppSettings("AdminCookieName"))("workerid")) Then
        '    Response.Redirect("../default.aspx")
        'End If
        val_SeriasSelect = Request.QueryString("v")
        val_User = Request.QueryString("user")
        '   Response.Write("2=" & Request.Form("fselect"))
        '  Response.End()


        If Not Page.IsPostBack Then



            '  Response.Write(fName & ":" & Request.Form("fselect"))
            ' Response.End()
            ' Dim ss As String
            '  ss = getSerias(CInt(Request.Form("fselect")))
            getSerias()

        Else

            ss = Request.Form("fselect")
            ssName = GetSelectSeriaName()
            Dim cScript As String
            If ss = 0 Then
                cScript = "<script language='javascript'>opener.document.getElementById('SeriesSelect').value = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; opener.document.getElementById('SelectSeriasAlt').src='../../images/select.png';opener.document.getElementById('SelectSeriasAlt').title='" & ssName & "';self.close(); </script>"

            Else
                cScript = "<script language='javascript'>opener.document.getElementById('SeriesSelect').value = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; opener.document.getElementById('SelectSeriasAlt').src='../../images/select_act.png';opener.document.getElementById('SelectSeriasAlt').title='" & ssName & "';self.close(); </script>"

            End If
            RegisterStartupScript("ReloadScrpt", cScript)

        End If

    End Sub
    Function GetSelectSeriaName()
        Dim result As String

        '  Response.Write("ss=" & ss)
        '  Response.End()
        If ss <> "" Then

            Dim cmdSelect As New SqlClient.SqlCommand("select  Series_Id,Series_Name from Series where Series_Id in (" & ss & ") ORDER BY Series_Name", conB)

            cmdSelect.CommandType = CommandType.Text
            conB.Open()

            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("Series_Name")) Then
                    If result <> "" Then
                        result = result & ", " & myReader("Series_Name")
                    Else
                        result = myReader("Series_Name")
                    End If


                End If

            End While
            If result = "" Then
                result = "הכל"
            End If
            conB.Close()

            Return result
        End If
    End Function
    Sub getSerias()

        Dim sql As String
        sql = "select  Series_Id,Series_Name as SeriaName from Series where 1=1 "

        If val_User <> "" And val_User <> "0" Then
            sql = sql & " and User_Id in (" & val_User & ")"
        End If
        sql = sql & "  ORDER BY Series_Name "
        ' Response.Write(sql)
        Dim cmdSelect As New SqlClient.SqlCommand(sql, conB)
        cmdSelect.CommandType = CommandType.Text
        conB.Open()
        fselect.DataTextField = "SeriaName"
        fselect.DataValueField = "Series_Id"
        fselect.DataSource = cmdSelect.ExecuteReader()
        fselect.DataBind()
        conB.Close()
        If fselect.Items.Count > 0 Then
            Dim list1 As New ListItem("הכל", "0")
            fselect.Items.Insert(0, list1)
        End If
        If val_SeriasSelect <> "" Then
            Dim parts As String() = val_SeriasSelect.Split(",")
            For Each part As String In parts
                fselect.Items.FindByValue(part).Selected = True
            Next
        End If


        '  val_SeriasSelect()
        'If fValue <> "" Then
        '    fselect.Items.FindByValue(fValue).Selected = True
        'End If

    End Sub
End Class
