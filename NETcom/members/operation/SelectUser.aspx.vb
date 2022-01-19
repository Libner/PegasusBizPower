Public Class SelectUser
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dtUsers As New DataTable
    Dim primKeydtUsers(0) As Data.DataColumn
    Dim myReader As System.Data.SqlClient.SqlDataReader
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Public Departure_GuideTelphone As String
    Protected WithEvents btnSubmit As UI.WebControls.Button
    Protected fName, fValue, DepartureCode, fTitle As String
    Protected WithEvents fselect As HtmlSelect
    Dim ss, ssName, val_UsersSelect As String

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
        val_UsersSelect = Request.QueryString("v")
        ' Response.Write("2=" & val_UsersSelect)
        '  Response.End()
      
      
        If Not Page.IsPostBack Then
         
       
          
            '  Response.Write(fName & ":" & Request.Form("fselect"))
            ' Response.End()
            ' Dim ss As String
            '  ss = getUsers(CInt(Request.Form("fselect")))
            getUsers()

        Else

            ss = Request.Form("fselect")
            ssName = GetSelectUserName()
            Dim cScript As String
            If ss = 0 Then
                cScript = "<script language='javascript'>opener.document.getElementById('UsersSelect').value = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; opener.document.getElementById('SelectUserAlt').src='../../images/select.png';opener.document.getElementById('SelectUserAlt').title='" & ssName & "';self.close(); </script>"

            Else
                cScript = "<script language='javascript'>opener.document.getElementById('UsersSelect').value = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; opener.document.getElementById('SelectUserAlt').src='../../images/select_act.png';opener.document.getElementById('SelectUserAlt').title='" & ssName & "';self.close(); </script>"

            End If
            RegisterStartupScript("ReloadScrpt", cScript)

        End If

    End Sub

    Function GetSelectUserName()
        Dim result As String

        '  Response.Write("ss=" & ss)
        '  Response.End()
        If ss <> "" Then
            Dim cmdSelect As New SqlClient.SqlCommand("SELECT USER_ID, FIRSTNAME + ' ' + LASTNAME as Username from Users  where  USER_ID in(" & ss & ") order by FIRSTNAME asc ,LASTNAME asc", conB)
            cmdSelect.CommandType = CommandType.Text
            conB.Open()

            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("Username")) Then
                    If result <> "" Then
                        result = result & ", " & myReader("Username")
                    Else
                        result = myReader("Username")
                    End If


                End If

            End While
            If result = "" Then
                result = "הכל"
            End If
            Return result
        End If
    End Function
    Sub getUsers()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT USER_ID, FIRSTNAME + ' ' + LASTNAME as Username from Users  where [ACTIVE]=1 and job_id=471 order by FIRSTNAME asc ,LASTNAME asc", conB)
        cmdSelect.CommandType = CommandType.Text
        conB.Open()
        fselect.DataTextField = "Username"
        fselect.DataValueField = "USER_ID"
        fselect.DataSource = cmdSelect.ExecuteReader()
        fselect.DataBind()
        conB.Close()
        If fselect.Items.Count > 0 Then
            Dim list1 As New ListItem("הכל", "0")
            fselect.Items.Insert(0, list1)
        End If
        If val_UsersSelect <> "" Then
            Dim parts As String() = val_UsersSelect.Split(",")
            For Each part As String In parts
                fselect.Items.FindByValue(part).Selected = True
            Next


            '  val_UsersSelect()

        End If

    End Sub
End Class
   