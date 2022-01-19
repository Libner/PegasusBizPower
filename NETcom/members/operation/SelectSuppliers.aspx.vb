Public Class SelectSuppliers
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dtUsers As New DataTable
    Dim primKeydtUsers(0) As Data.DataColumn
    Dim myReader As System.Data.SqlClient.SqlDataReader
    Dim cmdSelect As New SqlClient.SqlCommand
    ' Protected func As New include.funcs
    Public func As New bizpower.cfunc
    Public Departure_GuideTelphone As String
    Protected WithEvents btnSubmit As UI.WebControls.Button
    Protected fName, fValue, DepartureCode, fTitle As String
    Protected WithEvents fselect As HtmlSelect
    Dim ss, ssName, val_GuidesSelect As String
    Protected SeriasId As String

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
        val_GuidesSelect = Request.QueryString("v")
        ' Response.Write("2=" & val_GuidesSelect)
        '  Response.End()


        If Not Page.IsPostBack Then



            '  Response.Write(fName & ":" & Request.Form("fselect"))
            ' Response.End()
            ' Dim ss As String
            '  ss = getUsers(CInt(Request.Form("fselect")))
            getSuppliers()

        Else

            ss = Request.Form("fselect")
            ssName = GetSelectSuppliersName()
            Dim cScript As String
            If ss = 0 Then
                cScript = "<script language='javascript'>opener.document.getElementById('SuppliersSelect').value = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; opener.document.getElementById('SelectSuppliersAlt').src='../../images/select.png';opener.document.getElementById('SelectSuppliersAlt').title='" & ssName & "';self.close(); </script>"

            Else
                cScript = "<script language='javascript'>opener.document.getElementById('SuppliersSelect').value = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; opener.document.getElementById('SelectSuppliersAlt').src='../../images/select_act.png';opener.document.getElementById('SelectSuppliersAlt').title='" & ssName & "';self.close(); </script>"

            End If
            RegisterStartupScript("ReloadScrpt", cScript)

        End If

    End Sub
    Function GetSelectSuppliersName()
        Dim result As String

        '  Response.Write("ss=" & ss)
        '  Response.End()
        If ss <> "" Then
            Dim cmdSelect As New SqlClient.SqlCommand("SELECT supplier_Id, supplier_Name from Suppliers  where  supplier_Id in(" & ss & ") order by supplier_Name ", conB)
            cmdSelect.CommandType = CommandType.Text
            conB.Open()

            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("supplier_Name")) Then
                    If result <> "" Then
                        result = result & ", " & myReader("supplier_Name")
                    Else
                        result = myReader("supplier_Name")
                    End If


                End If

            End While
            conB.Close()
            If result = "" Then
                result = "הכל"
            End If
            Return result
        End If
    End Function
    Sub getSuppliers()
        Dim str As String

        ' Dim cmdSelect As New SqlClient.SqlCommand("SELECT supplier_Id, supplier_Name from Suppliers  order by supplier_Name ", conB)
        If Request.Cookies("bizpegasus")("Chief") <> "1" Then
            SeriasId = func.GetSeriasId(Request.Cookies("bizpegasus")("UserId"))
            If SeriasId <> "" Then
                str = "select supplier_Id,supplier_Name,Country_Id from Suppliers " & _
                  " where country_id in (select C.Country_Id from pegasus.dbo.Tours  T left join pegasus.dbo.Tours_Countries C on T.Tour_Id=C.Tour_Id " & _
                  " where SeriasId in(select SeriasId from Series  where user_id=" & Request.Cookies("bizpegasus")("UserId") & "))"
            Else
                str = "SELECT  supplier_Id,supplier_Name " & _
                    "FROM Suppliers order by supplier_Name"
            End If

            'Dim cmdSelect As New SqlClient.SqlCommand("select supplier_Id,Country_Id from Suppliers " & _
            '  " where country_id in (select C.Country_Id from pegasus.dbo.Tours  T left join pegasus.dbo.Tours_Countries C on T.Tour_Id=C.Tour_Id " & _
            '  " where SeriasId in(select SeriasId from Series  where user_id=" & Request.Cookies("bizpegasus")("UserId") & "))", conB)
            'Response.Write("uid=" & Request.Cookies("bizpegasus")("UserId"))
            'Response.End()



        Else
                str = "SELECT  supplier_Id,supplier_Name " & _
                       "FROM Suppliers order by supplier_Name"
                ' Dim cmdSelect As New SqlClient.SqlCommand("SELECT  supplier_Id,supplier_Name " & _
                '        "FROM Suppliers order by supplier_Name", conB)

            End If
            Dim cmdSelect As New SqlClient.SqlCommand(str, conB)

            'Dim cmdSelect As New SqlClient.SqlCommand("SELECT  supplier_Id,supplier_Name " & _
            '           "FROM Suppliers order by supplier_Name", conB)
            cmdSelect.CommandType = CommandType.Text
            conB.Open()
            fselect.DataTextField = "supplier_Name"
            fselect.DataValueField = "supplier_Id"
            fselect.DataSource = cmdSelect.ExecuteReader()
            fselect.DataBind()
            conB.Close()
            If fselect.Items.Count > 0 Then
                Dim list1 As New ListItem("הכל", "0")
                fselect.Items.Insert(0, list1)
            End If
            If val_GuidesSelect <> "" Then
                Dim parts As String() = val_GuidesSelect.Split(",")
                For Each part As String In parts
                    fselect.Items.FindByValue(part).Selected = True
                Next


                '  val_GuidesSelect()

            End If

    End Sub
End Class
