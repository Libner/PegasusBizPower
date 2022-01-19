Public Class SelectTours
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dtTours As New DataTable
    Dim primKeydtTours(0) As Data.DataColumn
    Dim myReader As System.Data.SqlClient.SqlDataReader
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Public Departure_GuideTelphone As String
    Protected WithEvents btnSubmit As UI.WebControls.Button
    Protected fName, fValue, DepartureCode, fTitle As String
    Protected WithEvents fselect As HtmlSelect
    Dim ss, ssName, val_ToursSelect As String

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
        val_ToursSelect = Request.QueryString("v")
        ' Response.Write("2=" & val_ToursSelect)
        '  Response.End()


        If Not Page.IsPostBack Then



            '  Response.Write(fName & ":" & Request.Form("fselect"))
            ' Response.End()
            ' Dim ss As String
            '  ss = getTours(CInt(Request.Form("fselect")))
            getTours()

        Else

            ss = Request.Form("fselect")
            ssName = GetSelectToursName()
            Dim cScript As String
            If ss = 0 Then
                cScript = "<script language='javascript'>opener.document.getElementById('ToursSelect').value = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; opener.document.getElementById('SelectToursAlt').src='../../images/select.png';opener.document.getElementById('SelectToursAlt').title='" & ssName & "';self.close(); </script>"

            Else
                cScript = "<script language='javascript'>opener.document.getElementById('ToursSelect').value = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; opener.document.getElementById('SelectToursAlt').src='../../images/select_act.png';opener.document.getElementById('SelectToursAlt').title='" & ssName & "';self.close(); </script>"

            End If
            RegisterStartupScript("ReloadScrpt", cScript)

        End If

    End Sub
    Function GetSelectToursName()
        Dim result As String

        '  Response.Write("ss=" & ss)
        '  Response.End()
        If ss <> "" Then
            Dim cmdSelect As New SqlClient.SqlCommand("SELECT Tour_Id, Tour_Name=case when isnull(Tour_PageTitle,'')='' then Tour_Name else Tour_PageTitle end from Tours  where  Tour_ID in(" & ss & ") order by Tour_Name", conPegasus)


            '    Dim cmdSelect As New SqlClient.SqlCommand("SELECT Tour_Id, Tour_Name from Tours  where  Tour_ID in(" & ss & ") order by Tour_Name", conPegasus)



            cmdSelect.CommandType = CommandType.Text
            conPegasus.Open()

            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("Tour_Name")) Then
                    If result <> "" Then
                        result = result & ", " & myReader("Tour_Name")
                    Else
                        result = myReader("Tour_Name")
                    End If


                End If

            End While
            conPegasus.Close()
            If result = "" Then
                result = "הכל"
            End If
            Return result
        End If
    End Function
    Sub getTours()
        Dim cmdSelect As New SqlClient.SqlCommand("set dateformat dmy;SELECT distinct  Tours.Tour_Id ,Tour_Name=case when isnull(Tour_PageTitle,'')='' then Tour_Name else Tour_PageTitle end " & _
               " FROM   Tours_Departures		INNER JOIN Tours ON Tours_Departures.Tour_Id = Tours.Tour_Id 		where	DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End) <= -1 " & _
               " order by  Tour_Name,Tours.Tour_Id ", conPegasus)
        'Dim cmdSelect As New SqlClient.SqlCommand("set dateformat dmy;SELECT distinct  Tours.Tour_Id ,Tour_Name " & _
        '          " FROM   Tours_Departures		INNER JOIN Tours ON Tours_Departures.Tour_Id = Tours.Tour_Id 		where	DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End) <= -1 " & _
        '          " order by Tours.Tour_Id desc", conPegasus)


        cmdSelect.CommandType = CommandType.Text
        conPegasus.Open()
        fselect.DataTextField = "Tour_Name"
        fselect.DataValueField = "Tour_Id"
        fselect.DataSource = cmdSelect.ExecuteReader()
        fselect.DataBind()
        conPegasus.Close()
        If fselect.Items.Count > 0 Then
            Dim list1 As New ListItem("הכל", "0")
            fselect.Items.Insert(0, list1)
        End If
        If val_ToursSelect <> "" Then
            Dim parts As String() = val_ToursSelect.Split(",")
            For Each part As String In parts
                fselect.Items.FindByValue(part).Selected = True
            Next


            '  val_ToursSelect()

        End If

    End Sub
End Class
