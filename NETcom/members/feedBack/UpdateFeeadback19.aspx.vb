Public Class UpdateFeeadback19
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Protected DepartureId, pId As Integer
    Dim cmdSelect As New SqlClient.SqlCommand

    Protected dateP, UserName As HtmlInputText
    Protected pContentText, pContentTextHeb As String
    Protected Userid As String
    Protected DepatureCode As String
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

    Protected ContentText As HtmlTextArea
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm


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
        Userid = Trim(Request.Cookies("bizpegasus")("UserId"))
        UserName.Value = func.GetUserNameEng(Userid)
        If IsNumeric(Request.QueryString("sdep")) Then
            DepartureId = CInt(Request.QueryString("sdep"))

            If DepartureId > 0 Then
                Dim sqlstr As String
                sqlstr = "SELECT Departure_Code, isGuaranteed  FROM Tours_Departures where Departure_Id=" & DepartureId
                Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
                Dim myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                Dim myReader = myCommand.ExecuteReader(CommandBehavior.SingleRow)
                While myReader.Read()
                    DepatureCode = myReader("Departure_Code")

                End While
                myConnection.Close()
            End If

            'If Page.IsPostBack Then
            If IsNumeric(Request.QueryString("id")) Then
                pId = Request.QueryString("id")
            End If

            GetData()
            If Page.IsPostBack Then
                SaveData()
            End If

        End If
    End Sub
    Public Sub GetData()
        'Response.Write("pId=" & pId)
        Dim cmdSelect As New SqlClient.SqlCommand("select Id,Date,Content_Text,Content_TextHeb,isNull(FIRSTNAME_ENG + ' ' + LASTNAME_ENG,'Automatic System') as Username from   FeedbackToSsupplier FS left join Users on FS.User_Id=Users.User_Id left join Suppliers on FS.supplier_Id=Suppliers.supplier_Id  where  Id = @pId order by Date,supplier_Name ", conB)
        cmdSelect.Parameters.Add("@pId", SqlDbType.Int).Value = CInt(pId)

        conB.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)

        While dr.Read()
            dateP.Value = dr("Date")
            UserName.Value = dr("Username")
            pContentText = dr("Content_Text")
            If Not IsDBNull(dr("Content_TextHeb")) Then
                pContentTextHeb = dr("Content_TextHeb")
            End If

        End While
        conB.Close()

    End Sub


    Public Sub SaveData()
        Dim UserId = Trim(Request.Cookies("bizpegasus")("UserId"))
        pContentText = Request.Form("ContentText")

        ''Update ''
        Dim cmdSelect = New SqlClient.SqlCommand("update FeedbackToSsupplier set User_Id=@UserId,Content_Text=@ContentText WHERE (Id = @pId)", conB)
        cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
        cmdSelect.Parameters.Add("@pId", SqlDbType.Int).Value = CInt(pId)
        cmdSelect.Parameters.Add("@ContentText", SqlDbType.NText).Value = pContentText
        conB.Open()
        cmdSelect.ExecuteNonQuery()
        conB.Close()


            '''----
            ''---  Send Mail

        

    
      
            Dim cScript As String

        cScript = "<script language='javascript'>alert('עדכון הערות לטבלת המשובים');window.opener.location.reload();self.close() ; </script>"
        RegisterStartupScript("ReloadScrpt", cScript)




    End Sub

End Class
