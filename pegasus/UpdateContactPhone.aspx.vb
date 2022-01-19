Public Class UpdateContactPhone
    Inherits System.Web.UI.Page


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
        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        Dim cmdSelect As New SqlClient.SqlCommand("SELECT CONTACT_ID, cellular FROM CONTACTS", con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        cmdSelect.ExecuteReader()


        '1:      .UPDATE(pegasus.dbo.Members)
        'SET pegasus.dbo.Members .Cellular = C.Cellular 
        'FROM bizpower_pegasus.dbo.Contacts C
        '        where(pegasus.dbo.Members.Contact_Id = C.Contact_Id And C.Contact_id = 55026)


        '2:      .UPDATE(pegasus.dbo.Members)
        '        pegasus.dbo.Members.Cellular = C.Cellular
        'FROM bizpower_pegasus.dbo.Contacts C
        '        where(pegasus.dbo.Members.Contact_Id = C.Contact_Id and   M.Cellular<>C.cellular
        '3.

        '        UPDATE(pegasus.dbo.Members)
        'SET pegasus.dbo.Members .Cellular = C.Cellular 
        'FROM bizpower_pegasus.dbo.Contacts C
        '        where(pegasus.dbo.Members.Contact_Id = C.Contact_Id And M.Cellular <> C.cellular)
        '4.
        '        select M.cellular,C.cellular,M.Contact_id,M.Member_Id
        'from pegasus.dbo.Members M left join bizpower_pegasus.dbo.CONTACTS C
        'on M.Contact_Id=C.Contact_Id
        'where
        ' M.Cellular<>C.cellular
        '5791 rows



        '5.  UPDATE pegasus.dbo.Members
        'SET pegasus.dbo.Members.Cellular = C.Cellular 
        'FROM bizpower_pegasus.dbo.Contacts C
        '        where
        '		pegasus.dbo.Members.Contact_Id = C.Contact_Id And pegasus.dbo.Members.cellular <> C.cellular


        cmdSelect.Dispose()
        con.Close()

    End Sub


End Class
