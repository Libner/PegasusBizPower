Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml
Imports System.Web.HttpContext
Imports System.Net
Imports System.Configuration
Public Class getWorkersByDepId
    Inherits System.Web.UI.Page
    Protected func As New include.funcs
    Dim cmdSelect As SqlClient.SqlCommand
    Dim dr As SqlClient.SqlDataReader
    Dim depId As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))



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
       
        Dim reader As New StreamReader(Page.Request.InputStream)
        Dim xmlData As String = reader.ReadToEnd()
        Dim doc As New XmlDocument

        doc.LoadXml(xmlData)

        Dim root As XmlNode
        root = doc.DocumentElement


        If Not IsNothing(root.SelectSingleNode("depId").FirstChild.Value()) Then
            '  If Trim(root.SelectSingleNode("depId").FirstChild.Value()) <> "" Then
            depId = Trim(root.SelectSingleNode("depId").FirstChild.Value())
            '  End If
        End If

        Dim resp As String = ""
        depId = Left(depId, Len(depId) - 1)

        Dim sqlStr As String = "select  User_Id,Email,(FIRSTNAME + Char(32) + LASTNAME) as [USER_NAME] from Users where  len(email)>0 and ACTIVE=1 and Department_Id in (" & depId & ") order by USER_NAME"
        Dim cmdSelect As New System.Data.SqlClient.SqlCommand(sqlStr, con)
        cmdSelect.CommandType = CommandType.Text

        'cmdSelect.Parameters.Add("@depId", SqlDbType.VarChar, 500).Value = depId
        con.Open()
        dr = cmdSelect.ExecuteReader()
        If dr.HasRows Then
            While dr.Read()
                resp = resp & "#" & dr("USER_NAME") & ":" & dr("Email")
            End While
            If Len(resp) > 0 Then
                resp = resp.Substring(1)
            End If
        End If
        cmdSelect.Dispose()
        dr.Close()
        con.Close()
        Response.Write(resp)
    End Sub

End Class
