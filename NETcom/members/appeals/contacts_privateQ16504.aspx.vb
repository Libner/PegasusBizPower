Imports System.Data.SqlClient
Public Class contacts_privateQ16504
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")

    Public contactID, cont_name, where_cont_name, urlSort, sort As String
    Public companyID, UserID, OrgID, quest_id As String
    Public lang_id As Integer
    Public dir_var, align_var, dir_obj_var As String
    Public arrTitles As New DataTable
    Protected WithEvents rptContacts As Repeater
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        UserID = Trim(Request.Cookies("bizpegasus")("UserID"))

        OrgID = Trim(Trim(Request.Cookies("bizpegasus")("ORGID")))

        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))

        If func.fixNumeric(lang_id) = 0 Then
            lang_id = 1
        End If
        If lang_id = 2 Then
            dir_var = "rtl"
            align_var = "left"
            dir_obj_var = "ltr"
        Else
            dir_var = "ltr"
            align_var = "right"
            dir_obj_var = "rtl"

        End If
        Dim cmdSelect As New SqlClient.SqlCommand
        Dim strsql As String
     
            strsql = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 3 Order By word_id"
            cmdSelect = New SqlClient.SqlCommand(strsql, con)
            Dim adTitles As New SqlDataAdapter
            adTitles.SelectCommand = cmdSelect
            con.Open()

            adTitles.Fill(arrTitles)

            con.Close()
            Dim dtnewRow As DataRow = arrTitles.NewRow()

            arrTitles.Rows.InsertAt(dtnewRow, 0)

            quest_id = Trim(Request.QueryString("quest_id"))
            If Trim(Request("cont_name")) <> "" Or Trim(Request.QueryString("cont_name")) <> "" Then
                cont_name = (Trim(Request("cont_name")))
                cont_name = HttpUtility.UrlDecode(cont_name, System.Text.Encoding.GetEncoding("utf-8"))

                where_cont_name = " and UPPER(ltrim(rtrim(CONTACT_NAME))) LIKE '" & UCase(func.sFix(cont_name)) & "%'"
            Else
                where_cont_name = ""
            End If
     
        urlSort = "contacts.asp?cont_name=" & cont_name & "&quest_id=" & quest_id
        sort = Request.QueryString("sort")
        If Trim(sort) = "" Then
            sort = 0
        End If
        Dim sortby(8)
        sortby(0) = "rtrim(ltrim(CONTACT_NAME))"
        sortby(1) = "rtrim(ltrim(company_name))"
        sortby(2) = "rtrim(ltrim(company_name)) DESC"
        sortby(3) = "date_update"
        sortby(4) = "date_update DESC"
        sortby(5) = "CONTACT_NAME"
        sortby(6) = "CONTACT_NAME DESC"

        If func.dbNullFix(Request("cont_name")) <> "" Then

            strsql = "SELECT contacts.contact_id,contacts.CONTACT_NAME,contacts.company_ID,company_name," & _
  " contacts.phone,contacts.cellular,contacts.messanger_name,contacts.email FROM contacts " & _
  " Inner Join Companies On contacts.company_id = companies.company_id " & _
  " WHERE contacts.ORGANIZATION_ID = " & Trim(OrgID) & " AND isNULL(private_flag,0) = '1' " & _
   where_cont_name & " order by " & sortby(sort)
            'Response.Write("<br>" & strsql)
            con.Open()
            cmdSelect = New SqlClient.SqlCommand(strsql, con)
            rptContacts.DataSource = cmdSelect.ExecuteReader
            rptContacts.DataBind()
            con.Close()
        End If
    End Sub




End Class