Public Class ReportGraph
    Inherits System.Web.UI.Page
    Dim func As New bizpower.cfunc
    Protected guideId, currentYear, FromDate, ToDate, makor, res1, res2, Guide_Phone, Guide_Email, title, label2, label1 As String

    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim rs_Makor As System.Data.SqlClient.SqlDataReader

    Dim sqlstrMakor As New SqlClient.SqlCommand

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
        currentYear = Request("currentYear")

        FromDate = DateValue(Trim(Request("dateStart")))
        ToDate = DateValue(Trim(Request("dateEnd")))
        '	 response.Write FromDate &":"& ToDate
        '	 response.end
        con.Open()
        '
       Dim sqlstrMakor As New SqlClient.SqlCommand("SET DATEFORMAT DMY;SELECT distinct FIELD_VALUE from  FORM_VALUE " & _
        "  Left Join APPEALS on APPEALS.APPEAL_ID=FORM_VALUE.APPEAL_ID    where " & _
        " (FORM_VALUE.FIELD_ID = 40812 OR  FORM_VALUE.FIELD_ID = 40776) AND (LTRIM(RTRIM(FORM_VALUE.FIELD_VALUE)) <> '') " & _
        " and (APPEAL_DATE BETWEEN COALESCE('" & FromDate & "',APPEAL_DATE) AND  COALESCE('" & ToDate & "',APPEAL_DATE)) ", con)
        rs_Makor = sqlstrMakor.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_Makor.Read()

            If Not IsDBNull(rs_Makor("FIELD_VALUE")) Then
                makor = rs_Makor("FIELD_VALUE")
                res1 = 10
                label1 = label1 & "{label: """ & rs_Makor("FIELD_VALUE") & """, y:" & res1 & ",indexLabel: """ & res1 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            End If

        End While
        con.Close()

     




        'If Guide_Phone <> "" Then
        '    strXLS.Add("<tr style='background:#808080;color:#ffffff'><td colspan=6 align=right class=xl24 style='background:#808080;color:#ffffff'>") : strXLS.Add("טלפון:") : strXLS.Add(Guide_Phone) : strXLS.Add(" </TD></TR>")
        'End If
        'If Guide_Email <> "" Then
        '    strXLS.Add("<tr style='background:#808080;color:#ffffff'><td colspan=6 align=right class=xl24 style='background:#808080;color:#ffffff'>") : strXLS.Add(Guide_Email) : strXLS.Add(" </TD></TR>")
        'End If




    End Sub

End Class
