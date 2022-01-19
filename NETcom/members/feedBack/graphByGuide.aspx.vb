Public Class graphByGuide
    Inherits System.Web.UI.Page
    Dim func As New bizpower.cfunc
    Protected guideId, currentYear, FromDate, ToDate, Guide_Name, res1, res2, Guide_Phone, Guide_Email, title, label2, label1 As String

    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim rs_Guide, rs_r As System.Data.SqlClient.SqlDataReader

    Dim sqlstrGuide As New SqlClient.SqlCommand

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

        guideId = Request("guide_id")
        currentYear = Request("currentYear")

        FromDate = "1/01/" & currentYear
        ToDate = "31/12/" & currentYear
        '	 response.Write FromDate &":"& ToDate
        '	 response.end
        conPegasus.Open()
        Dim sqlstrGuide As New SqlClient.SqlCommand("SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name,Guide_Phone,Guide_Email  FROM Guides  where Guide_Id=" & guideId, conPegasus)
        rs_Guide = sqlstrGuide.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_Guide.Read()

            If Not IsDBNull(rs_Guide("Guide_Name")) Then
                Guide_Name = Replace(rs_Guide("Guide_Name"), """", "''")
            End If
            If Not IsDBNull(rs_Guide("Guide_Phone")) Then
                Guide_Phone = rs_Guide("Guide_Phone")
            End If
            If Not IsDBNull(rs_Guide("Guide_Email")) Then
                Guide_Email = rs_Guide("Guide_Email")
            End If
        End While
        conPegasus.Close()
        title = "משובים דוח שנתי למדריך " & Guide_Name
        conPegasus.Open()
        Dim s = "Exec dbo.[get_GuideFeedbackReport]  '" & guideId & "','" & FromDate & "','" & ToDate & "'"
        'Response.Write(s)
        '' Response.End()
        Dim sql As New SqlClient.SqlCommand("Exec dbo.[get_GuideFeedbackReport]  '" & guideId & "','" & FromDate & "','" & ToDate & "'", conPegasus)
        rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_r.Read()
            res1 = rs_r("Tour_Grade")

            res2 = rs_r("Guide_Grade")
            label2 = label2 & "{label: """ & func.QFix(rs_r("Departure_Code")) & " / מס משובים " & rs_r("CountFeedBack") & """, y: " & res2 & ",indexLabel: """ & res2 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            label1 = label1 & "{label: """ & func.QFix(rs_r("Departure_Code")) & """, y:" & res1 & ",indexLabel: """ & res1 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"


        End While
        conPegasus.Close()





        'If Guide_Phone <> "" Then
        '    strXLS.Add("<tr style='background:#808080;color:#ffffff'><td colspan=6 align=right class=xl24 style='background:#808080;color:#ffffff'>") : strXLS.Add("טלפון:") : strXLS.Add(Guide_Phone) : strXLS.Add(" </TD></TR>")
        'End If
        'If Guide_Email <> "" Then
        '    strXLS.Add("<tr style='background:#808080;color:#ffffff'><td colspan=6 align=right class=xl24 style='background:#808080;color:#ffffff'>") : strXLS.Add(Guide_Email) : strXLS.Add(" </TD></TR>")
        'End If




    End Sub

End Class
