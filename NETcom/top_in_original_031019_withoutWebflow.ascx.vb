Public Class top_in
    Inherits System.Web.UI.UserControl
    Public numOfLink As String = "0"
    Public numOfTab As String = "0"
    Dim arr_bars_top(3, 1) As String
    Dim arr_bars(5, 1) As String
    Public strLocal As String
    Dim top_bar_count, bar_count, i, count, j As Integer
    Dim sql_obj, sqlstr, parent_id, parent_order, userId, orgID, lang_id, barID, barTitle, _
    barVisible, barParent, barURL, href, dis As String
    Dim links_array() As String
    Public dir_var, align_var, dir_obj_var As String
    Dim strPH As String
    Dim rs_parents As System.Data.SqlClient.SqlDataReader
    Dim rs_bar As System.Data.SqlClient.SqlDataReader
    Dim rs_obj As System.Data.SqlClient.SqlDataReader
    Protected WithEvents conn As System.Data.SqlClient.SqlConnection
    Protected WithEvents conn1 As System.Data.SqlClient.SqlConnection
    Protected WithEvents BarTopPH As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents BarBottomPH As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents catCMD As System.Data.SqlClient.SqlCommand
    Protected func As New bizpower.cfunc

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
        strLocal = "http://" & Trim(Request.ServerVariables("SERVER_NAME"))
        If Len(Trim(Request.ServerVariables("SERVER_PORT"))) > 0 Then
            strLocal = strLocal & ":" & Trim(Request.ServerVariables("SERVER_PORT"))
        End If
        strLocal = strLocal & Application("VirDir") & "/"
        userId = Trim(Request.Cookies("bizpegasus")("UserId"))
        orgID = Trim(Request.Cookies("bizpegasus")("OrgID"))
        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))

        If IsNumeric(lang_id) = False Or Trim(lang_id) = "" Then
            lang_id = "1"
        End If
        If lang_id = "2" Then
            dir_var = "rtl" : align_var = "left" : dir_obj_var = "ltr"
        Else
            dir_var = "ltr" : align_var = "right" : dir_obj_var = "rtl"
        End If

        conn = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        conn1 = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        'sql_obj = "Select bar_id, bar_title, is_visible From dbo.bar_users_table('" & orgID & "','" & userId & "')" & _
        '" WHERE parent_id IS NULL Order By parent_order, bar_order"
        sql_obj = "Exec dbo.get_user_bars '" & orgID & "','" & userId & "',0"
        catCMD = New System.Data.SqlClient.SqlCommand(sql_obj, conn)
        conn.Open()
        rs_obj = catCMD.ExecuteReader(CommandBehavior.CloseConnection)
        top_bar_count = 0
        While rs_obj.Read()
            arr_bars_top(0, top_bar_count) = func.dbNullFix(rs_obj("bar_id"))
            arr_bars_top(1, top_bar_count) = func.dbNullFix(rs_obj("bar_title"))
            arr_bars_top(2, top_bar_count) = func.dbNullFix(rs_obj("is_visible"))

            top_bar_count = top_bar_count + 1
            ReDim Preserve arr_bars_top(3, top_bar_count)
        End While
        rs_obj.Close()

        'sql_obj = "Select bar_id, bar_title, is_visible, parent_id, bar_url, parent_order From " & _
        '" dbo.bar_users_table('" & orgID & "','" & userId & "')" & _
        '" WHERE parent_id IS NOT NULL Order By Parent_Order, bar_order"
        sql_obj = "Exec dbo.get_user_bars '" & orgID & "','" & userId & "',1"
  
        catCMD = New System.Data.SqlClient.SqlCommand(sql_obj, conn)
        conn.Open()
        rs_obj = catCMD.ExecuteReader(CommandBehavior.CloseConnection)
        bar_count = 0
        While rs_obj.Read()
            arr_bars(0, bar_count) = func.dbNullFix(rs_obj("bar_id"))
            arr_bars(1, bar_count) = func.dbNullFix(rs_obj("bar_title"))
            arr_bars(2, bar_count) = func.dbNullFix(rs_obj("is_visible"))
            arr_bars(3, bar_count) = func.dbNullFix(rs_obj("parent_order"))
            arr_bars(4, bar_count) = func.dbNullFix(rs_obj("bar_url"))
            bar_count = bar_count + 1
            ReDim Preserve arr_bars(5, bar_count)
        End While
        'Response.Write(arr_bars.)
        rs_obj.Close()


        ReDim Preserve links_array(top_bar_count)
        'sqlstr = "Select bar_id, bar_order FROM bars WHERE parent_id IS NULL"
        sqlstr = "Exec dbo.get_user_bars_links '" & orgID & "','" & userId & "'"
        '   Response.Write(sqlstr)
        '  Response.End()
        catCMD = New System.Data.SqlClient.SqlCommand(sqlstr, conn)
        conn.Open()
        rs_parents = catCMD.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_parents.Read()
            parent_id = func.dbNullFix(rs_parents("bar_id"))
            parent_order = func.dbNullFix(rs_parents("bar_order"))
            '     links_array(parent_order) = func.dbNullFix(rs_parents(2))
            sqlstr = "Select TOP 1 bar_url FROM dbo.bar_users_table('" & orgID & "','" & userId & "')" & _
            "  WHERE is_visible = '1' AND parent_id = " & parent_id & " Order BY bar_order"
            catCMD = New System.Data.SqlClient.SqlCommand(sqlstr, conn1)
            conn1.Open()
            rs_bar = catCMD.ExecuteReader(CommandBehavior.SingleRow)
            If rs_bar.Read() Then
                links_array(parent_order) = rs_bar(0)
            End If
            rs_bar.Close()
            conn1.Close()
        End While
        rs_parents.Close()

        ' creating top bar
        strPH = ""
        For i = 0 To UBound(arr_bars_top, 2) - 1
            barID = arr_bars_top(0, i)
            barTitle = arr_bars_top(1, i)
            barVisible = arr_bars_top(2, i)
            If Trim(barVisible) = "1" Then
                'href = "href='" & strLocal & CStr(links_array(i + 1)) & "' "
                href = "href='" & CStr(links_array(i + 1)) & "' "
                dis = ""

                strPH += "<td  align=center nowrap valign=top dir=" & dir_obj_var & ">"
                strPH += "<A class='link_top"
                If Trim(barID) = Trim(numOfTab + 1) Then
                    strPH += "_over"
                End If
                strPH += "' " & href & dis & "  target=_self>" & barTitle & "</A>"
                strPH += "</td><td width=5 nowrap></td>"
            End If
        Next
        If strPH = "" Then
            strPH = "<td></td>"
        End If
        BarTopPH.Controls.Add(New LiteralControl(strPH))

        'creating bottom bar
        strPH = ""
        count = 0
        For j = 0 To UBound(arr_bars, 2) - 1
            barID = arr_bars(0, j)
            barTitle = arr_bars(1, j)
            barVisible = arr_bars(2, j)
            barParent = arr_bars(3, j)
            barURL = arr_bars(4, j)
            If CInt(barParent) = CInt(numOfTab) + 1 Then
                If Trim(barVisible) = "1" Then
                    ' href = "href='" & strLocal & barURL & "' "
                    href = "href='" & barURL & "' "
                    dis = ""
                    strPH += "<TD bgcolor=#F0C000 nowrap dir=" & dir_obj_var & ">"
                    strPH += "<a " & href & dis & " target=_self class='title_tab"
                    If Trim(CStr(count)) = Trim(CStr(numOfLink)) Then
                        strPH += "_over"
                    End If
                    strPH += "'>" & barTitle & "</a></TD>"
                    strPH += "<TD width=2 bgcolor=#FFFFFF nowrap style='border-bottom: solid 1px #6F6D6B;'>&nbsp;</TD>"
                    count += 1
                End If
            End If
        Next
        If strPH = "" Then
            strPH = "<td></td>"
        End If
        BarBottomPH.Controls.Add(New LiteralControl(strPH))

    End Sub

End Class
