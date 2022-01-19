Public Class editVouchers_Provider_Status
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

    Dim cmdSelect As New SqlClient.SqlCommand
    Public func As New bizpower.cfunc

    Public Departure_GuideTelphone As String
    Protected WithEvents btnSubmit, btnCancel As UI.WebControls.Button
    Protected fName, fValue, DepartureCode, fTitle As String
    Protected WithEvents dtlSupplier As DataList

    Protected values As Array

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


        If IsNumeric(Request.QueryString("DepartureId")) Then
            DepartureId = CInt(Request.QueryString("DepartureId"))
        End If
        fName = Request.QueryString("fName")
        DepartureCode = Request.QueryString("DepartureCode")

     

        If Not Page.IsPostBack Then
            If DepartureId > 0 Then
                cmdSelect = New SqlClient.SqlCommand("SELECT supplier_Id FROM Tours_Departures " & _
                " WHERE (Departure_Id = @Departure_Id)", conPegasus)
                cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
                conPegasus.Open()
                Try
                    fValue = cmdSelect.ExecuteScalar
                Catch ex As Exception

                End Try
                conPegasus.Close()
                If fValue = "" Then
                    fValue = 0
                End If
                values = fValue.Split(",")
                'Response.Write(values.Length())


                getSuppliers()
            End If
        Else
            'Dim i As Integer
            'For i = 0 To Request.Form().Count - 1
            '    Response.Write(Request.Form(i) & "=" & Request.Form(i).GetType.Name & "<BR>")

            'Next
            Dim Item
            Dim strArr() As String

            For Each Item In Request.Form
                ' Response.Write(Item & "-" & InStr(Item, "Vouchers_Provider", 0) & "-")

                If InStr(Item, "Vouchers_Provider", 0) > 0 Then
                    ' Response.Write(Item & "-" & Request.Form(Item))
                    strArr = Item.Split("_")
                    '  Response.Write("strArr=" & strArr(2))
                    cmdSelect = New SqlClient.SqlCommand("update VouchersToSuppliers set  Vouchers_Status=' " & Request.Form(Item) & "' where  supplier_Id = @supplier_Id and  (Departure_Id = @Departure_Id)", conB)
                    cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
                    cmdSelect.Parameters.Add("@supplier_Id", SqlDbType.Int).Value = CInt(strArr(2))

                    conB.Open()
                    cmdSelect.ExecuteNonQuery()
                    conB.Close()

                End If

            Next
            ' Response.Write("rrrr")

            '     Response.Write(Request.Form("sVouchers_Provider_3"))
            ' Response.End()


            '////------------update VouchersToSuppliers

            'Dim chk, chkIsMain As HtmlInputCheckBox
            ''''Dim i As Integer, sId As String
            ''''Dim txtField As String
            ''''', sId As String
            '''''sId = ""
            ''''For i = 0 To dtlSupplier.Items.Count - 1
            ''''    '  sId = dtlSupplier.Items(i).DataItem("supplier_Id")
            ''''    ' Response.Write(sId)
            ''''    ' Response.End()
            ''''    Dim txtSel As HtmlSelect
            ''''    txtSel = CType(dtlSupplier.Items(i).FindControl("sVouchers_Provider"), HtmlSelect)
            ''''    txtField = txtSel.Items(txtSel.SelectedIndex).Value 'dtlSupplier.Items(i).FindControl("sVouchers_Provider")
            ''''    Response.Write(txtField)
            ''''    Response.End()
            ''''    '    chk = CType(dtlSupplier.Items(i).FindControl("chkSpec"), HtmlInputCheckBox)
            ''''    '    If chk.Checked Then
            ''''    '        cmdSelect = New SqlClient.SqlCommand("Ins_VouchersToSuppliers", conB)
            ''''    '        cmdSelect.CommandType = CommandType.StoredProcedure
            ''''    '        cmdSelect.Parameters.Add("@supplier_Id", SqlDbType.Int).Value = CInt(chk.Attributes("value"))
            ''''    '        cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)

            ''''    '        conB.Open()
            ''''    '        cmdSelect.ExecuteNonQuery()
            ''''    '        conB.Close()

            ''''    '        'If sId = "" Then
            ''''    '        '    sId = CInt(chk.Attributes("value"))
            ''''    '        'Else
            ''''    '        '    sId = sId & "," & CInt(chk.Attributes("value"))
            ''''    '        'End If

            ''''    '    End If
            ''''Next
            '''''  Response.Write("1=" & sId)
            ''''' Response.End()  ''''''''''''
            ''''''delete --
            '''''  Response.Write("sId=" & sId)
            ''''' Response.End()
            '''''cmdSelect = New SqlClient.SqlCommand("Del_VouchersToSuppliers", conB)
            '''''cmdSelect.CommandType = CommandType.StoredProcedure
            '''''cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            '''''cmdSelect.Parameters.Add("@supplier_Id", SqlDbType.VarChar, 150).Value = sId
            '''''conB.Open()
            '''''cmdSelect.ExecuteNonQuery()
            '''''conB.Close()

            ''''''delete---



            ''cmdSelect = New SqlClient.SqlCommand("update Tours_Departures set supplier_Id = '" & sId & "' WHERE (Departure_Id = @Departure_Id)", conPegasus)
            ''cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
            ''conPegasus.Open()
            ''cmdSelect.ExecuteNonQuery()
            ''conPegasus.Close()
            '''  Response.Write(fName & ":" & Request.Form("fselect"))
            ''' Response.End()
            ''Dim ss As String
            '''  ss = func.GetSelectSupplierName(sId)
            ''ss = sId


            Dim cScript As String
            ''If ss <> "" Then
            ''    '  cScript = "<script language='javascript'>opener.document.getElementById('supplier_Id_" & DepartureId & "').innerHTML = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; self.close(); </script>"
            ''    cScript = "<script language='javascript'>opener.document.getElementById('Imgsupplier_Id_row" & DepartureId & "').src='../../images/v.png';self.close(); </script>"
            ''    'Response.Write(cScript)
            ''    '  Response.End()
            ''Else
            ''    cScript = "<script language='javascript'>opener.document.getElementById('Imgsupplier_Id_row" & DepartureId & "').src='../../images/select.png';opener.document.getElementById('Imgsupplier_Id_row" & DepartureId & "').title='';self.close(); </script>"

            ''End If
            Dim status = func.GetVouchers_ProviderStatus(DepartureId)

            cScript = "<script language='javascript'>opener.document.getElementById('Vouchers_Provider_" & DepartureId & "').innerHTML='" & status & "';self.close(); </script>"
            RegisterStartupScript("ReloadScrpt", cScript)

            ' cScript = "<script language='javascript'>opener.document.getElementById('Vouchers_Provider_" & DepartureId & "').innerHTML='uuuuu';self.close(); </script>"
            ' RegisterStartupScript("ReloadScrpt", cScript)
        End If

    End Sub
    Sub dtlSupplier_onDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataListItemEventArgs) Handles dtlSupplier.ItemDataBound
        'Set categories link for selected row in categories repeater
        ''If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
        ''    Dim chkSpec As HtmlInputCheckBox
        ''    Dim i As Integer
        ''    chkSpec = e.Item.FindControl("chkSpec")
        ''    chkSpec.Attributes("name") = "chkSpec"
        ''    chkSpec.Attributes.Add("value", e.Item.DataItem("supplier_Id"))
        ''    'Response.Write("fValue=" & fValue)
        ''    For i = 0 To values.Length() - 1
        ''        If chkSpec.Value = values(i) Then
        ''            chkSpec.Checked = True
        ''        End If
        ''    Next



        '  End If
        'Set categories link for selected row in categories repeater
   
    End Sub
    Public Function GetSupplierName(ByVal Sid As Integer) As String
        Dim sqlstr As String
        sqlstr = "SELECT supplier_Name  FROM    Suppliers where supplier_ID=" & Sid

        Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
        myConnection.Open()

        Dim SName = myCommand.ExecuteScalar
        myConnection.Close()
        Return SName


    End Function
    Sub getSuppliers()
        ''           Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id,S.supplier_Name,Departure_Id,Vouchers_Status ,Country_Name " & _
        ''     " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
        ''   " left join pegasus.dbo.Countries on pegasus.dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), conB)

        Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id,S.supplier_Name,Departure_Id,Vouchers_Status ,Country_Name " & _
           " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
           " left join pegasus.dbo.Countries on pegasus.dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), conB)


        cmdSelect.CommandType = CommandType.Text
        conB.Open()


        dtlSupplier.DataSource = cmdSelect.ExecuteReader()
        dtlSupplier.DataBind()
        conB.Close()

    End Sub


    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        cmdSelect = New SqlClient.SqlCommand("update Tours_Departures set supplier_Id = null WHERE (Departure_Id = @Departure_Id)", conPegasus)
        cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
        conPegasus.Open()
        cmdSelect.ExecuteNonQuery()
        conPegasus.Close()
        '  Response.Write(fName & ":" & Request.Form("fselect"))
        ' Response.End()


        Dim cScript As String

        '  cScript = "<script language='javascript'>opener.document.getElementById('supplier_Id_" & DepartureId & "').innerHTML = '" & Replace(ss, "'", "&#" & AscW(Chr(39)) & ";") & "'; self.close(); </script>"
        cScript = "<script language='javascript'>opener.document.getElementById('Imgsupplier_Id_row" & DepartureId & "').src='../../images/select.png';opener.document.getElementById('Imgsupplier_Id_row" & DepartureId & "').title='';self.close(); </script>"
        ' Response.Write(cScript)
        ' Response.End()

        RegisterStartupScript("ReloadScrpt", cScript)
    End Sub
End Class