Imports System.Data.SqlClient
Public Class getSeriasByCountry
    Inherits System.Web.UI.Page

    Protected COUNTRYID, SERID As String
    Protected func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    Public selStr As New System.Text.StringBuilder("")


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        SERID = 0
        If Request("SERID") <> "null" And Request("SERID") <> "" Then
            If IsNumeric(Trim(Request("SERID"))) Then
                If Trim(Request("SERID")) <> "0" Then
                    SERID = Trim(Request("SERID"))
                End If
            End If
        End If

        COUNTRYID = 0

        If Request("COUNTRYID") <> "null" And Request("COUNTRYID") <> "" Then
            If IsNumeric(Trim(Request("COUNTRYID"))) Then
                If Trim(Request("COUNTRYID")) <> "0" Then
                    COUNTRYID = CInt(Request("COUNTRYID"))
                End If
            End If
        End If

        Dim cmdSelect As New SqlClient.SqlCommand
        Dim strsql As String
        If func.fixNumeric(COUNTRYID) > 0 Then
            strsql = "SELECT  distinct SeriasId,Series_Name FROM dbo.Tours T " & _
                                                   "inner join " & BizpowerDBName & ".dbo.Series S on T.SeriasId=S.Series_Id where T.Country_CRM=@COUNTRYID and T.SeriasId is not null ORDER BY  Series_Name"
        Else
            strsql = "SELECT  distinct SeriasId,Series_Name FROM dbo.Tours T " & _
                                                    "inner join " & BizpowerDBName & ".dbo.Series S on T.SeriasId=S.Series_Id where T.SeriasId is not null  ORDER BY  Series_Name"
        End If

        cmdSelect = New System.Data.SqlClient.SqlCommand(strsql, conPegasus)
        If COUNTRYID > 0 Then
            cmdSelect.Parameters.Add("@COUNTRYID", SqlDbType.Int).Value = CInt(COUNTRYID) 'where serias id=
        End If
        Dim tblSerias As New DataTable
        Dim adSerias As New SqlDataAdapter
        adSerias.SelectCommand = cmdSelect
        conPegasus.Open()

        adSerias.Fill(tblSerias)

        conPegasus.Close()
        selStr.Append("<option value=""0""></option>")
        For dr As Integer = 0 To tblSerias.Rows.Count - 1
            '  If dr("Departure_id") = TOURID Then
            'selStr.Append("<option value=""" & dr("Departure_Id") & """ selected>" & dr("Departure_Name") & "</option>")

            'Else
            If tblSerias.Rows.Count = 1 Or CInt(SERID) = CInt(tblSerias.Rows(dr)("SeriasId")) Then
                selStr.Append("<option value=""" & tblSerias.Rows(dr)("SeriasId") & """ selected>" & tblSerias.Rows(dr)("Series_Name") & "</option>")
            Else
                selStr.Append("<option value=""" & tblSerias.Rows(dr)("SeriasId") & """>" & tblSerias.Rows(dr)("Series_Name") & "</option>")
            End If

        Next

        Response.Write(selStr)

    End Sub




End Class