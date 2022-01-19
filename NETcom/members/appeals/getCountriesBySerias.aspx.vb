Imports System.Data.SqlClient
Public Class getCountriesBySerias
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
        'cmdSelect = New System.Data.SqlClient.SqlCommand("SELECT  distinct SeriasId,Series_Name FROM dbo.Tours T " & _
        '                                           "left join " & BizpowerDBName & ".dbo.Series S on T.SeriasId=S.Series_Id   where T.Country_CRM=@COUNTRYID  ORDER BY  Series_Name", conPegasus)
        Dim strsql As String
        If func.fixNumeric(SERID) > 0 Then
            strsql = "SELECT distinct Country_Id, Country_Name from " & BizpowerDBName & ".dbo.Countries " & _
        " inner join dbo.Tours T on T.Country_CRM=Countries.Country_Id   where T.SeriasId=@SeriasId order by Country_Name"
        Else
            strsql = "SELECT distinct Country_Id, Country_Name from " & BizpowerDBName & ".dbo.Countries " & _
        " inner join dbo.Tours T on T.Country_CRM=Countries.Country_Id  order by Country_Name"
        End If
        cmdSelect = New SqlClient.SqlCommand(strsql, conPegasus)
        If SERID > 0 Then
            cmdSelect.Parameters.Add("@SeriasId", SqlDbType.Int).Value = CInt(SERID) 'where serias id=
        End If
        Dim tblCountries As New DataTable
        Dim adCountries As New SqlDataAdapter
        adCountries.SelectCommand = cmdSelect
        conPegasus.Open()

        adCountries.Fill(tblCountries)

        conPegasus.Close()

        selStr.Append("<option value=""0""></option>")
        For dr As Integer = 0 To tblCountries.Rows.Count - 1

            '  If dr("Departure_id") = TOURID Then
            'selStr.Append("<option value=""" & dr("Departure_Id") & """ selected>" & dr("Departure_Name") & "</option>")

            'Else
            If tblCountries.Rows.Count = 1 Or CInt(COUNTRYID) = CInt(tblCountries.Rows(dr)("Country_Id")) Then
                selStr.Append("<option value=""" & tblCountries.Rows(dr)("Country_Id") & """ selected>" & tblCountries.Rows(dr)("Country_Name") & "</option>")
            Else
                selStr.Append("<option value=""" & tblCountries.Rows(dr)("Country_Id") & """>" & tblCountries.Rows(dr)("Country_Name") & "</option>")
            End If

            '  End If
        Next
        Response.Write(selStr)

    End Sub




End Class