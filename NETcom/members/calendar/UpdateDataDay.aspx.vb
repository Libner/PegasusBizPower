Imports System.Data.SqlClient

Public Class UpdateDataDay
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmd As New SqlClient.SqlCommand
    Public parName, selV, selVName, selDateTypeId, DayTypeColor, DayFontColor As String
    Public isHoliday, wdayNumber As Integer




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
        parName = Request.Form("name")
        selV = Request.Form("selV")
        isHoliday = Request.Form("isH")
        wdayNumber = Request.Form("wday")
        Select Case selV
            Case "1"
                selVName = "יום עבודה מלא"
                selDateTypeId = "1"
            Case "2"
                selVName = "חצי יום עבודה"
                selDateTypeId = "0.5"
            Case "3"
                selVName = "אין עבודה"
                selDateTypeId = "0"
        End Select
        If wdayNumber = 7 Then
            DayTypeColor = "#ff0000"
            DayFontColor = "#ffffff"
        Else
            If isHoliday = "1" Then
                Select Case selV
                    Case "1"
                        DayTypeColor = "#008000"
                        DayFontColor = "#ffffff"
                    Case "2"
                        DayTypeColor = "#FFCE42"
                        DayFontColor = "#000000"
                    Case "3"
                        DayTypeColor = "#F16529"
                        DayFontColor = "#000000"
                End Select
            Else
                Select Case selV
                    Case "1"
                        DayTypeColor = "#7CFC00"
                        DayFontColor = "#000000"
                    Case "2"
                        DayTypeColor = "#FFC0CB"
                        DayFontColor = "#000000"
                    Case "3"
                        DayTypeColor = "#ff0000"
                        DayFontColor = "#ffffff"
                End Select

            End If
        End If



        con.Open()

        cmd = New System.Data.SqlClient.SqlCommand("Update DimDate set DayTypeName=@selVName,DayTypeId=@selDateTypeId,DayTypeColor=@DayTypeColor,DayFontColor=@DayFontColor where DateKey=@parName", con)
        cmd.Parameters.Add("@selVName", SqlDbType.VarChar).Value = selVName
        cmd.Parameters.Add("@selDateTypeId", SqlDbType.VarChar).Value = selDateTypeId
        cmd.Parameters.Add("@parName", SqlDbType.DateTime).Value = CDate(parName)
        cmd.Parameters.Add("@DayTypeColor", SqlDbType.VarChar).Value = DayTypeColor
        cmd.Parameters.Add("@DayFontColor", SqlDbType.VarChar).Value = DayFontColor

        cmd.CommandType = CommandType.Text
        ' cmd.Parameters.Add("@chkValue", SqlDbType.Int).Value = CInt(chk.Value())
        'cmd.Parameters.Add("@ProfId", SqlDbType.Int).Value = CInt(chk.Attributes("value"))
        cmd.ExecuteNonQuery()
        con.Close()
        Response.Write(selVName & "_" & DayTypeColor & "_" & DayFontColor)
    End Sub

End Class
