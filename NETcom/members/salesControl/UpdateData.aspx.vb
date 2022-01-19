Imports System.Data.SqlClient

Public Class UpdateData1
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmd As New SqlClient.SqlCommand
    Public parName, logUid, Uid, parValue, parDep As String




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
        Dim UserID = Trim(Request.Cookies("bizpegasus")("UserID"))


        parName = Request.Form("name")
        Uid = Request("Uid")
        logUid = Year(Request("Uid")) & "-" & Month(Request("Uid")) & "-" & Day(Request.Form("Uid"))
        Dim pM, pY, pD As String
        Dim strV As Array
        strV = logUid.Split("-")
        If strV.Length > 0 Then
            pM = strV(1)
            pY = strV(0)
            pD = strV(2)
        End If


        parValue = Request.Form("value")
        parDep = Request.Form("dep")
        'parName = "call_Sales"
        'Uid = "2017-09-01"
        'parValue = 10


        con.Open()
        'Dim sql = "Update DimDate set " & parName & "=" & parValue & " where DateKey=" & Uid
        'Response.Write(sql)

        cmd = New System.Data.SqlClient.SqlCommand("SET DATEFORMAT dmy;Update Departments_Data set " & parName & "=@parValue where DateKey=@Uid and departmentId=@depid", con)
        cmd.Parameters.Add("@parValue", SqlDbType.VarChar).Value = parValue
        cmd.Parameters.Add("@Uid", SqlDbType.VarChar).Value = Uid
        cmd.Parameters.Add("@depid", SqlDbType.VarChar).Value = parDep

        cmd.CommandType = CommandType.Text
        cmd.ExecuteNonQuery()
        con.Close()
        Dim Change_Type = "עדכון"
        Dim Table_ID = 111 '"שליטה חודשי / שנתי in table bars
        Dim Change_Table = "שליטה חודשי / שנתי"

        Dim ChangeSubTable As String
        Dim Object_Title As String

        If parName = "call_Sales" Then
            Object_Title = "שיחות מכירה נכנסות"
            ChangeSubTable = "טבלת מעקב ושליטה יומית"
        ElseIf parName = "call_Service" Then
            Object_Title = "שיחות שירות נכנסות"
            ChangeSubTable = "טבלת מעקב ושליטה יומית"
        End If
        ' Dim s = "SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO ChangesSalesControl (" & _
        '" Department_Id,Year,Month,Day,Change_Type,Table_ID,Change_Table,ChangeSubTable,Object_Title,User_ID,Change_Date,IP)" & _
        '" VALUES (" & parDep & "," & pY & "," & pM & "," & pD & ",'" & Change_Type & "'," & Table_ID & ",'" & Change_Table & "','" & ChangeSubTable & "','" & Object_Title & "'," & UserID & ",GetDate(),'" & Request.ServerVariables("REMOTE_ADDR") & "')"
        ' Response.Write(s)
        ' Response.End()


        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO ChangesSalesControl (" & _
       " Department_Id,Year,Month,Day,Change_Type,Table_ID,Change_Table,ChangeSubTable,Object_Title,User_ID,Change_Date,IP)" & _
       " VALUES (" & parDep & "," & pY & "," & pM & "," & pD & ",'" & Change_Type & "'," & Table_ID & ",'" & Change_Table & "','" & ChangeSubTable & "','" & Object_Title & "'," & UserID & ",GetDate(),'" & Request.ServerVariables("REMOTE_ADDR") & "')", con)
        cmdInsert.CommandType = CommandType.Text
        '   cmdInsert.Parameters.Add("@region_name", SqlDbType.VarChar, 150).Value = CStr(regionName)
        '    cmdInsert.Parameters.Add("@max_order", SqlDbType.Int).Value = MaxOrder
        con.Open()
        cmdInsert.ExecuteNonQuery()
        con.Close()


        ' Response.Write(selVName & "_" & DayTypeColor & "_" & DayFontColor)
    End Sub

End Class
