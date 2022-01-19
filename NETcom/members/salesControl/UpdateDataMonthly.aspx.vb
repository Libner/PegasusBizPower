Imports System.Data.SqlClient
Public Class UpdateDataMonthly
    Inherits System.Web.UI.Page

    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmd As New SqlClient.SqlCommand
    Public parName, Uid, parValue, dep As String


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
        ' parName = "Avrg_Call"
        Uid = Request.Form("Uid") '"12_2017" month -dep - year
        parValue = Request.Form("value")
        '  dep = Request.Form("dep")
        Dim pM, pY As String
        Dim strV As Array
        strV = Uid.Split("_")
        If strV.Length > 0 Then
            pM = strV(0)
            dep = strV(1)
            pY = strV(2)
        End If
        '  Dim pD As String = ""


        con.Open()
        'Dim sql = "Update DimDate set " & parName & "=" & parValue & " where DateKey=" & Uid
        'Response.Write(sql)

        cmd = New System.Data.SqlClient.SqlCommand("update_Departments_Data_Monthly", con)

        cmd.Parameters.Add("@parName", SqlDbType.VarChar).Value = parName
        cmd.Parameters.Add("@parValue", SqlDbType.VarChar).Value = parValue
        cmd.Parameters.Add("@pM", SqlDbType.Char, 2).Value = CInt(pM)
        cmd.Parameters.Add("@pY", SqlDbType.Char, 4).Value = CInt(pY)
        cmd.Parameters.Add("@dep", SqlDbType.Int).Value = CInt(dep)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        con.Close()
        Dim Change_Type = "עדכון"
        Dim Table_ID = 112 'מסך הגדרות יעדים in table bars
        Dim Change_Table = "מסך הגדרות יעדים"

        Dim ChangeSubTable As String
        Dim Object_Title As String
        If parName = "Sales_Point" Then
            Object_Title = "יעד מכירות שנה נוכחית"
            ChangeSubTable = "טבלת הגדרות יעדים שנתיים למכירות ולשירות וביצוע בפועל"
        ElseIf parName = "KishuritService_Point" Then
            Object_Title = "יעד קישוריות שירות של שנה נוכחית (אל מול שיחות שירות) (באחוזים) "
            ChangeSubTable = "טבלת הגדרות יעדים שנתיים למכירות ולשירות וביצוע בפועל"
        ElseIf parName = "ActiveTime_Point" Then
            Object_Title = "יעד זמן אקטיב לנציג שירות שנה נוכחית "
            ChangeSubTable = "טבלת הגדרות יעדים שנתיים למכירות ולשירות וביצוע בפועל"
        End If
        '   Dim s = "SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO ChangesSalesControl (" & _
        '         " Department_Id,Year,Month,Day,Change_Type,Table_ID,Change_Table,ChangeSubTable,Object_Title,User_ID,Change_Date,IP)" & _
        '       " VALUES (" & dep & "," & pY & "," & pM & ",'','" & Change_Type & "'," & Table_ID & ",'" & Change_Table & "','" & ChangeSubTable & "','" & Object_Title & "'," & UserID & ",GetDate(),'" & Request.ServerVariables("REMOTE_ADDR") & "')"
        'Response.Write(s)
        'Response.End()
        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO ChangesSalesControl (" & _
       " Department_Id,Year,Month,Day,Change_Type,Table_ID,Change_Table,ChangeSubTable,Object_Title,User_ID,Change_Date,IP)" & _
      " VALUES (" & dep & "," & pY & "," & pM & ",'','" & Change_Type & "'," & Table_ID & ",'" & Change_Table & "','" & ChangeSubTable & "','" & Object_Title & "'," & UserID & ",GetDate(),'" & Request.ServerVariables("REMOTE_ADDR") & "')", con)
        cmdInsert.CommandType = CommandType.Text
        '   cmdInsert.Parameters.Add("@region_name", SqlDbType.VarChar, 150).Value = CStr(regionName)
        '    cmdInsert.Parameters.Add("@max_order", SqlDbType.Int).Value = MaxOrder
        con.Open()
        cmdInsert.ExecuteNonQuery()
        con.Close()




    End Sub


End Class
