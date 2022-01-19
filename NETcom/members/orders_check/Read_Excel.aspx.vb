Imports System.Data.SqlClient
Imports System.Data.Oledb
Imports System.Text.RegularExpressions

Public Class Read_Excel
    Inherits System.Web.UI.Page
    Protected object_name, sMessage As String
    Dim con As New SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected func As New bizpower.cfunc
    Protected validExcel As Boolean = True
    Protected IsTextFile As Boolean = True
    Private Const strRegex As String = "\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b"

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
        ''''Session.Timeout = 60

        If IsNothing(Request.Cookies("bizpegasus")) Then
            Response.Redirect("/")
        End If

        If Not IsNumeric(Request.Cookies("bizpegasus")("UserId")) Then
            Response.Redirect("/")
        End If

        Dim table As String = Trim(Request.QueryString("table"))
        Select Case LCase(table)
            Case "gilboa" : object_name = "דוח גלבוע"
        End Select
        Dim filename As String = Trim(table) & ".xls"

        GetExcelWorkSheet(Server.MapPath(Request.ApplicationPath & "/download/import/"), filename, 0, table)

    End Sub

    Private Function ColumnEqual(ByVal A As Object, ByVal B As Object) As Boolean
        ' Compares two values to see if they are equal. Also compares DBNULL.Value.
        'Note: If your DataTable contains object fields, then you must extend this
        ' function to handle them in a meaningful way if you intend to group on them.
        If (A Is DBNull.Value And B Is DBNull.Value) Then ' both are DBNull.Value
            Return True
        End If
        If (A Is DBNull.Value Or B Is DBNull.Value) Then '  only one is DBNull.Value
            Return False
        End If
        A = Trim(CStr(A)) : B = Trim(CStr(B))
        Return (A.Equals(B))    ' value type standard comparison
    End Function

    Private Sub GetExcelWorkSheet(ByVal pathName As String, ByVal fileName As String, ByVal workSheetNumber As Integer, _
ByVal table As String)
        Dim dt As New DataTable
        Dim ext As String = System.IO.Path.GetExtension(fileName)
        Dim ExcelConnection As OleDbConnection
        Dim ExcelCommand As New OleDbCommand

        ExcelConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
         "Data Source=" & pathName & "/" & fileName & ";" & "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;""")

        Try
            ExcelConnection.Open()
        Catch ex As Exception
            validExcel = False
        End Try

        If validExcel Then
            ExcelCommand.Connection = ExcelConnection

            Dim ExcelSheets As DataTable = ExcelConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, _
            New Object() {Nothing, Nothing, Nothing, "TABLE"})
            Dim SpreadSheetName As String = "[" & ExcelSheets.Rows(workSheetNumber)("TABLE_NAME").ToString() & "]"

            Dim XlsDA As OleDbDataAdapter
            ExcelCommand.CommandText = "SELECT TOP 100 PERCENT *  FROM " & SpreadSheetName
            XlsDA = New OleDbDataAdapter(ExcelCommand)
            XlsDA.Fill(dt)
            ExcelConnection.Close()
        Else
            ExcelConnection.Close()
            'Try read as text file
            'Get a StreamReader class that can be used to read the file
            Dim objStreamReader As System.IO.StreamReader
            Try
                objStreamReader = System.IO.File.OpenText(pathName & "/" & fileName)
            Catch ex As Exception
                IsTextFile = False
            End Try
            objStreamReader.Close()

            If IsTextFile Then

                Dim txtConnection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pathName & ";Extended Properties=""text;HDR=Yes;FMT=Delimited""")
                Dim txtCommand As New OleDbCommand("SELECT TOP 100 PERCENT * FROM " & fileName, txtConnection)
                Try
                    txtConnection.Open()
                    Dim da As New OleDb.OleDbDataAdapter(txtCommand)
                    da.Fill(dt)
                    txtConnection.Close()
                Catch ex As Exception
                    txtConnection.Close()
                    IsTextFile = False
                End Try
            End If
        End If

        If Not IsTextFile And Not validExcel Then 'Try read html xls file
            Dim htmlXslConnection As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & pathName & "/" & fileName & ";" & _
               "Extended Properties=""HTML Import;HDR=YES;IMEX=1""")

            Dim isHtmlXls As Boolean = False
            Try
                htmlXslConnection.Open()
                isHtmlXls = True
            Catch ex As Exception
                isHtmlXls = False
            End Try

            If isHtmlXls Then
                Dim ExcelSheets As DataTable = htmlXslConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, _
            New Object() {Nothing, Nothing, Nothing, "TABLE"})
                Dim SpreadSheetName As String = "[" & ExcelSheets.Rows(workSheetNumber)("TABLE_NAME").ToString() & "]"

                Dim txtCommand As New OleDbCommand("SELECT TOP 100 PERCENT * FROM " & SpreadSheetName, htmlXslConnection)
                Dim XlsDA As New OleDbDataAdapter(txtCommand)
                XlsDA.Fill(dt)
                htmlXslConnection.Close()
            End If

        End If

        'Dim dltCommand As New SqlClient.SqlCommand("DELETE FROM " & table, con)
        'con.Open()
        'dltCommand.ExecuteNonQuery()
        'con.Close()

        'For ii As Integer = 0 To dt.Columns.Count - 1
        '    Response.Write(dt.Columns(ii).ColumnName & ",")
        'Next
        'Response.End()

        'insert data into sql db
        For ss As Integer = 0 To dt.Rows.Count - 1
            Dim ServiceName = dt.Rows(ss)(0)
            If IsDBNull(ServiceName) Then
                ServiceName = ""
            Else
                ServiceName = Trim(ServiceName)
            End If
            Dim PFile = dt.Rows(ss)(1)
            If IsDBNull(PFile) Then
                PFile = ""
            Else
                PFile = Trim(PFile)
            End If
            Dim Pax = dt.Rows(ss)(2)
            If IsDBNull(Pax) Then
                Pax = ""
            Else
                Pax = Trim(Pax)
            End If

            If IsNumeric(Pax) And Len(ServiceName) > 0 Then
                Dim cmdInsert As New SqlClient.SqlCommand("gilboa_insert_docket", con)
                cmdInsert.CommandType = CommandType.StoredProcedure
                cmdInsert.Parameters.Add("@ServiceName", SqlDbType.Char, 10).Value = Trim(ServiceName)
                cmdInsert.Parameters.Add("@PFile", SqlDbType.Char, 10).Value = Trim(PFile)
                cmdInsert.Parameters.Add("@Pax", SqlDbType.Int).Value = Trim(Pax)
                'For Each prm As SqlClient.SqlParameter In cmdInsert.Parameters
                '    Response.Write(prm.ParameterName & " = " & prm.Value & "<br>")
                'Next
                'Response.End()
                con.Open()
                Dim tmp = cmdInsert.ExecuteScalar()
                con.Close()
                cmdInsert.Dispose()
                'Response.Write("PFile = " & PFile & ", ServiceName=" & tmp & "<br>")
            End If
        Next

        Dim cmdUpdate As New SqlClient.SqlCommand("gilboa_update_docket_status", con)
        cmdUpdate.CommandType = CommandType.StoredProcedure
        con.Open()
        cmdUpdate.ExecuteNonQuery()
        con.Close()

        sMessage = "<p>יבוא קובץ גילבוע הסתיים בהצלחה</p>"

    End Sub

End Class