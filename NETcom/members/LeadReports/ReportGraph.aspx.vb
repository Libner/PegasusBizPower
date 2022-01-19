Public Class ReportGraph
    Inherits System.Web.UI.Page
    Protected start_date, end_date, title, label1, label2, label3, makor As String
    Protected pAppDate, pAPPID, pQUESTIONS_ID, pCONTACT_ID As String

    Protected sumContactus As Integer = 0
    Protected sum16504, sum16504Status5 As Integer

    Dim func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim con1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim con2 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))


    Dim rs_Makor, dr, drForm As System.Data.SqlClient.SqlDataReader

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
        If Trim(Request("dateStart")) <> "" Then
            start_date = DateValue(Trim(Request("dateStart")))
        Else
            ' start_date = "1/01/2016"
            ' start_date = "1/12/2017"

        End If
        If Trim(Request("dateEnd")) <> "" Then
            end_date = DateValue(Trim(Request("dateEnd")))
        Else

            '  end_date = "31/12/2017"
        End If
        title = "דוח מקורות הגעה " & start_date & " - " & end_date
        con.Open()
        Dim sqlstrMakor As New SqlClient.SqlCommand("SET DATEFORMAT DMY;SELECT distinct FIELD_VALUE from  FORM_VALUE " & _
    "  Left Join APPEALS on APPEALS.APPEAL_ID=FORM_VALUE.APPEAL_ID    where " & _
    " (FORM_VALUE.FIELD_ID = 40812 OR  FORM_VALUE.FIELD_ID = 40776 OR  FORM_VALUE.FIELD_ID = 40820) AND (LTRIM(RTRIM(FORM_VALUE.FIELD_VALUE)) <> '') " & _
    " and (APPEAL_DATE BETWEEN COALESCE('" & start_date & "',APPEAL_DATE) AND  COALESCE('" & end_date & "',APPEAL_DATE)) ", con)
        rs_Makor = sqlstrMakor.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_Makor.Read()

            If Not IsDBNull(rs_Makor("FIELD_VALUE")) Then
                makor = Trim(rs_Makor("FIELD_VALUE"))
                Dim cmdTmpCount = New SqlClient.SqlCommand("SET DATEFORMAT DMY;SELECT Count(APPEALS.Appeal_Id) as sumContactus FROM dbo.APPEALS Left Join FORM_VALUE " & _
              " on APPEALS.APPEAL_ID=FORM_VALUE.APPEAL_ID " & _
              " where (FORM_VALUE.FIELD_ID = 40812 OR  FORM_VALUE.FIELD_ID = 40776 OR  FORM_VALUE.FIELD_ID = 40820) AND (LTRIM(RTRIM(FORM_VALUE.FIELD_VALUE)) like '" & makor & "') " & _
              " and (APPEAL_DATE BETWEEN COALESCE('" & start_date & "',APPEAL_DATE) AND  COALESCE('" & end_date & "',APPEAL_DATE))", con1)
                con1.Open()
                dr = cmdTmpCount.ExecuteReader(CommandBehavior.SingleRow)
                If dr.Read() Then
                    sumContactus = dr("sumContactus")
                Else
                    sumContactus = 0
                End If
                con1.Close()
                '   Dim sqlForm
                Dim s As String
                ' s = "Exec dbo.get_LeadReportMakor  @date_start='" & start_date & "', @date_end='" & end_date & "',makor='" & makor & " '"
                ' Response.Write(s)
                ' Response.End()
                Dim cmdTmpCount1 = New SqlClient.SqlCommand("Exec dbo.get_LeadReportMakor  @date_start='" & start_date & "', @date_end='" & end_date & "',@makor='" & Trim(makor) & "'", con1)

                con1.Open()
                dr = cmdTmpCount1.ExecuteReader(CommandBehavior.CloseConnection)
                sum16504 = 0
                sum16504Status5 = 0
                While dr.Read()
                    ' Response.Write("dr.Read1")
                    pAppDate = dr("APPEAL_DATE")
                    pCONTACT_ID = dr("CONTACT_ID")
                    pAPPID = dr("APPEAL_Id")
                    pQUESTIONS_ID = dr("QUESTIONS_ID")
                    
                    Dim sqlForm = New SqlClient.SqlCommand("Exec dbo.get_LeadReport_16504  @CONTACT_ID='" & pCONTACT_ID & "', @AppDate='" & pAppDate & "'", con2)
                    s = "Exec dbo.get_LeadReport_16504  @CONTACT_ID='" & pCONTACT_ID & "', @AppDate='" & pAppDate & "'"
                    '   Response.Write(s & "<BR>")
                    '  Response.End()
                    con2.Open()
                    drForm = sqlForm.ExecuteReader(CommandBehavior.CloseConnection)
                    If drForm.Read() Then
                        sum16504 = sum16504 + 1
                        If drForm("APPEAL_STATUS") = 5 Then
                            sum16504Status5 = sum16504Status5 + 1
                        End If

                    End If

                    con2.Close()
                End While
                con1.Close()





                If label1 <> "" Then
                    label1 = label1 & "," & "{ label: """ & makor & """, y: " & sumContactus & " }"
                    label2 = label2 & "," & "{ label: """ & makor & """, y:" & sum16504 & " }"
                    label3 = label3 & "," & "{ label: """ & makor & """, y:" & sum16504Status5 & "}"

                Else
                    label1 = "{ label: """ & makor & """, y:" & sumContactus & " }"
                    label2 = "{ label: """ & makor & """, y: " & sum16504 & " }"
                    label3 = "{ label: """ & makor & """, y: " & sum16504Status5 & "}"
                End If


            End If

        End While
        con.Close()
        '  Response.Write("111")
        ' Response.End()
        ' Response.Write(label1)


        '  label1 = "{ label: ""banana"", y: 10 },{ label: ""orange"", y: 15 },{ label: ""apple"", y: 20 }, { label: ""mango"", y: 25 },{ label: ""grape"", y: 30 }"
        '  label2 = "{label: ""banana"", y: 23 },{ label: ""orange"", y: 33 },{ label: ""apple"", y: 48 },{ label: ""mango"", y: 37 },{ label: ""grape"", y: 20 }"
        '    label3 = "{ label:""banana"", y: 10 }, { label: ""orange"", y: 27 },{ label: ""apple"", y: 42 },{ label: ""mango"", y: 33 },{ label: ""grape"", y: 26 }"
    End Sub

End Class
