Imports System.Data.SqlClient
Public Class SalesEfficiency
    Inherits System.Web.UI.Page
    Protected dateStart, dateEnd As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected func As New bizpower.cfunc
    Public dtDep, dtUser, dtCountry As New DataTable
    Dim primKeydtDep(0), primKeydtUser(0), primKeydtCountry(0) As Data.DataColumn
    Protected WithEvents rptData As Repeater
    Protected WithEvents seldep, selUser, selCountry As System.Web.UI.HtmlControls.HtmlSelect
    Public dep, usr, depName, country, RadioType As String
    Protected WithEvents Button1, Button2 As System.Web.UI.WebControls.Button
    Public SumVar16504, SumVar16470_40811_out, SumVar16470_40105, SumVar16470_40811_in, SumVar16735_40660, SumVar16735_40660Bitul, sumColumn5, sum16735To16504 As Integer
    Protected WorkDays, sumNumberDaysWork, SumVar1 As Decimal

    Protected SumVarnumberOf16735_16470totalperUser As Integer
    Protected pdateStart, pdateEnd As String


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
        SumVar1 = 0
        Button1.Text = "בצע"
        Button2.Text = "ייצוא לאקסל"
        'Button2.Visible = False
        If Request("dateStart") = "" Then
            dateStart = "1/" & Month(Now()) & "/" & Year(Now())
        Else
            dateStart = Request("dateStart")
        End If
        If Request("dateEnd") = "" Then
            dateEnd = Now().ToString("dd/MM/yyyy")
        Else
            dateEnd = Request("dateEnd")
        End If

        If Request("seldep") = "" And Request("selUser") = "" Then
            dep = "1"
        Else
            dep = Request("seldep")

        End If
        If Request("selUser") <> "" Then
            usr = Request("selUser")

        End If
        If Request("RadioType") <> "" Then
            RadioType = Request("RadioType")
        Else
            RadioType = 1
        End If
        ''   Response.Write("usr=" & Year(dateStart) & "/" & Month(dateStart) & "/" & Day(dateStart))
        depName = func.GetSelectDepName(dep)

        'Dim ss = "SET DATEFORMAT YMD;SELECT count(id) as WorkDays " & _
        '                      "  FROM DimDate  WHERE DateDiff(d,DateKey, convert(smalldatetime,''' + @start_date + ''',101)) <= 0  " & _
        '                      "  AND DateDiff(d,DateKey, convert(smalldatetime,''' + @end_date + ''',101)) >= 0  and  IsHoliday=0 and DayTypeId<>'0'"
        'Response.Write(ss)
        'Response.End()
        pdateStart = Year(dateStart) & "/" & Month(dateStart) & "/" & Day(dateStart)
        pdateEnd = Year(dateEnd) & "/" & Month(dateEnd) & "/" & Day(dateEnd)

        Dim cmdSelect = New SqlCommand("SET DATEFORMAT YMD;SELECT sum(cast(DayTypeId as  DECIMAL(9,1))) as WorkDays " & _
                "  FROM DimDate  WHERE DateDiff(d,DateKey, convert(smalldatetime,'" & pdateStart & "',101)) <= 0  " & _
                "  AND DateDiff(d,DateKey, convert(smalldatetime,'" & pdateEnd & "',101)) >= 0  and DayTypeId<>'0'", con)
        ''and  IsHoliday=0 
        'cmdSelect.Parameters.Add("@start_date", SqlDbType.VarChar, 10).Value = Year(dateStart) & "/" & Month(dateStart) & "/" & Day(dateStart)
        'cmdSelect.Parameters.Add("@end_date", SqlDbType.VarChar, 10).Value = Year(dateEnd) & "/" & Month(dateEnd) & "/" & Day(dateEnd)

        '''   Response.Write(cmdSelect.CommandText)

        'Response.End()
        con.Open()
        WorkDays = cmdSelect.ExecuteScalar
        'Response.Write("WorkDays=" & WorkDays)
        con.Close()

''''
'''SELECT U.User_Id,FIRSTNAME, LASTNAME, (SELECT COUNT(L.Log_Id) FROM Users_LogIn L where User_id=U.User_Id group by User_id,LogInDate)
'''  AS LogEnter  
'''		 FROM dbo.Users U 
'''	--left join jobs J ON u.job_Id=J.job_Id
	  
'''sET DATEFORMAT YMD;SELECT COUNT(L.Log_Id) as l FROM Users_LogIn L 
'''where 
'''DateDiff(d,LogInDate, convert(smalldatetime,'2018/10/01',101)) <= 0 
'''              AND DateDiff(d,LogInDate, convert(smalldatetime,'2018/10/30',101)) >= 0
'''group by User_id

'''select LogInDate,DateDiff(d,LogInDate, convert(smalldatetime,'2018/10/01',101)), DateDiff(d,LogInDate, convert(smalldatetime,'2018/10/30',101))  from Users_LogIn
''''


        If Not IsPostBack Then
            GetData()


            GetDep()
            GetUsers()
            GetCountry()

        End If


    End Sub
    Sub GetData()
        Dim s As String = ""
        Dim dr As SqlDataReader

        If RadioType = 2 Then
            ' '            s = "SET DATEFORMAT DMY;SELECT FIRSTNAME + Char(32) + LASTNAME AS User_Name ,departmentName from Users " & _
            ' '" left JOIN Departments ON Users.Department_Id = Departments.departmentId " & _
            ' '" where  ACTIVE = 1 and  User_Id in (" & usr & ") order by departmentName,FIRSTNAME,LASTNAME  "

            ' s = "  SET DATEFORMAT DMY;select FIRSTNAME + Char(32) + LASTNAME AS User_Name ,departmentName, Count(*) AS Total_Amount" & _
            '" from Users   left join Users_LogIn on Users_LogIn.USER_ID=Users.User_Id " & _
            '" inner join Departments  ON Users.Department_Id = Departments.departmentId " & _
            '" where  ACTIVE = 1 and  Users.User_Id in (" & usr & ") group by  FIRSTNAME, LASTNAME ,departmentName" & _
            ''" order by departmentName,FIRSTNAME, LASTNAME"
            Dim cmdSelectProc As New SqlClient.SqlCommand("get_SalesEfficiencyByUser_User", con)
            cmdSelectProc.CommandType = CommandType.StoredProcedure
            cmdSelectProc.Parameters.Add("@user", SqlDbType.VarChar, 200).Value = usr
            cmdSelectProc.Parameters.Add("@dateStart", SqlDbType.VarChar, 10).Value = pdateStart
            cmdSelectProc.Parameters.Add("@dateEnd", SqlDbType.VarChar, 10).Value = pdateEnd
            con.Open()
            dr = cmdSelectProc.ExecuteReader(CommandBehavior.CloseConnection)
            '   Response.Write(usr & "--" & pdateStart & "--" & pdateEnd)
        Else

            '            s = "SET DATEFORMAT DMY;SELECT FIRSTNAME + Char(32) + LASTNAME AS User_Name ,departmentName from Users " & _
            '" left JOIN Departments ON Users.Department_Id = Departments.departmentId " & _
            '" where  ACTIVE = 1 and  Department_Id in (" & dep & ") order by departmentName,FIRSTNAME,LASTNAME  "
            '''s = "  SET DATEFORMAT DMY;select FIRSTNAME + Char(32) + LASTNAME AS User_Name ,departmentName, Count(*) AS Total_Amount" & _
            '''   " from Users   left join Users_LogIn on Users_LogIn.USER_ID=Users.User_Id " & _
            '''   " inner join Departments  ON Users.Department_Id = Departments.departmentId " & _
            '''   " where  ACTIVE = 1 and  Department_Id in (" & dep & ")  group by  FIRSTNAME, LASTNAME ,departmentName" & _
            '''   " order by departmentName,FIRSTNAME, LASTNAME"
            '--and  Department_Id in (" & dep & ")

            's = "SET DATEFORMAT DMY; SELECT Users.User_Id,FIRSTNAME + Char(32) + LASTNAME AS User_Name ,departmentName," & _
            '  " count(Users_LogIn.Log_id) AS numberOfDays FROM Users LEFT JOIN  Users_LogIn ON  Users_LogIn.USER_ID=Users.User_Id" & _
            '  " left join Departments  ON Users.Department_Id = Departments.departmentId  where  ACTIVE = 1 and   Department_Id in (" & dep & ")" & _
            '    " group by Users.User_Id,FIRSTNAME + Char(32) + LASTNAME ,departmentName order by departmentName,FIRSTNAME + Char(32) + LASTNAME "
            '' s = "SET DATEFORMAT DMY; SELECT Users.User_Id,FIRSTNAME + Char(32) + LASTNAME AS User_Name ,departmentName," & _
            '    " (select count(Users_LogIn.Log_id)  from Users_LogIn where  Users_LogIn.USER_ID=Users.User_Id " & _
            '    "  and  DateDiff(d,LogInDate, convert(smalldatetime,'" & pdateStart & "',101)) <= 0 " & _
            '    "  and DateDiff(d,LogInDate, convert(smalldatetime,'" & pdateEnd & "',101)) >= 0 " & _
            '    "  )		AS numberOfDays   FROM Users LEFT JOIN  Users_LogIn ON  Users_LogIn.USER_ID=Users.User_Id " & _
            '    " left join Departments  ON Users.Department_Id = Departments.departmentId  where  ACTIVE = 1 and   Department_Id in  (" & dep & ")" & _
            '    " group by Users.User_Id,FIRSTNAME + Char(32) + LASTNAME ,departmentName order by departmentName,FIRSTNAME + Char(32) + LASTNAME "
            'Response.Write("s=" & s & "<BR>")

            Dim cmdSelectProc As New SqlClient.SqlCommand("get_SalesEfficiencyByUser", con)
            cmdSelectProc.CommandType = CommandType.StoredProcedure
            cmdSelectProc.Parameters.Add("@depId", SqlDbType.VarChar, 20).Value = dep
            cmdSelectProc.Parameters.Add("@dateStart", SqlDbType.VarChar, 10).Value = pdateStart
            cmdSelectProc.Parameters.Add("@dateEnd", SqlDbType.VarChar, 10).Value = pdateEnd
            con.Open()
            dr = cmdSelectProc.ExecuteReader(CommandBehavior.CloseConnection)

            'Response.Write("get_SalesEfficiencyByUser=-" & dep & "--" & pdateStart & "--" & pdateEnd)
            ' Response.End()
        End If
        '''  Response.Write("s=" & s & "<BR>")
        '   Response.End()
        ' Dim cmdSelect = New SqlCommand(s, con)
        '"  WHERE CalendarYear = @currYear and CalendarMonth=@month  and DD.departmentId=@dep ORDER BY CalendarDay, CalendarMonth ", con)
        '
        'cmdSelect.Parameters.Add("@currYear", SqlDbType.Int).Value = CInt(currYear)
        'cmdSelect.Parameters.Add("@month", SqlDbType.Int).Value = CInt(currMonth)
        'cmdSelect.Parameters.Add("@dep", SqlDbType.VarChar, 50).Value = dep

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptData.DataSource = dr
            rptData.DataBind()
            rptData.Visible = True
        Else
            rptData.Visible = False

        End If


        con.Close()

    End Sub
    Sub GetCountry()

        Dim cmdSelect = New SqlCommand("select Country_Name,Country_Id From dbo.Countries order by Country_Name", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtCountry)
        con.Close()
        primKeydtCountry(0) = dtCountry.Columns("Country_Id")
        dtCountry.PrimaryKey = primKeydtCountry
        selCountry.Items.Clear()
        '  Dim list1 As New ListItem("הכל", "0")
        '  seldep.Items.Add(list1)
        For i As Integer = 0 To dtCountry.Rows.Count - 1
            Dim list As New ListItem(dtCountry.Rows(i)("Country_Name"), dtCountry.Rows(i)("Country_Id"))
            If Request("selCountry") > 0 And Request("selCountry") = dtCountry.Rows(i)("Country_Id") Then
                list.Selected = True
            ElseIf i = 0 Then
                list.Selected = True
                country = dtCountry.Rows(i)("Country_Id")
            End If
            selCountry.Items.Add(list)
        Next
        'Response.Write("dep=" & dep)
        ' Response.End()

    End Sub
    Sub GetUsers()
        Dim cmdSelect = New SqlCommand("Select USER_ID, FIRSTNAME + char(32) + LASTNAME as UserName  from Users where  ACTIVE = 1 order by FIRSTNAME,LASTNAME", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtUser)
        con.Close()
        primKeydtUser(0) = dtUser.Columns("USER_ID")
        dtUser.PrimaryKey = primKeydtUser
        selUser.Items.Clear()
        '  Dim list1 As New ListItem("הכל", "0")
        '  seldep.Items.Add(list1)
        For i As Integer = 0 To dtUser.Rows.Count - 1
            Dim list As New ListItem(dtUser.Rows(i)("UserName"), dtUser.Rows(i)("USER_ID"))
            If Request("selUser") > 0 And Request("selUser") = dtUser.Rows(i)("USER_ID") Then
                list.Selected = True
            ElseIf i = 0 Then
                list.Selected = True
                usr = dtUser.Rows(i)("USER_ID")
            End If
            selUser.Items.Add(list)
        Next
        'Response.Write("dep=" & dep)
        ' Response.End()


    End Sub
    Sub GetDep()
        Dim cmdSelect = New SqlCommand("Select departmentId, departmentName,PriorityLevel from Departments order by departmentName", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtDep)
        con.Close()
        primKeydtDep(0) = dtDep.Columns("departmentId")
        dtDep.PrimaryKey = primKeydtDep
        seldep.Items.Clear()
        '  Dim list1 As New ListItem("הכל", "0")
        '  seldep.Items.Add(list1)
        For i As Integer = 0 To dtDep.Rows.Count - 1
            Dim list As New ListItem(dtDep.Rows(i)("departmentName"), dtDep.Rows(i)("departmentId"))
            'If (Request("seldep") > 0 And Request("seldep") = dtDep.Rows(i)("departmentId")) Then
            '   Response.Write("dep=" & dep & "<BR>")
            '  Response.Write("dep=" & dtDep.Rows(i)("departmentId"))
            If dep = dtDep.Rows(i)("departmentId") Then
                list.Selected = True
                'ElseIf i = 0 Then
                '    list.Selected = True
                '    dep = dtDep.Rows(i)("departmentId")
            End If
            seldep.Items.Add(list)
        Next
        'Response.Write("dep=" & dep)
        ' Response.End()


    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
       
        GetData()
    End Sub
    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
     
        GetData()
    End Sub

    Private Sub rptData_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptData.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim UserId As String
            Dim pDays As Label
            UserId = e.Item.DataItem("user_Id")
            Dim var1 As Decimal
            Dim numberOfDays As Decimal

            Dim var16504 As Integer
            Dim var16470_40811_out As Integer
            Dim var16470_40811_in As Integer
            Dim var16735_40660 As Integer
            Dim var16735_40660Bitul As Integer
            Dim var16470_40105 As Integer   '-  כמות טפסי "תיעודי שיחה" יוצאת כמות שמכילה טקסטים של "אין מענה" 
            Dim var16735To16504 As Integer ''--numberOf16735To16504 	 מתוכם בכמה בוצעה מכירה column 8 
            Dim varnumberOf16735_16470totalperUser As Integer  'column 12 כמות מכירות ב"תהליך מלא
            '    Response.Write(WorkDays & ":" & e.Item.DataItem("numberOfDays"))
            '    Response.End()
            If IsDBNull(e.Item.DataItem("numberOfDays")) Then
                numberOfDays = 0

            Else
                numberOfDays = CDbl(e.Item.DataItem("numberOfDays"))
            End If
            var1 = CDbl(WorkDays) - numberOfDays
            pDays = e.Item.FindControl("pDays")
            If var1 > 0 Then
                pDays.Text = "<a href='#' style='font-color:#000000' onClick=openDays('" & UserId & "');>" & var1 & "</a>"
            End If

            sumNumberDaysWork = sumNumberDaysWork + CDbl(e.Item.DataItem("numberOfDays")) '' sum day work column1 
            SumVar1 = SumVar1 + var1
            var16504 = e.Item.DataItem("numberOf16504")
            SumVar16504 = SumVar16504 + var16504

            var16470_40811_out = e.Item.DataItem("numberOf16470_40811_out")
            SumVar16470_40811_out = SumVar16470_40811_out + var16470_40811_out

            var16470_40811_in = e.Item.DataItem("numberOf16470_40811_in")
            SumVar16470_40811_in = SumVar16470_40811_in + var16470_40811_in

            If IsNumeric(e.Item.DataItem("numberOf16735_40660")) Then
                var16735_40660 = e.Item.DataItem("numberOf16735_40660")
                SumVar16735_40660 = SumVar16735_40660 + var16735_40660
            End If

            If IsNumeric(e.Item.DataItem("numberOf16735_40660Bitul")) Then
                var16735_40660Bitul = e.Item.DataItem("numberOf16735_40660Bitul")
                SumVar16735_40660Bitul = SumVar16735_40660Bitul + var16735_40660Bitul
            End If
            If IsNumeric(e.Item.DataItem("numberOf16470_40105")) Then
                var16470_40105 = e.Item.DataItem("numberOf16470_40105")
            Else
                var16470_40105 = 0
            End If

            '==============Added by Mila 25/02/20020: 
            '=============="תיעודי השיחה היוצאים" - "אין מענה" 
            SumVar16470_40105 = SumVar16470_40105 + var16470_40105
            '===============================================================================

            ''--  column=5 כמות כללית של סך השיחות
            sumColumn5 = sumColumn5 + var16504 + var16470_40811_in + var16470_40105
            ''--  

            If IsNumeric(e.Item.DataItem("numberOf16735To16504")) Then
                var16735To16504 = e.Item.DataItem("numberOf16735To16504")
                sum16735To16504 = sum16735To16504 + var16735To16504
            End If
            If IsNumeric(e.Item.DataItem("numberOf16735_16470totalperUser")) Then
                varnumberOf16735_16470totalperUser = e.Item.DataItem("numberOf16735_16470totalperUser")
                SumVarnumberOf16735_16470totalperUser = SumVarnumberOf16735_16470totalperUser + varnumberOf16735_16470totalperUser
            End If






            '''   Response.Write("SumVar1=" & SumVar1 & "<BR>")
        End If

    End Sub
End Class
