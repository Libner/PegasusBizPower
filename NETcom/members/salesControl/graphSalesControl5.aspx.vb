Public Class graphSalesControl5
    Inherits System.Web.UI.Page
    Dim func As New include.funcs
    Protected currentYear, res1, res1All, res2, res2All, title, label2, label1, pMonthName, proc2, proc1 As String
    Protected sumLabel1, sumLabel2, proc, sumProc, rezProc As Decimal
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim rs_Guide, rs_r As System.Data.SqlClient.SqlDataReader

    Dim sqlstrGuide As New SqlClient.SqlCommand

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        currentYear = Request("currentYear")
        '--טופס מתעניין בטיול--16504

        For i As Integer = 1 To 12
            ' MonthName = Month(i)
            pMonthName = MonthName(i)
            res1 = 0
            res2 = 0

            '    ' title = "דוח מכירות יעד מול ביצוע שנת " & currentYear
            con.Open()
            Dim sql As New SqlClient.SqlCommand("SET DATEFORMAT DMY;SELECT sum(IsNull(P16504,0) )as P16504, CalendarYear FROM DimDate  WHERE (CalendarYear =" & currentYear & _
             " or CalendarYear =" & currentYear - 1 & " ) " & _
            " and CalendarMonth=" & i & " group by CalendarYear order by CalendarYear", con)
            rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
            While rs_r.Read()
                If rs_r("CalendarYear") = currentYear Then
                    res1 = rs_r("P16504") 'sum P16504 currentYear 
                Else
                    res2 = rs_r("P16504") 'sum P16504 currentYear-1
                End If
            End While
            con.Close()
            con.Open()
            '''prodId = 16735	טופס רישום חתום
            Dim sqlCurr As New SqlClient.SqlCommand("SET DATEFORMAT DMY;select count(APPEAL_ID) as sumA  from Appeals  where QUESTIONS_ID=16735 and RelationId in ( select APPEAL_ID from Appeals where QUESTIONS_ID=16504 " & _
                    "  and (year(APPEAL_DATE)=" & currentYear & ")  and Month(APPEAL_DATE)=" & i & " and appeal_status=5) ", con)
            If Not IsDBNull(sqlCurr.ExecuteScalar()) Then
                res1All = sqlCurr.ExecuteScalar()
            Else
                res1All = 0
            End If
            con.Close()
            con.Open()


            Dim sqlPrevY As New SqlClient.SqlCommand("SET DATEFORMAT DMY;select count(APPEAL_ID) as sumA  from Appeals  where QUESTIONS_ID=16735 and RelationId in ( select APPEAL_ID from Appeals where QUESTIONS_ID=16504 " & _
                           "  and (year(APPEAL_DATE)=" & currentYear - 1 & " )  and Month(APPEAL_DATE)=" & i & "	and appeal_status=5) ", con)
            If Not IsDBNull(sqlPrevY.ExecuteScalar()) Then
                res2All = sqlPrevY.ExecuteScalar()
            Else
                res2All = 0
            End If
            con.Close()





            'select count(APPEAL_ID) as sumA  from Appeals  
            'where QUESTIONS_ID=16735 and RelationId in 
            '( select APPEAL_ID from Appeals where QUESTIONS_ID=16504
            '  and (
            '		 year(APPEAL_DATE)=2016 or year(APPEAL_DATE)=2017
            '	  )
            '	   and Month(APPEAL_DATE)=11 	   and appeal_status=5) 



            'select count(APPEAL_ID) as sumA from Appeals  
            'where QUESTIONS_ID=16735 and RelationId in 
            '(select APPEAL_ID from Appeals where QUESTIONS_ID=16504  and year(APPEAL_DATE)=2016 and Month(APPEAL_DATE)=1) and appeal_status=5




            'select count(APPEAL_ID) as sumA ,year(APPEAL_DATE) from Appeals  
            'where QUESTIONS_ID=16735 and RelationId in 
            '(select APPEAL_ID from Appeals where QUESTIONS_ID=16504  and (year(APPEAL_DATE)=2016 or year(APPEAL_DATE)=2017) and Month(APPEAL_DATE)=1) and appeal_status=5
            'group by year(APPEAL_DATE)



            'select count(APPEAL_ID) as sumA ,year(APPEAL_DATE) from Appeals  
            'where QUESTIONS_ID=16735 and RelationId in 
            '( select APPEAL_ID from Appeals where QUESTIONS_ID=16504
            '  and (
            '		 year(APPEAL_DATE)=2016 or year(APPEAL_DATE)=2017
            '	  )
            '	   and Month(APPEAL_DATE)=1) 
            '	   and appeal_status=5

            '''group by year(APPEAL_DATE)

            '    rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
            '    While rs_r.Read()
            '        If rs_r("CalendarYear") = currentYear Then
            '            res1 = rs_r("P16504")
            '        Else
            '            res2 = rs_r("P16504")
            '        End If

            '        ' res2 = rs_r("Sales_Point")


            '    End While
            '    If IsNumeric(res1) Then
            '        sumLabel1 = sumLabel1 + res1
            '    End If
            '    If IsNumeric(res2) Then
            '        sumLabel2 = sumLabel2 + res2
            '    End If
            '    If res2 > 0 Then
            '        proc = (res1 / res2) * 100
            '    Else
            '        proc = 0
            '    End If
            '    'label2 = label2 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """, y: " & res2 & ",indexLabel: """ & res2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            '    'label1 = label1 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """,, y:" & res1 & ",indexLabel: """ & res1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"

            '    label2 = label2 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """, y: " & res2 & ",indexLabel: """ & res2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            '    label1 = label1 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """, y:" & res1 & ",indexLabel: """ & res1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            '  res1 = 0
            '  res2 = 0
            If res2All > 0 Then
                '   proc2 = String.Format("{0:0.0}", (res2 / res2All) * 100)
                proc2 = String.Format("{0:0.0}", (res2All / res2) * 100)
            Else

                proc2 = 0
            End If
            If res1All > 0 Then
                '    proc1 = String.Format("{0:0.0}", (res1 / res1All) * 100)
                proc1 = String.Format("{0:0.0}", (res1All / res1) * 100)
            Else
                proc1 = 0
            End If

            If proc2 > 0 Then
                rezProc = String.Format("{0:0.0}", (proc1 / proc2) * 100)
            End If

            ' label2 = label2 & "{label: """ & pMonthName & " " & """, y: " & proc2 & ",indexLabel: """ & proc2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            ' label1 = label1 & "{label: """ & pMonthName & " " & """, y:" & proc1 & ",indexLabel: """ & proc1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"

            label2 = label2 & "{label: """ & pMonthName & " " & rezProc & "%" & """, y: " & proc2 & ",indexLabel: """ & String.Format("{0:0.0}", proc2) & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            label1 = label1 & "{label: """ & pMonthName & " " & rezProc & "%" & """, y:" & proc1 & ",indexLabel: """ & String.Format("{0:0.0}", proc1) & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"

            '    con.Close()
        Next
        'If sumLabel2 > 0 Then
        '    sumProc = (sumLabel1 / sumLabel2) * 100
        'Else
        '    sumProc = 0
        'End If

        ''   label2 = label2 & "{label: """ & "שנתי" & " " & String.Format("{0:0.0}", sumProc) & "%" & """, y: " & sumLabel2 & ",indexLabel: """ & sumLabel2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        ''   label1 = label1 & "{label: """ & "שנתי" & " " & String.Format("{0:0.0}", sumProc) & "%" & """,, y:" & sumLabel1 & ",indexLabel: """ & sumLabel1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        'label2 = label2 & "{label: """ & "שנתי" & """, y: " & sumLabel2 & ",indexLabel: """ & sumLabel2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        'label1 = label1 & "{label: """ & "שנתי" & """, y:" & sumLabel1 & ",indexLabel: """ & sumLabel1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"




    End Sub
End Class
