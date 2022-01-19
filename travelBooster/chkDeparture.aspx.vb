Imports System.Web.Services.Protocols
Imports System.Web.Mail
Imports System.Xml
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.COMException
Imports System.Text.RegularExpressions

    Public Class chkDeparture
        Inherits System.Web.UI.Page
    Protected func As New include.funcs
        Dim travelBoosfunc As New travelBoost.funcs
    Protected con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New SqlClient.SqlCommand


           Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Session.LCID = 1037
        Dim isValid As String = "NO"
        Dim departureId, tourId, SeriesId As Integer
        Dim departureCode As String
        Dim TB_departureCode, TB_seriesCode, TB_departureDate, TB_departurePrice, TB_TotalSeats, TB_SeatsTaken, TB_AvailCount, TB_productID As String

        departureId = travelBoosfunc.fixNumeric(Request("depId"))
        departureId = 6070
        Response.Write("departureId=" & departureId)
        ' Response.End()
        If departureId > 0 Then
            Dim seriesCode, departureDate, departureDateEnd, price As String
            cmdSelect = New SqlClient.SqlCommand("SELECT Departure_Id,Departure_Code, Tours_Departures.Tour_Id, Price, Departure_Date, Departure_Date_End,Tours.SeriasId FROM Tours_Departures INNER JOIN Tours ON Tours_Departures.Tour_Id = Tours.Tour_Id " & _
               " where Departure_Id=@depId", con)
            cmdSelect.CommandType = CommandType.Text
            cmdSelect.Parameters.Add("@depId", SqlDbType.Int).Value = CInt(departureId)
            con.Open()
            Dim rd As SqlClient.SqlDataReader
            rd = cmdSelect.ExecuteReader
            If rd.Read Then
                tourId = travelBoosfunc.fixNumeric(rd("Tour_Id"))
                SeriesId = travelBoosfunc.fixNumeric(rd("seriasId"))
                departureCode = travelBoosfunc.dbNullFix(rd("Departure_Code"))
                departureDate = travelBoosfunc.dbNullFix(rd("Departure_Date"))
                departureDateEnd = travelBoosfunc.dbNullFix(rd("Departure_Date_End"))
                price = Regex.Match(travelBoosfunc.dbNullFix(rd("Price")), "^\d+").Value
                con.Close()
            End If
            'Response.Write(" seriasId=" & SeriesId)
            tourId = 100
            departureCode = "NYT0411"
            departureDate = "11/04/2019"
            departureDateEnd = "19/04/2019"
            price = "1390$"
            SeriesId = 7
            Response.Write("SeriesId=" & SeriesId)

            If SeriesId > 0 Then
                cmdSelect = New SqlClient.SqlCommand("SELECT Series_Name FROM Series  where Series_Id=@SeriesId", conB)
                cmdSelect.CommandType = CommandType.Text
                cmdSelect.Parameters.Add("@SeriesId", SqlDbType.Int).Value = CInt(SeriesId)
                conB.Open()
                Try
                    seriesCode = cmdSelect.ExecuteScalar
                Catch ex As Exception
                    'Response.Write(" seriesCode=" & Err.Description)
                End Try
                conB.Close()
            End If

            Response.Write(" seriesCode=" & seriesCode)

            Response.Write(" departureDate=" & IsDate(departureDate))
            Response.Write(" price=" & price)
            If seriesCode <> "" And departureDate <> "" And price <> "" Then
                If IsDate(departureDate) And travelBoosfunc.fixNumeric(price) > 0 Then
                    'check if exists at least one place: 1 room for 1 adult
                    Response.Write("yyyy=")
                    Dim xmldoc As New XmlDocument
                    xmldoc = travelBoosfunc.getDeparture(seriesCode, CDate(departureDate), CDate(departureDateEnd), 1, 1, 0)
                    Response.Write("111=" & xmldoc.Value)
                    Response.End()
                    'get price for general group - קהל רחב:
                    ' <priceLevelID>47</priceLevelID><priceLevelName>קהל רחב</priceLevelName> 

                    If Not xmldoc.SelectSingleNode("body/product[priceLevelID='47']") Is Nothing Then
                        'check if code and price in pegasus DB are corresponded to travelBoost WS 
                        Try
                            TB_seriesCode = xmldoc.SelectSingleNode("body/product[priceLevelID='47']/tourCode1").InnerText
                            TB_departureCode = xmldoc.SelectSingleNode("body/product[priceLevelID='47']/tourCode2").InnerText
                            TB_departureDate = xmldoc.SelectSingleNode("body/product[priceLevelID='47']/startDate").InnerText
                            TB_departurePrice = xmldoc.SelectSingleNode("body/product[priceLevelID='47']/priceTotal").InnerText
                            TB_productID = xmldoc.SelectSingleNode("body/product[priceLevelID='47']/productID").InnerText
                        Catch ex As Exception

                        End Try
                        'Try
                        '    TB_TotalSeats = xmldoc.SelectNodes("body/product/classes/class")(0).Attributes("total").Value
                        'Catch ex As Exception

                        'End Try
                        'Try
                        '    TB_SeatsTaken = xmldoc.SelectNodes("body/product/classes/class")(0).Attributes("taken").Value
                        'Catch ex As Exception

                        'End Try
                        Try
                            TB_AvailCount = xmldoc.SelectNodes("body/product[priceLevelID='47']/classes/class")(0).Attributes("availCount").Value 'number of Available places
                        Catch ex As Exception

                        End Try

                        Response.Write("<br>departureDate=" & departureDate)
                        Response.Write(" TB_departureDate=" & TB_departureDate)
                        Response.Write("<br>departureCode=" & travelBoosfunc.dbNullDate(CDate(departureDate), "MMdd"))
                        Response.Write(" TB_departureCode=" & TB_departureCode)
                        Response.Write("<br>price=" & price)
                        Response.End()
                        If TB_seriesCode = seriesCode And TB_departureCode = Right(departureCode, 4) And _
                                      travelBoosfunc.dbNullDate(CDate(TB_departureDate), "yyyy-MM-dd") = travelBoosfunc.dbNullDate(CDate(departureDate), "yyyy-MM-dd") And _
                                      travelBoosfunc.dbNullDate(CDate(departureDate), "MMdd") = TB_departureCode And _
                                     Math.Floor(travelBoosfunc.fixNumeric(price)) = Math.Floor(travelBoosfunc.fixNumeric(TB_departurePrice) / 2) Then
                            If travelBoosfunc.fixNumeric(TB_AvailCount) > 0 Then
                                isValid = "OK"
                            Else
                                isValid = "NOSEATS"
                            End If
                        End If
                    End If
                End If


                If isValid <> "OK" Then
                    'send mail to admin with permissions
                    Dim strMailTo As String
                    cmdSelect = New SqlClient.SqlCommand(" SELECT  Workers.Email FROM Workers_To_Actions INNER JOIN Workers ON Workers_To_Actions.Worker_ID = Workers.worker_Id " & _
" WHERE Workers_To_Actions.Action_Id = @ActionId", con)
                    cmdSelect.CommandType = CommandType.Text
                    cmdSelect.Parameters.Add("@ActionId", SqlDbType.Int).Value = 22 'actionId for receiving warrnings about TB
                    con.Open()
                    Dim rdrm As SqlClient.SqlDataReader
                    rdrm = cmdSelect.ExecuteReader
                    Do While rdrm.Read
                        strMailTo = strMailTo & travelBoosfunc.dbNullFix(rdrm("Email"))
                    Loop
                    con.Close()
                    If strMailTo <> "" Then


                        Dim objMail As New System.Web.Mail.MailMessage

                        objMail.From = "info@pegasusisrael.co.il"
                        ''''Response.Write("<br> mailto=" & strMailTo)
                        objMail.To = "faina@cyberserve.co.il"  '''strMailTo '"mila@cyberserve.co.il" 
                        '''' objMail.Bcc.Add("mila@cyberserve.co.il")
                        'objMail.Bcc = "mila@cyberserve.co.il"
                        objMail.Subject = " TBחוסר התאמת נתונים בין אתר החברה למערכת ה"

                        Dim PHstr As String
                        PHstr = ""

                        PHstr += "<html><head><meta charset=utf-8></head>"

                        PHstr += "<body>"


                        '*************************************************************************************

                        PHstr += "<h2 dir='rtl'> בעקבות כניסה לחוצץ יציאות קרובות של הסדרה " & seriesCode & "</h2>"
                        PHstr += "<h3 dir='rtl'>התגלה חוסר התאמה בטיול " & departureCode & "</h3>"
                        PHstr += "<h3 dir='rtl'>להלן הנתונים כפי מופיעים בכל אחת מן המערכות:</h3>"
                        PHstr += "<table cellpadding=5 cellspacing=0 border='1' align='right' dir='rtl'>"
                        PHstr += "<tr><th align='right' dir='rtl'>מערכת</th><th align='right' dir='rtl'>קוד טיול</th><th align='right' dir='rtl'>תאריך יציאה</th><th align='right' dir='rtl'>מחיר טיול</th>"
                        PHstr += "<tr>"
                        PHstr += "<td align='right' dir='rtl'>אתר</td>"
                        PHstr += "<td align='right' dir='rtl'>" & departureCode & "</td>"
                        PHstr += "<td align='right' dir='rtl'>" & travelBoosfunc.dbNullDate(CDate(departureDate), "yyyy-MM-dd") & "</td>"
                        PHstr += "<td align='right' dir='rtl'>" & price & "</td>"
                        PHstr += "</tr>"
                        PHstr += "<tr>"
                        PHstr += "<td align='right' dir='rtl'>TB</td>"
                        PHstr += "<td align='right' dir='rtl'>" & TB_seriesCode & TB_departureCode & "</td>"
                        PHstr += "<td align='right' dir='rtl'>"
                        If TB_departureDate <> "" Then
                            If IsDate(TB_departureDate) Then
                                PHstr += travelBoosfunc.dbNullDate(CDate(TB_departureDate), "yyyy-MM-dd")
                            End If
                        End If
                        PHstr += "</td>"
                        PHstr += "<td align='right' dir='rtl'>" & TB_departurePrice & "</td>"
                        PHstr += "</tr>"
                        PHstr += "<tr>"
                        PHstr += "<td align='right' dir='rtl'>קיים פער</td>"
                        PHstr += "<td align='right' dir='rtl'>" & IIf(TB_seriesCode = seriesCode And TB_departureCode = Right(departureCode, 4), "לא", "כן") & "</td>"
                        PHstr += "<td align='right' dir='rtl'>"
                        If TB_departureDate <> "" Then
                            If IsDate(TB_departureDate) Then
                                PHstr += IIf(travelBoosfunc.dbNullDate(CDate(TB_departureDate), "yyyy-MM-dd") = travelBoosfunc.dbNullDate(CDate(departureDate), "yyyy-MM-dd") And travelBoosfunc.dbNullDate(CDate(departureDate), "MMdd") = TB_departureCode, "לא", "כן")
                            Else
                                PHstr += "כן"
                            End If
                        Else
                            PHstr += "כן"
                        End If
                        PHstr += "</td>"
                        PHstr += "<td align='right' dir='rtl'>" & IIf(Math.Floor(travelBoosfunc.fixNumeric(price)) = Math.Floor(travelBoosfunc.fixNumeric(TB_departurePrice / 2)), "לא", "כן") & "</td>"
                        PHstr += "</tr>"
                        PHstr += "</table>"
                        PHstr += "</body>"
                        objMail.Body = PHstr
                        PHstr += "</html>"
                        objMail.BodyEncoding = System.Text.Encoding.UTF8

                        ' objMail.IsBodyHtml = True

                        objMail.BodyFormat = Mail.MailFormat.Html


                        Try
                            'Dim mSmtpClient As New System.Net.Mail.SmtpClient("john2")
                            '' Send the mail message
                            ' '''''''TEST - no sent
                            'mSmtpClient.Send(objMail)
                            ' If LCase(Request.ServerVariables("SERVER_NAME")).IndexOf("www.pegasusisrael.co.il") >= 0 Then

                            System.Web.Mail.SmtpMail.Send(objMail)
                            '   End If
                        Catch ex As Exception
                            Response.Write(Err.Description)
                        End Try
                        objMail = Nothing
                    End If

                End If
            End If
        End If


        Response.Write(isValid & "_" & TB_productID)

    End Sub

End Class