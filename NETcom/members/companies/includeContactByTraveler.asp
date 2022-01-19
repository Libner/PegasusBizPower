   <!--mashov travelers begin -->
    <%	sqlstrMashovTraveler = "Select  s.traveler_Id ,FirstName,LastName,IdNumber,FeedBack_Status, Tours_Departures.Departure_Id, Tours_Departures.Tour_Id, Tours_Departures.Departure_Code, Tours_Departures.Date_Begin,Tours_Departures.Departure_Date, (Guide_FName + ' ' + Guide_LName) as Guide_Name , " & _
    " Tour_Name=case when isnull(Tour_PageTitle,'')='' then Tour_Name else Tour_PageTitle end FROM bizpower_pegasus_test.[dbo].sms_FeedBackTo_Travelers  as s " & _
    " left join Tours_Departures on s.Departure_id=Tours_Departures.Departure_id " & _
  " INNER JOIN Tours ON Tours_Departures.Tour_Id = Tours.Tour_Id INNER JOIN Tours_Categories ON Tours.Category_Id = Tours_Categories.Category_Id " & _
" left JOIN Departments ON Tours.DeparureId = Departments.Dep_Id	" & _
" left join Guides on Tours_Departures.Guide_Id=Guides.Guide_Id " & _
 " left join bizpower_pegasus_test.dbo.Travelers on s.traveler_Id= bizpower_pegasus_test.dbo.Travelers.traveler_Id " & _
 " left join FeedBack_Form  on FeedBack_Form.traveler_Id=s.traveler_Id " & _
  " and FeedBack_Form.Departure_Id=s.Departure_Id	" & _
 " where s.Contact_Id=" &  ContactId & " and Company_id=" & Companyid &" and " & Application("pegasusDBName") & ".dbo.getFeedBackFaqMax("& ContactId &", Tours_Departures.departure_id)"& ">0"

'Response.Write sqlstrMashov
'response.End   
		set MashovListTraveler = conPegasus.Execute(sqlstrMashovTraveler)
	
		If not MashovListTraveler.eof Then
			recCount = 1
			else
			recCount = 0
		End If		
'response.Write "recCount="& recCount
    %>
    <%if recCount=1 then%>
    <tr>
        <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
            <table cellpadding="0" cellspacing="0" dir="<%=dir_var%>" width="100%" dir="<%=dir_var%>" id="Table9">
                <tr>
                    <td width="100%"><a name="table_SMS"></a>
                        <table cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>" id="Table10">
                            <tr>
                                <td class="title_form" width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;משוב מטייל<%'=trim(Request.Cookies("bizpegasus")("TasksMulti"))%>&nbsp;<font color="#E6E6E6">(<%=company_name & " - " & contacter%>)</font>&nbsp;</td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td width="100%" valign="top" dir="<%=dir_var%>">
                        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF" dir="<%=dir_var%>" id="Table11">
                            <tr>
                                <td align="center" class="title_sort" width="29" nowrap>&nbsp;</td>
                                <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" width="265" nowrap>&nbsp;</td>
                                <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width="75" nowrap class="title_sort">שם המדריך</td>
                                <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width="45" nowrap class="title_sort"><span style="color: rgb(0, 0, 0); font-family: Arial; font-size: 12px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: -webkit-right; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; display: inline !important; float: none; background-color: rgb(211, 211, 211);">טיול </span></td>
                                <td width="90" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort">קוד טיול</a></td>
                                <td width="90" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort">תאריך יציאה</td>
                                <td width="120" nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">סטטוס מילוי</td>
                            </tr>
                            <%   while not MashovListTraveler.EOF  	%>
                            <%select case MashovListTraveler("FeedBack_Status")
    case 1
    FeedBack_Status="חלקי"
    SendButton=1
    OpenForm=1
    case 2
    FeedBack_Status="מלא"
     SendButton=0
         OpenForm=1
   case  else 
   FeedBack_Status="בהמתנה"
    SendButton=1  'שלח משוב 
   OpenForm=0 
    end select
                            %>

                            <tr>
                                <td align="<%=align_var%>" dir="<%=dir_var%>" class="card<%=class_%>" valign="top"><%if SendButton=1 and Request.Cookies("bizpegasus")("UserSendSmsGeneralScreen")="1" then%>
                                    <a class="button_edit_1" href="javascript:void window.open('../feedback/SendMailFeedBack.aspx?DepartureId=<%=MashovListTraveler("Departure_Id")%>&companyId=<%=CompanyId%>&contactId=<%=ContactId%>','SendSMSFeedback','top=70, left=70, width=520, height=520, scrollbars=1');" onclick="return window.confirm('?האם ברצונך לשלוח מייל לאיש קשר');" style="width: 120px; background-color: #F7CD65; color: #000000">שלח משוב במייל</a>
                                    <a class="button_edit_1" href="javascript:void window.open('../feedback/SendSMSFeedBack.aspx?DepartureId=<%=MashovListTraveler("Departure_Id")%>&companyId=<%=CompanyId%>&contactId=<%=ContactId%>','SendSMSFeedback','top=70, left=70, width=520, height=520, scrollbars=1');" onclick="return window.confirm('?האם ברצונך לשלוח סמס לאיש קשר');" style="width: 120px; background-color: #F7CD65; color: #000000">SMS שלח משוב </a>
                                    <%end if%>
                                </td>
                                <td align="center" class="card<%=class_%>" width="29" nowrap><%if OpenForm=1 then%><a class="button_edit_1" style="width: 120px" href="javascript:window.open('../Feedback/Feedback.aspx?status=<%=MashovListTraveler("FeedBack_Status")%>&depId=<%=MashovListTraveler("Departure_Id")%>&companyId=<%=CompanyId%>&contactId=<%=ContactId%>','AddCat','top=10,left=10,width=900,height=450,scrollbars=1')" dir="<%=dir_obj_var%>">&nbsp; פרטי המשוב</a><%end if%></td>
                                <td align="<%=align_var%>" class="card<%=class_%>" valign="top"><%=MashovListTraveler("Guide_Name")%></td>
                                <td align="right" class="card<%=class_%>" valign="top"><%=MashovListTraveler("Tour_Name")%></td>
                                <td align="<%=align_var%>" class="card<%=class_%>" valign="top">&nbsp;<%=MashovListTraveler("Departure_Code")%>&nbsp;</td>
                                <td align="<%=align_var%>" class="card<%=class_%>" valign="top">&nbsp;<%=MashovListTraveler("Date_Begin")%>&nbsp;</td>
                                <td align="<%=align_var%>" class="card<%=class_%>" valign="top"><a class="link_categ">&nbsp;<%=FeedBack_Status%>&nbsp;</td>
                            </tr>



                            <%MashovListTraveler.MoveNext
    Wend%>
                        </table>
                    </td>
                </tr>

            </table>
        </td>
    </tr>
    <%end if%>
    <!--mashov-->