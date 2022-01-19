<%
dim catname(6)
catname(0)= "קשרי לקוחות"
catname(1)= "דיוור"
catname(2)= "משובים וסקרים"
catname(3)= "דוחות"
catname(4)= "כלכלה ופיננסים"
catname(5)= "אשפים"

dim subcatname(9,5)

subcatname(9,3)="ניהול משימות"
subcatname(9,2)="פעילות לקוח"
subcatname(9,1)="תיקי לקוחות"
subcatname(9,0)="מועדון לקוחות"

subcatname(8,3)="אתר אירוע"
subcatname(8,2)="אתר כנס"
subcatname(8,1)="אתר קמפיין"
subcatname(8,0)="אתר תדמיתי"

subcatname(7,3)="עמדות עובדים"
subcatname(7,2)="סקר שוק"
subcatname(7,1)="הערכת עובדים"
subcatname(7,0)="סקר לקוחות"

subcatname(6,4)="עדכונים וחדשות"
subcatname(6,3)="פרסום אירוע"
subcatname(6,2)="רישום לכנס"
subcatname(6,1)="מבצע עם רישום"
subcatname(6,0)="פרסום מבצע"

subcatname(5,3)="קשרי לקוחות"
subcatname(5,2)="אתרים יעודיים"
subcatname(5,1)="משובים וסקרים"
subcatname(5,0)="דיוור פרסומי"

If trim(chief)  = "1" Then
subcatname(4,5)="Bizpower desktop"
subcatname(4,4)="המחרה"
subcatname(4,3)="שעות העבודה שלי"
subcatname(4,2)="עלות העבודה"
subcatname(4,1)="פעילויות"
subcatname(4,0)="עובדים"
Else
subcatname(4,0)="שעות העבודה שלי"
End if

subcatname(3,0)="דוחות משובים וסקרים"
subcatname(3,1)="דוחות פיננסים"

subcatname(2,0)="קבוצות דיוור"
subcatname(2,1)="טפסי משוב"
subcatname(2,2)="משובים"
subcatname(2,3)="הפצות"

subcatname(1,0)="קבוצות דיוור"
subcatname(1,1)="דפים מעוצבים"
subcatname(1,2)="טפסי רישום"
subcatname(1,3)="הפצות"

subcatname(0,0)="לקוחות"
subcatname(0,1)="אנשי קשר"
subcatname(0,2)="משימות"

dim subcatlink(9,5)

subcatlink(6,4)="../../default.asp?wizard_id=5"
subcatlink(6,3)="../../default.asp?wizard_id=2"
subcatlink(6,2)="../../default.asp?wizard_id=4"
subcatlink(6,1)="../../default.asp?wizard_id=3"
subcatlink(6,0)="../../default.asp?wizard_id=1"

subcatlink(5,1)=null
subcatlink(5,0)=null

If trim(chief)  = "1" Then
subcatlink(4,5)="../projects/download.asp"
subcatlink(4,4)="../projects/prices.asp"
subcatlink(4,3)="../projects/hours_calendar.asp"
subcatlink(4,2)="../projects/projects_reports.asp"
subcatlink(4,1)="../projects/default.asp"
subcatlink(4,0)="../workers/default.asp"
Else
subcatlink(4,0)="../projects/hours.asp"
End If

subcatlink(3,0)="../reports/reports.asp"
subcatlink(3,1)="../reports/default.asp"

subcatlink(2,0)="../groups_clients/default.asp"
subcatlink(2,1)="../products/questions.asp"
subcatlink(2,2)="../appeals/app_list.asp"
subcatlink(2,3)="../products/products.asp"

subcatlink(1,0)="../groups_clients/default.asp"
subcatlink(1,1)="../pages/default.asp"
subcatlink(1,2)="../products/questions.asp"
subcatlink(1,3)="../products/products.asp"

subcatlink(0,0)="../companies/companies.asp"
subcatlink(0,1)="../companies/contacts.asp"
subcatlink(0,2)="../tasks/default.asp"

%>