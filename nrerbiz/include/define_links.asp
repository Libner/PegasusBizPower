<%
dim catname(6)
catname(0)= "קשרי לקוחות"
catname(1)= "דיוור"
catname(2)= "משובים וסקרים"
catname(3)= "דוחות"
catname(4)= "כלכלה ופיננסים"
catname(5)= "אשפים"

dim subcatname(10,5)

subcatname(10,0)="פרטי ארגון"
subcatname(10,1)="המחרה"

subcatname(9,3)="ניהול משימות"
subcatname(9,2)="פרויקט לקוח"
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
subcatname(5,4)="כלכלה ופיננסים"  ' +

subcatname(4,5)="Bizpower desktop"
subcatname(4,4)="תזרים מזומנים"
subcatname(4,3)="שעות העבודה שלי"
subcatname(4,2)="פרויקטים עתידיים ותמחור"
subcatname(4,1)="פרויקטים ועלות עבודה"
subcatname(4,0)="עובדים ותפקידים"

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

dim subcatlink(11,5)

subcatlink(6,4)="../../default.asp?wizard_id=5"
subcatlink(6,3)="../../default.asp?wizard_id=2"
subcatlink(6,2)="../../default.asp?wizard_id=4"
subcatlink(6,1)="../../default.asp?wizard_id=3"
subcatlink(6,0)="../../default.asp?wizard_id=1"

subcatlink(10,0)="../../default.asp?wizard_id=6"  'פרטי ארגון 
subcatlink(10,1)="../../default.asp?wizard_id=7"  ' המחרה 

subcatlink(5,1)=null
subcatlink(5,0)=null

subcatlink(4,5)="../projects/download.asp"
subcatlink(4,4)="../projects/cash_flow.asp"
subcatlink(4,3)="../projects/hours_calendar.asp"
subcatlink(4,2)="../projects/prices.asp"
subcatlink(4,1)="../projects/default.asp"
subcatlink(4,0)="../workers/default.asp"

subcatlink(3,0)="../reports/reports.asp"
subcatlink(3,1)="../projects/projects_reports.asp"

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