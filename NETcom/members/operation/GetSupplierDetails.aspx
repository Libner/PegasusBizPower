<%@ Page Language="vb" AutoEventWireup="false" Codebehind="GetSupplierDetails.aspx.vb" Inherits="bizpower_pegasus.GetSupplierDetails" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>GetSupplierDetails</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<table border="0" cellpadding="1" cellspacing="1">
				<tr>
					<td align="right" width="330" nowrap>
						<span id="word3" name="word3"><%=func.vFix(supplier_Name)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="word3" name="word3"><!--שם סדרה--> 
								שם ספק</span></b>&nbsp;</td>
				</tr>
				<%if func.vFix(Country_Name)<>"" then%>
				<TR>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(Country_Name)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span1" name="word3">מדינה בה נותן 
								שירותים</span></b>&nbsp;</td>
				</TR>
				<%end if%>
				<%if func.vFix(supplier_Name1)="" and supplier_Job1="" and supplier_Tel1="" and  supplier_Phone1="" and  supplier_Email1="" and  supplier_Descr1="" then%>
				<TR>
					<td colspan="2" align="center">&nbsp;<b><span id="Span8" name="word4"><!--שם סוג--> 1 פרטי 
								איש קשר </span></b>&nbsp;</td>
					</TD>
				</TR>
				<%end if%>
				<%if func.vFix(supplier_Name1)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Name1)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span2" name="word4"><!--שם סוג--> 
								שם </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Job1)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						</span><%=func.vFix(supplier_Job1)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span9" name="word4"><!--שם סוג--> 
								תפקיד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<tr>
					<td align="right" width="330" nowrap> <span><%=func.vFix(supplier_Tel1)%></span> / <span><%=func.vFix(supplier_Ext1)%></span>
						
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span10" name="word4"><!--שם סוג--> 
								טלפון משרד </span></b>&nbsp;</td>
				</tr>
				<%if func.vFix(supplier_Phone1)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Phone1)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span11" name="word4"><!--שם סוג--> 
								נייד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
					<%if func.vFix(supplier_Email1)<>"" then%>
			
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Email1)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span12" name="word4"><!--שם סוג--> 
								מייל </span></b>&nbsp;</td>
				</tr>
				<%end if%>
					<%if func.vFix(supplier_Descr1)<>"" then%>
			
				<tr>
					<td align="right" width="330" nowrap>
					<span><%=func.vFix(supplier_Descr1)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span13" name="word4"><!--שם סוג--> 
								הערות</span></b>&nbsp;</td>
				</tr><%end if%>
				<%if func.vFix(trim(supplier_Name2))<>"" or func.vFix(trim(supplier_Job2))<>"" or func.vFix(trim(supplier_Tel2))<>"" or  func.vFix(trim(supplier_Phone2))<>"" or  func.vFix(trim(supplier_Email2))<>"" or  func.vFix(trim(supplier_Descr2))<>"" then%>
							<TR>
					<td colspan="2" align="center">&nbsp;<b><span id="Span14" name="word4"><!--שם סוג--> 2 פרטי 
								איש קשר </span></b>&nbsp;</td>
					</TD>
				</TR>
				<%end if%>
					<%if func.vFix(supplier_Name2)<>"" then%>
					<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Name2)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span15" name="word4"><!--שם סוג--> 
								שם </span></b>&nbsp;</td>
				</tr><%end if%>
				<%if func.vFix(supplier_Job2)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
					<span><%=func.vFix(supplier_Job2)%>-</span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span16" name="word4"><!--שם סוג--> 
								תפקיד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
					<%if func.vFix(supplier_Tel2)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap><span><%=func.vFix(supplier_Tel2)%></span> / 
						<span><%=func.vFix(supplier_Ext2)%></span>
					
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span17" name="word4"><!--שם סוג--> 
								טלפון משרד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
					<%if func.vFix(supplier_Phone2)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Phone2)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span18" name="word4"><!--שם סוג--> 
								נייד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Email2)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
					<span><%=func.vFix(supplier_Email2)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span19" name="word4"><!--שם סוג--> 
								מייל </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Descr2)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Descr2)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span20" name="word4"><!--שם סוג--> 
								הערות</span></b>&nbsp;</td>
				</tr>
				<%end if%>
			<%if func.vFix(trim(supplier_Name3))<>"" or func.vFix(trim(supplier_Job3))<>"" or func.vFix(trim(supplier_Tel3))<>"" or  func.vFix(trim(supplier_Phone3))<>"" or  func.vFix(trim(supplier_Email3))<>"" or  func.vFix(trim(supplier_Descr3))<>"" then%>
			
				<TR>
					<td colspan="2" align="center">&nbsp;<b><span id="Span21" name="word4"><!--שם סוג--> 3 פרטי 
								איש קשר </span></b>&nbsp;</td>
					</TD>
				</TR>
				<%end if%>
				<%if func.vFix(supplier_Name3)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
					<span><%=func.vFix(supplier_Name3)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span22" name="word4"><!--שם סוג--> 
								שם </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Job3)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Job3)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span23" name="word4"><!--שם סוג--> 
								תפקיד </span></b>&nbsp;</td>
				</tr><%end if%>
					<%if func.vFix(supplier_Tel3)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap><span><%=func.vFix(supplier_Tel3)%></span> / 
						<span><%=func.vFix(supplier_Ext3)%></span>
						 
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span24" name="word4"><!--שם סוג--> 
								טלפון משרד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
					<%if func.vFix(supplier_Phone3)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Phone3)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span25" name="word4"><!--שם סוג--> 
								נייד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
					<%if func.vFix(supplier_Email3)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Email3)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span26" name="word4"><!--שם סוג--> 
								מייל </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Descr3)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
					<span><%=func.vFix(supplier_Descr3)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span27" name="word4"><!--שם סוג--> 
								הערות</span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(trim(supplier_Name4))<>"" or func.vFix(trim(supplier_Job4))<>"" or func.vFix(trim(supplier_Tel4))<>"" or  func.vFix(trim(supplier_Phone4))<>"" or  func.vFix(trim(supplier_Email4))<>"" or  func.vFix(trim(supplier_Descr4))<>"" then%>
				<TR>
					<td colspan="2" align="center">&nbsp;<b><span id="Span28" name="word4"><!--שם סוג--> 4 פרטי 
								איש קשר </span></b>&nbsp;</td>
					</TD>
				</TR>
				<%end if%>
				<%if func.vFix(supplier_Name4)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Name4)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span29" name="word4"><!--שם סוג--> 
								שם </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Job4)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Job4)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span30" name="word4"><!--שם סוג--> 
								תפקיד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Tel4)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap><span><%=func.vFix(supplier_Tel4)%></span> / 
						<span><%=func.vFix(supplier_Ext4)%></span>
						
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span31" name="word4"><!--שם סוג--> 
								טלפון משרד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Phone4)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Phone4)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span32" name="word4"><!--שם סוג--> 
								נייד </span></b>&nbsp;</td>
				</tr>
				<%end if%>
					<%if func.vFix(supplier_Email4)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
						<span><%=func.vFix(supplier_Email4)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span33" name="word4"><!--שם סוג--> 
								מייל </span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<%if func.vFix(supplier_Descr4)<>"" then%>
				<tr>
					<td align="right" width="330" nowrap>
					<span><%=func.vFix(supplier_Descr4)%></span>
					</td>
					<td width="150" nowrap align="right">&nbsp;<b><span id="Span34" name="word4"><!--שם סוג--> 
								הערות</span></b>&nbsp;</td>
				</tr>
				<%end if%>
				<tr>
					<td colspan="2" height="15"></td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
