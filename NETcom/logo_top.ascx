<%@ Control Language="vb" AutoEventWireup="false" Codebehind="logo_top.ascx.vb" Inherits="bizpower_pegasus2018.logo_top_webflowCSS" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
		<%if ShowLogo = "1" then%>
		<%Dim show_logo as Boolean
  Dim align_logo as String
  show_logo = true : align_logo = "center"
  If Not isNothing(Request.Cookies("bizpegasus")("UseBizLogo")) Then
	If trim(Request.Cookies("bizpegasus")("UseBizLogo")) = "0" Then 
		show_logo = false : align_logo = "left"
	End If	
  End If%>
		<div class="biz_nav_section">
			<div class="bizpower_top">
				<div class="biz_pegasus_login">
					<div class="biz_login_name"><%=trim(user_name)%></div>
					<div></div>
					<a href="<%=Application("VirDir")%>/exit.asp" target="_self" class="biz_exit">יציאה</a>
				</div>
				<div class="biz_logo_group w-clearfix">
					<div class="biz_logo"><img src="<%=Application("VirDir")%>/NETcom/webflowCSS/images/top_biz_logo.png" alt="" style="max-width: 100%;"></div>
					<div class="biz_logo_txt">pegasus</div>
					<div class="biz_logo_line">/</div>
				</div>
			</div>
		</div>
		<%end if ' show logo%>
