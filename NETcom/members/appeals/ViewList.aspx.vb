Imports System.Data
Imports System.Data.SqlClient

Public Class ViewList
    Inherits System.Web.UI.Page

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

    Protected WithEvents RLi As Repeater
    Protected PHPagerList, PHPagerPage As System.Web.UI.WebControls.PlaceHolder

    Protected IsDataBound As Boolean = False
    Dim LoginMembercode, sessionID As String
    Protected siteAudience As String
    Dim currentPageNumber, recordsPerPage, returnNumberOfPages As Integer

    Protected WithEvents txtSearch As System.Web.UI.WebControls.TextBox

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim aCookie As System.Web.HttpCookie
        aCookie = Request.Cookies("SL")

        If Not (aCookie Is Nothing) Then
            If Trim(aCookie.Values("membercode")) <> "" Then
                LoginMembercode = CDbl(aCookie.Values("membercode"))
                sessionID = Server.UrlDecode(Request.Cookies("SL")("MSG"))
                siteAudience = aCookie.Values("siteAudience")
            Else
                LoginMembercode = ""
                sessionID = ""
            End If

        Else
            LoginMembercode = ""
            sessionID = ""
        End If

        If Trim(LoginMembercode) = "" Then
            Response.End()
        End If

        If Trim(LoginMembercode) <> "" Then

            If Trim(Request.QueryString("Page")) <> "" Then
                currentPageNumber = Request.QueryString("Page")
            Else
                currentPageNumber = 1
            End If

            If Trim(Request.Form("txtSearch")) <> "" Then
                currentPageNumber = 1
                recordsPerPage = 100
            Else
                recordsPerPage = 32
            End If

            Dim myConnection As SqlConnection = New SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
            Dim myCommand As SqlCommand

            Dim sqlStr As String

            myCommand = New SqlCommand("usp_selectMemberMemoListNew", myConnection)
            myCommand.CommandType = CommandType.StoredProcedure

            Dim parameterLoginMembercode As SqlParameter = New SqlParameter("@login_membercode", SqlDbType.Int, 4)
            parameterLoginMembercode.Value = LoginMembercode
            myCommand.Parameters.Add(parameterLoginMembercode)

            Dim parameterCurrentPageNumber As SqlParameter = New SqlParameter("@currentPageNumber", SqlDbType.Int, 4)
            parameterCurrentPageNumber.Value = currentPageNumber
            myCommand.Parameters.Add(parameterCurrentPageNumber)

            Dim parameterRecordsPerPage As SqlParameter = New SqlParameter("@RecordsPerPage", SqlDbType.Int, 4)
            parameterRecordsPerPage.Value = recordsPerPage
            myCommand.Parameters.Add(parameterRecordsPerPage)

            Dim parameterOutput_returnNumberOfPages As New SqlParameter("@returnNumberOfPages", SqlDbType.Int, 4)
            parameterOutput_returnNumberOfPages.Direction = ParameterDirection.Output
            myCommand.Parameters.Add(parameterOutput_returnNumberOfPages)

            Dim parameterTxtSearch As SqlParameter = New SqlParameter("@txtsearch", SqlDbType.VarChar, 50)
            parameterTxtSearch.Value = Trim(killChars(Request.Form("txtSearch")))
            myCommand.Parameters.Add(parameterTxtSearch)
            Try
                myConnection.Open()
                RLi.DataSource = myCommand.ExecuteReader()
                RLi.DataBind()
                If RLi.Items.Count = 0 Then
                    RLi.Visible = False
                Else
                    RLi.Visible = True
                End If
            Finally
                myCommand.Dispose()
                myConnection.Close()
                myConnection.Dispose()
            End Try

            returnNumberOfPages = parameterOutput_returnNumberOfPages.Value
        End If

        If returnNumberOfPages > 1 Then
            PHPagerList.Visible = True
            PHPagerPage.Controls.Add(New LiteralControl(pagerList(currentPageNumber, returnNumberOfPages, 6)))
        Else
            PHPagerList.Visible = False
        End If
    End Sub

    Private Sub RLi_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles RLi.ItemDataBound
        Dim divPic, divNW, divMembername As System.Web.UI.HtmlControls.HtmlGenericControl
        Dim lblMemo_membername, lblMemoPhone, lblMemoText As Label
        Dim PHMessage As System.Web.UI.WebControls.PlaceHolder

        Dim strImgOnline, intoCode, strIntoCode, altIntoCode, altSearchIntoCode As String
        Dim strPicName As String = ""

        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            IsDataBound = True

            If Not IsDBNull(e.Item.DataItem("picName")) And Not IsDBNull(e.Item.DataItem("picType")) Then
                If e.Item.DataItem("picType") = 0 Then
                    strPicName = "/images/pics/nophoto-100-100.jpg"
                ElseIf e.Item.DataItem("picType") = 2 Then
                    strPicName = "/images/pics/private-100-100.jpg"
                ElseIf e.Item.DataItem("picType") = 3 Then
                    strPicName = "/images/pics/x-100-100.jpg"
                ElseIf e.Item.DataItem("picType") = 1 Then
                    strPicName = Application("atrafdatingCentralSiteUrl") & "/crop100/" & e.Item.DataItem("picName")
                End If
            Else
                strPicName = "/images/pics/nophoto-100-100.jpg"
            End If

            divPic = e.Item.FindControl("divPic")
            divPic.Attributes.Add("style", "background-image: url(" & strPicName & ");")
            divPic.Attributes.Add("href", "/members/memberDetails.aspx?" & EnCrypt("membercode=" & e.Item.DataItem("membercode")))

            divNW = e.Item.FindControl("divNW")
            divNW.InnerHtml = "<a href="""" onclick=""return memberDetailsWindow('" & EnCrypt("membercode=" & e.Item.DataItem("membercode")) & "')""><img src='/images/nw.gif' border=0 vspace=""3""></a>"

            If e.Item.DataItem("online") & "" = "1" Then
                strImgOnline = "<img src=""/images/online.png"" class=""va-m-l"">"
            Else
                strImgOnline = "<img src=""/images/offline.png"" class=""va-m-l"">"
            End If

            divMembername = e.Item.FindControl("divMembername")
            divMembername.InnerHtml = strImgOnline & "<a href=""../members/memberDetails.aspx?" & EnCrypt("membercode=" & e.Item.DataItem("membercode")) & """ class=""lcrop-n"" dir=""rtl""><span dir=""rtl"">" & vFix(e.Item.DataItem("memberName") & "") & "</span></a>"

            lblMemo_membername = e.Item.FindControl("lblMemo_membername")
            If Trim(e.Item.DataItem("memo_membername") & "") <> "" Then
                lblMemo_membername.Text = "שם: <b>" & e.Item.DataItem("memo_membername") & "" & "</b><br>"
            End If

            lblMemoPhone = e.Item.FindControl("lblMemoPhone")
            If Trim(e.Item.DataItem("memo_phone") & "") <> "" Then
                lblMemoPhone.Text = "טלפון: <b>" & e.Item.DataItem("memo_phone") & "" & "</b><br>"
            End If

            lblMemoText = e.Item.FindControl("lblMemoText")
            If Not IsDBNull(e.Item.DataItem("memo_text")) Then
                lblMemoText.Text = HtmlBreak(e.Item.DataItem("memo_text"))
                ''lblMemoText.Attributes.Add("title", e.Item.DataItem("memo_text"))
                ''lblMemoText.Attributes.Add("alt", e.Item.DataItem("memo_text"))
            End If

            PHMessage = e.Item.FindControl("PHMessage")
            If Not IsDBNull(e.Item.DataItem("membercode")) Then
                If e.Item.DataItem("online") & "" = "1" Then
                    PHMessage.Controls.Add(New LiteralControl("<div class=""lcrop-ms-on""><a class=""lcrop-ms-on"" href="""" ONCLICK=""return popupjMessage(" & e.Item.DataItem("membercode") & ");"" dir=rtl >שלח הודעה</a></div>"))
                Else
                    PHMessage.Controls.Add(New LiteralControl("<div class=""lcrop-ms-off""><a class=""lcrop-ms-off"" href="""" ONCLICK=""return popupjMessage(" & e.Item.DataItem("membercode") & ");"" dir=rtl >שלח הודעה</a></div>"))
                End If
                PHMessage.Controls.Add(New LiteralControl("<div class=""lcrop-action lcrop-action-plus""><a title=""עדכן תזכורת"" href='' onclick=""return memo_window(" & e.Item.DataItem("membercode") & "," & e.Item.DataItem("memoID") & ",'','" & LoginMembercode & "','" & sessionID & "');"" class=""lcrop-action"">עדכן תזכורת</a></div>"))
            Else
                PHMessage.Controls.Add(New LiteralControl("<div class=""lcrop-action lcrop-action-plus""><a title=""עדכן תזכורת"" href='' onclick=""return memo_window(''," & e.Item.DataItem("memoID") & ",'','" & LoginMembercode & "','" & sessionID & "');"" class=""lcrop-action"">עדכן תזכורת</a></div>"))
            End If

            PHMessage.Controls.Add(New LiteralControl("<div class=""lcrop-action lcrop-action-del""><a title=""מחק תזכורת"" href=""javascript:DeleteMemo('" & e.Item.DataItem("memoID") & "','" & LoginMembercode & "','" & sessionID & "');"" class=""lcrop-action"">מחק תזכורת</a></div>"))
        End If

        If e.Item.ItemType = ListItemType.Separator Then
            If (e.Item.ItemIndex + 1) Mod 8 = 0 Then
                Dim cntrl As System.Web.UI.LiteralControl = e.Item.Controls(0)
                If (e.Item.ItemIndex + 1) = 16 Or siteAudience = "girls" Then
                    cntrl.Text = "<div class=""lcrop-h-sep""></div><div class=""lcrop-line""></div><div class=""gallery-banner"" align=""center""><iframe src=""/banners/getNormalBanner.aspx"" allowtransparency=""true"" width=""468"" height=""70"" frameborder=""0"" scrolling=""no""></iframe></div><div class=""lcrop-h-sep""></div><div class=""lcrop-line""></div>"
                Else
                    cntrl.Text = "<div class=""lcrop-h-sep""></div><div class=""lcrop-line""></div><div class=""gallery-banner"" align=""center""><iframe src=""/banners/getPirsumBanner.aspx"" allowtransparency=""true"" width=""590"" height=""70"" frameborder=""0"" scrolling=""no""></iframe></div><div class=""lcrop-h-sep""></div><div class=""lcrop-line""></div>"
                End If
            End If
        End If
    End Sub
End Class
