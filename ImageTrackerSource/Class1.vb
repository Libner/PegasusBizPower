Imports System
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Data.SqlClient

Namespace ImageTracker

    Public Class TrackRequest1
        Implements IHttpModule
        Private pattern As String = "/images/(?<key>.*)\.aspx"
        Private logoFile As String = "~/images/Powered-By-BP2.gif"
        Private dbConnectionString As String = "Server=(local);User=bizpower_beterem_user;Password=cyber;database=bizpower_beterem"

        Public Sub New()
        End Sub 'New


        '/ <summary>
        '/ Required by the interface IHttpModule
        '/ </summary>
        Public Sub Dispose() Implements System.Web.IHttpModule.Dispose

        End Sub 'Dispose


        '/ <summary>
        '/ Required by the interface IHttpModule
        '/ Wires up the BeginRequest event.
        '/ </summary>
        Public Sub Init(ByVal Appl As System.Web.HttpApplication) Implements System.Web.IHttpModule.Init

            AddHandler Appl.BeginRequest, AddressOf GetImage_BeginRequest
        End Sub 'Init


        '/ <summary>
        '/ Extracts the email key from the url and saves to the database. Also serves up the image
        '/ </summary>
        '/ <param name="sender">HttpApplication</param>
        '/ <param name="args">Not used</param>
        Public Sub GetImage_BeginRequest(ByVal sender As Object, ByVal args As System.EventArgs)
            'cast the sender to a HttpApplication object
            Dim application As System.Web.HttpApplication = CType(sender, System.Web.HttpApplication)

            Dim url As String = application.Request.Path 'get the url path
            'create the regex to match for becon images
            If url.IndexOf("CuteSoft_Client") < 1 Then

                Dim r As New Regex(pattern, RegexOptions.Compiled Or RegexOptions.IgnoreCase)
                If r.IsMatch(url) Then
                    Dim mc As MatchCollection = r.Matches(url)
                    If Not (mc Is Nothing) And mc.Count > 0 Then
                        Dim key As String = mc(0).Groups("key").Value
                        SaveToDB(key)
                    End If

                    'now send the image to the client
                    application.Response.ContentType = "image/gif"
                    application.Response.WriteFile(application.Request.MapPath(logoFile))

                    'end the resposne
                    application.Response.End()
                End If
            End If
        End Sub 'GetImage_BeginRequest


        '/ <summary>
        '/ saves the key to the database
        '/ </summary>
        '/ <param name="key">key database value</param>
        Private Sub SaveToDB(ByVal key As String)
            If key Is Nothing Or key.Trim().Length = 0 Then
                Return
            End If
            'normally you would use a stored procedure, but the sql statement is written out for simplicity
            Dim arr As Array
            arr = key.Split("_")
            If IsArray(arr) And UBound(arr) > 0 Then
                Dim prod_id = arr(0)
                Dim client_id = arr(1)
                If IsNumeric(prod_id) And IsNumeric(client_id) Then
                    Dim sqlText As String = String.Format("UPDATE PRODUCT_CLIENT SET IS_OPENED = 1 WHERE PRODUCT_ID = " & prod_id & " AND PEOPLE_ID = " & client_id)
                    Dim cmd As New SqlCommand(sqlText, New SqlConnection(dbConnectionString))
                    cmd.Connection.Open()
                    cmd.ExecuteNonQuery()
                    cmd.Connection.Close()
                End If
            End If
        End Sub 'SaveToDB 
    End Class 'TrackRequest1
End Namespace 'ImageTracker 