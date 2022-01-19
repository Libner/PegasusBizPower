Imports System.Web.HttpContext
Imports System.Text
Imports System.IO
Imports System.Web.SessionState
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Services.Protocols
Imports System.Net
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Uri
Namespace travelBoost

    Public Class funcs
        '======================
        '   send XML in format:
        '   <root>
        '          <header>
        '                <lang>Eng</lang>
        '                <xmlVersion>11</xmlVersion>
        '                <oem>0000</oem>
        '                <agentID>0000</agentID>
        '                <password>0000</password>
        '          </header>
        '          <body>
        '            <opCode>Op-Get-Product</opCode>
        '          </body>
        '   </root>

        Public Function fixNumeric(ByVal objVal As Object) As Double
            Dim num As Double
            Try

                If IsNothing(objVal) Then
                    num = 0
                Else
                    If IsNumeric(objVal) Then
                        num = CDbl(objVal)
                    Else
                        num = 0
                    End If
                End If
            Catch ex As Exception
                num = 0
            End Try
            fixNumeric = num
        End Function
        Function dbNullDate(ByVal objVal As Object, ByVal formatString As String) As String
            Try
                If IsNothing(objVal) Then
                    dbNullDate = ""
                ElseIf IsDBNull(objVal) Then
                    dbNullDate = ""
                Else
                    If IsDate(objVal) Then
                        If formatString <> "" Then
                            dbNullDate = CDate(objVal).ToString(formatString)
                        Else
                            dbNullDate = CDate(objVal)
                        End If
                    End If
                End If

            Catch ex As Exception
                dbNullDate = ""
            End Try
        End Function
        Function dbNullFix(ByVal objVal As Object) As String
            Try

                If IsNothing(objVal) Then
                    dbNullFix = ""
                Else
                    If IsDBNull(objVal) Then
                        dbNullFix = ""
                    Else
                        Try
                            dbNullFix = Trim(objVal.ToString)
                        Catch ex As Exception
                            dbNullFix = ""
                        End Try
                    End If
                End If
            Catch ex As Exception
                dbNullFix = ""
            End Try
        End Function
        Function sendXMLRequest(ByVal url As String, ByVal postXMLData As System.Xml.XmlDocument) As System.Xml.XmlDocument
            ''   Dim WSUri As New System.Uri(url, UriKind.Absolute)
            HttpContext.Current.Response.Write("<br>sendXMLRequest")
            HttpContext.Current.Response.End()
            Dim objRequest As HttpWebRequest = CType(HttpWebRequest.Create(url), HttpWebRequest)
            Dim StrData As String

            StrData = postXMLData.OuterXml

            objRequest.Method = "POST"
            Dim byte1 As Byte() = System.Text.Encoding.UTF8.GetBytes(StrData)
            objRequest.ContentLength = byte1.Length
            objRequest.ContentType = "application/x-www-form-urlencoded" '"text/xml"
            objRequest.KeepAlive = False

            Dim streamToSend As System.IO.Stream = objRequest.GetRequestStream()
            streamToSend.Write(byte1, 0, byte1.Length)
            streamToSend.Close()


            Dim objResponse As HttpWebResponse = objRequest.GetResponse()

            Dim StreamReader As System.IO.StreamReader
            StreamReader = New System.IO.StreamReader(objResponse.GetResponseStream())

            Dim xmldoc As New System.Xml.XmlDocument

            xmldoc.LoadXml(StreamReader.ReadToEnd())
            HttpContext.Current.Response.Write("<br>xmldoc=" & xmldoc.OuterXml)
            ' xmldoc.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download") & "/getLogin_" & Date.Now.Minute & ".xml")
            sendXMLRequest = xmldoc

        End Function


        Function getAreas() As XmlDocument
            Dim xmldocOutPut As New XmlDocument
            'xmldocOutPut.AppendChild(xmldocOutPut.CreateXmlDeclaration("1.0", "utf-8", Nothing))


            Dim nodeLevel1, nodeLevel2, nodeLevel3, nodeLevel4 As XmlNode

            nodeLevel1 = xmldocOutPut.CreateElement("root")
            'header
            nodeLevel2 = xmldocOutPut.CreateElement("header")

            nodeLevel3 = xmldocOutPut.CreateElement("oem")
            nodeLevel3.InnerText = "G4I_B2B"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("agentID")
            nodeLevel3.InnerText = "XML"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("password")
            nodeLevel3.InnerText = "cyber@xml1"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("xmlVersion")
            nodeLevel3.InnerText = "10"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            'body
            nodeLevel2 = xmldocOutPut.CreateElement("body")



            nodeLevel3 = xmldocOutPut.CreateElement("opCode")
            nodeLevel3.InnerText = "OP_GET_AREAS"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("servType")
            nodeLevel3.InnerText = "HOTEL"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            xmldocOutPut.AppendChild(nodeLevel1)

            xmldocOutPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/req_Areas.xml")

            Dim xmldocInPut As New XmlDocument
            xmldocInPut = sendXMLRequest(HttpContext.Current.Application("T_URL"), xmldocOutPut)
            xmldocInPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/resp_Areas.xml")

            getAreas = xmldocInPut
        End Function

        Function getPriceLevelsTours() As XmlDocument
            Dim xmldocOutPut As New XmlDocument
            'xmldocOutPut.AppendChild(xmldocOutPut.CreateXmlDeclaration("1.0", "utf-8", Nothing))


            Dim nodeLevel1, nodeLevel2, nodeLevel3, nodeLevel4 As XmlNode

            nodeLevel1 = xmldocOutPut.CreateElement("root")
            'header
            nodeLevel2 = xmldocOutPut.CreateElement("header")

            nodeLevel3 = xmldocOutPut.CreateElement("oem")
            nodeLevel3.InnerText = "G4I_B2B"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("agentID")
            nodeLevel3.InnerText = "XML"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("password")
            nodeLevel3.InnerText = "cyber@xml1"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("xmlVersion")
            nodeLevel3.InnerText = "10"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            'body
            nodeLevel2 = xmldocOutPut.CreateElement("body")



            nodeLevel3 = xmldocOutPut.CreateElement("opCode")
            nodeLevel3.InnerText = "OP_GET_PRICELEVELS"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("servType")
            nodeLevel3.InnerText = "TOUR"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            xmldocOutPut.AppendChild(nodeLevel1)

            xmldocOutPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/req_PriceLevelsTours.xml")

            Dim xmldocInPut As New XmlDocument
            xmldocInPut = sendXMLRequest(HttpContext.Current.Application("T_URL"), xmldocOutPut)
            xmldocInPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/resp_PriceLevelsTours.xml")

            getPriceLevelsTours = xmldocInPut
        End Function
        Function getPriceLevelsRooms() As XmlDocument
            Dim xmldocOutPut As New XmlDocument
            'xmldocOutPut.AppendChild(xmldocOutPut.CreateXmlDeclaration("1.0", "utf-8", Nothing))


            Dim nodeLevel1, nodeLevel2, nodeLevel3, nodeLevel4 As XmlNode

            nodeLevel1 = xmldocOutPut.CreateElement("root")
            'header
            nodeLevel2 = xmldocOutPut.CreateElement("header")

            nodeLevel3 = xmldocOutPut.CreateElement("oem")
            nodeLevel3.InnerText = "G4I_B2B"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("agentID")
            nodeLevel3.InnerText = "XML"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("password")
            nodeLevel3.InnerText = "cyber@xml1"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("xmlVersion")
            nodeLevel3.InnerText = "10"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            'body
            nodeLevel2 = xmldocOutPut.CreateElement("body")



            nodeLevel3 = xmldocOutPut.CreateElement("opCode")
            nodeLevel3.InnerText = "OP_GET_PRICELEVELS"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("servType")
            nodeLevel3.InnerText = "HOTEL"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            xmldocOutPut.AppendChild(nodeLevel1)

            xmldocOutPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/req_PriceLevelsRooms.xml")

            Dim xmldocInPut As New XmlDocument
            xmldocInPut = sendXMLRequest(HttpContext.Current.Application("T_URL"), xmldocOutPut)
            xmldocInPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/resp_PriceLevelsRooms.xml")

            getPriceLevelsRooms = xmldocInPut
        End Function

        Function getVacations() As XmlDocument
            Dim xmldocOutPut As New XmlDocument
            'xmldocOutPut.AppendChild(xmldocOutPut.CreateXmlDeclaration("1.0", "utf-8", Nothing))


            Dim nodeLevel1, nodeLevel2, nodeLevel3, nodeLevel4 As XmlNode

            nodeLevel1 = xmldocOutPut.CreateElement("root")
            'header
            nodeLevel2 = xmldocOutPut.CreateElement("header")

            nodeLevel3 = xmldocOutPut.CreateElement("oem")
            nodeLevel3.InnerText = "G4I_B2B"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("agentID")
            nodeLevel3.InnerText = "XML"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("password")
            nodeLevel3.InnerText = "cyber@xml1"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("xmlVersion")
            nodeLevel3.InnerText = "10"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            'body
            nodeLevel2 = xmldocOutPut.CreateElement("body")



            nodeLevel3 = xmldocOutPut.CreateElement("opCode")
            nodeLevel3.InnerText = "OP_GET_VACATIONS"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("startDate")
            nodeLevel3.InnerText = "2018-07-01"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("origAreaID")
            nodeLevel3.InnerText = ""
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("areaID")
            nodeLevel3.InnerText = "292"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("deltaDays")
            nodeLevel3.InnerText = "0"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("currency")
            nodeLevel3.InnerText = "EUR"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("getAgentComm")
            nodeLevel3.InnerText = "0"
            nodeLevel2.AppendChild(nodeLevel3)



            'nodeLevel3 = xmldocOutPut.CreateElement("opCode")
            'nodeLevel3.InnerText = "OP_GET_PRODUCT"
            'nodeLevel2.AppendChild(nodeLevel3)

            'nodeLevel3 = xmldocOutPut.CreateElement("TourCodeA")
            'nodeLevel3.InnerText = "VBB"
            'nodeLevel2.AppendChild(nodeLevel3)

            'nodeLevel3 = xmldocOutPut.CreateElement("TourCodeB")
            'nodeLevel3.InnerText = "0920"
            'nodeLevel2.AppendChild(nodeLevel3)

            'nodeLevel3 = xmldocOutPut.CreateElement("dateFrom")
            'nodeLevel3.InnerText = "2018-09-20"
            'nodeLevel2.AppendChild(nodeLevel3)


            nodeLevel1.AppendChild(nodeLevel2)

            xmldocOutPut.AppendChild(nodeLevel1)

            xmldocOutPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/req_Vacations.xml")

            Dim xmldocInPut As New XmlDocument
            xmldocInPut = sendXMLRequest(HttpContext.Current.Application("T_URL"), xmldocOutPut)
            xmldocInPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/resp_Vacations.xml")

            getVacations = xmldocInPut
        End Function


        Function getDeparture(ByVal seriesCode As String, ByVal departureDate As Date, ByVal departureDateEnd As Date, ByVal roomCount As Integer, ByVal paxPerRoom As Integer, ByVal chdPerRoom As Integer) As XmlDocument
            Dim xmldocOutPut As New XmlDocument
            xmldocOutPut.AppendChild(xmldocOutPut.CreateXmlDeclaration("1.0", "utf-8", Nothing))


            Dim nodeLevel1, nodeLevel2, nodeLevel3, nodeLevel4 As XmlNode
            Dim elLevel4 As XmlElement

            nodeLevel1 = xmldocOutPut.CreateElement("root")
            'header
            nodeLevel2 = xmldocOutPut.CreateElement("header")

            nodeLevel3 = xmldocOutPut.CreateElement("oem")
            nodeLevel3.InnerText = HttpContext.Current.Application("T_oem")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("agentID")
            nodeLevel3.InnerText = HttpContext.Current.Application("T_agentID")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("password")
            nodeLevel3.InnerText = HttpContext.Current.Application("T_password")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("xmlVersion")
            nodeLevel3.InnerText = HttpContext.Current.Application("T_xmlVersion")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            'body
            nodeLevel2 = xmldocOutPut.CreateElement("body")

            nodeLevel3 = xmldocOutPut.CreateElement("opCode")
            nodeLevel3.InnerText = "OP_GET_TOURS"
            nodeLevel2.AppendChild(nodeLevel3)

            'nodeLevel3 = xmldocOutPut.CreateElement("areaID")
            'nodeLevel3.InnerText = "2624"
            'nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("depDateFrom")
            nodeLevel3.InnerText = departureDate.ToString("yyyy-MM-dd") '"2018-09-27"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("depDateTo")
            nodeLevel3.InnerText = departureDateEnd.ToString("yyyy-MM-dd") '"2018-09-27"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("tourCode1")
            nodeLevel3.InnerText = seriesCode '"VBB"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("tourCode2")
            nodeLevel3.InnerText = departureDate.ToString("MMdd") '"0920"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("currency")
            nodeLevel3.InnerText = departureDate.ToString("USD")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("roomDescCount")
            nodeLevel4 = xmldocOutPut.CreateElement("roomDesc")
            nodeLevel4.InnerText = roomCount '"1"by default
            Dim attr As XmlAttribute
            attr = xmldocOutPut.CreateAttribute("paxPerRoom")
            attr.Value = paxPerRoom '"1" by default
            nodeLevel4.Attributes.Append(attr)
            attr = xmldocOutPut.CreateAttribute("chdPerRoom")
            attr.Value = chdPerRoom '"0" by default
            nodeLevel4.Attributes.Append(attr)

            nodeLevel3.AppendChild(nodeLevel4)
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)



            xmldocOutPut.AppendChild(nodeLevel1)
            xmldocOutPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/req_Departure.xml")

            Dim xmldocInPut As New XmlDocument
            xmldocInPut = sendXMLRequest(HttpContext.Current.Application("T_URL"), xmldocOutPut)
            xmldocInPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/resp_Departure_" & seriesCode & "-" & departureDate.ToString("MMdd") & ".xml")

            getDeparture = xmldocInPut

        End Function


        Function getTourDepartures(ByVal seriesCode As String) As XmlDocument

            Dim xmldocOutPut As New XmlDocument
            xmldocOutPut.AppendChild(xmldocOutPut.CreateXmlDeclaration("1.0", "utf-8", Nothing))


            Dim nodeLevel1, nodeLevel2, nodeLevel3, nodeLevel4 As XmlNode
            Dim elLevel4 As XmlElement

            nodeLevel1 = xmldocOutPut.CreateElement("root")
            'header
            nodeLevel2 = xmldocOutPut.CreateElement("header")

            nodeLevel3 = xmldocOutPut.CreateElement("oem")
            nodeLevel3.InnerText = HttpContext.Current.Application("T_oem")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("agentID")
            nodeLevel3.InnerText = HttpContext.Current.Application("T_agentID")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("password")
            nodeLevel3.InnerText = HttpContext.Current.Application("T_password")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("xmlVersion")
            nodeLevel3.InnerText = HttpContext.Current.Application("T_xmlVersion")
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            'body
            nodeLevel2 = xmldocOutPut.CreateElement("body")

            nodeLevel3 = xmldocOutPut.CreateElement("opCode")
            nodeLevel3.InnerText = "OP_GET_TOURS"
            nodeLevel2.AppendChild(nodeLevel3)

            'nodeLevel3 = xmldocOutPut.CreateElement("areaID")
            'nodeLevel3.InnerText = "2624"
            'nodeLevel2.AppendChild(nodeLevel3)


            nodeLevel3 = xmldocOutPut.CreateElement("depDateFrom")
            nodeLevel3.InnerText = Date.Today.ToString("yyyy-MM-dd") '"2018-09-27"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("depDateTo")
            nodeLevel3.InnerText = DateAdd(DateInterval.Year, 2, Date.Today).ToString("yyyy-MM-dd") '"2018-09-27"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("tourCode1")
            nodeLevel3.InnerText = seriesCode '"VBB"
            nodeLevel2.AppendChild(nodeLevel3)

            '' <priceLevelID>47</priceLevelID><priceLevelName>קהל רחב</priceLevelName> 
            'nodeLevel3 = xmldocOutPut.CreateElement("priceLevelID")
            'nodeLevel3.InnerText = "47"
            'nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("currency")
            nodeLevel3.InnerText = "USD"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("roomDescCount")
            nodeLevel4 = xmldocOutPut.CreateElement("roomDesc")
            nodeLevel4.InnerText = "1" 'by default

            Dim attr As XmlAttribute
            attr = xmldocOutPut.CreateAttribute("paxPerRoom")
            attr.Value = "2" 'by default - double room - get price for one pax in the double room (total price /2)
            nodeLevel4.Attributes.Append(attr)
            attr = xmldocOutPut.CreateAttribute("chdPerRoom")
            attr.Value = "0" 'by default
            nodeLevel4.Attributes.Append(attr)

            nodeLevel3.AppendChild(nodeLevel4)
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            xmldocOutPut.AppendChild(nodeLevel1)
            xmldocOutPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/req_Tour_" & seriesCode & ".xml")

            Dim xmldocInPut As New XmlDocument
            xmldocInPut = sendXMLRequest(HttpContext.Current.Application("T_URL"), xmldocOutPut)
            xmldocInPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/resp_Tour_" & seriesCode & ".xml")

            getTourDepartures = xmldocInPut
        End Function

        Function test() As XmlDocument
            Dim xmldocOutPut As New XmlDocument
            xmldocOutPut.AppendChild(xmldocOutPut.CreateXmlDeclaration("1.0", "utf-8", Nothing))


            Dim nodeLevel1, nodeLevel2, nodeLevel3, nodeLevel4 As XmlNode

            nodeLevel1 = xmldocOutPut.CreateElement("root")
            'header
            nodeLevel2 = xmldocOutPut.CreateElement("header")

            nodeLevel3 = xmldocOutPut.CreateElement("oem")
            nodeLevel3.InnerText = "G4I_B2B"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("agentID")
            nodeLevel3.InnerText = "XML"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("password")
            nodeLevel3.InnerText = "cyber@xml1"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel3 = xmldocOutPut.CreateElement("xmlVersion")
            nodeLevel3.InnerText = "10"
            nodeLevel2.AppendChild(nodeLevel3)

            nodeLevel1.AppendChild(nodeLevel2)

            'body
            nodeLevel2 = xmldocOutPut.CreateElement("body")

            nodeLevel3 = xmldocOutPut.CreateElement("opCode")
            nodeLevel3.InnerText = "OP_GET_PRODUCT"
            nodeLevel2.AppendChild(nodeLevel3)


            nodeLevel1.AppendChild(nodeLevel2)

            xmldocOutPut.AppendChild(nodeLevel1)
            xmldocOutPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/req_Test.xml")

            Dim xmldocInPut As New XmlDocument
            xmldocInPut = sendXMLRequest("http://b2b-pegasusuat.travelbooster.com/UI_NET/XmlServerApi.aspx?FromStartDate=16Aug18&ToStartDate=23Aug18&TourCodeA=VBB&TourCodeB=0816&NumOfAdult=2", xmldocOutPut)
            xmldocInPut.Save(HttpContext.Current.Server.MapPath(HttpContext.Current.Application("virDir") & "/download/travelBooster") & "/resp_Test.xml")

            test = xmldocInPut
        End Function
    End Class

    '  Class RijndaelFuncs

        '    Shared Function EncryptStringToBytes(ByVal plainText As String, ByVal Key() As Byte, ByVal IV() As Byte) As Byte()
        '        ' Check arguments.
        '        If plainText Is Nothing OrElse plainText.Length <= 0 Then
        '            Throw New ArgumentNullException("plainText")
        '        End If
        '        If Key Is Nothing OrElse Key.Length <= 0 Then
        '            Throw New ArgumentNullException("Key")
        '        End If
        '        If IV Is Nothing OrElse IV.Length <= 0 Then
        '            Throw New ArgumentNullException("IV")
        '        End If
        '        Dim encrypted() As Byte
        '        ' Create an Rijndael object
        '        ' with the specified key and IV.
        '        Using(rijAlg = Rijndael.Create())

        '        rijAlg.Key = Key
        '        rijAlg.IV = IV

        '        ' Create an encryptor to perform the stream transform.
        '        Dim encryptor As ICryptoTransform = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV)
        '        ' Create the streams used for encryption.
        '            Using msEncrypt As New MemoryStream()
        '                Using csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
        '                    Using swEncrypt As New StreamWriter(csEncrypt)

        '        'Write all data to the stream.
        '        swEncrypt.Write(plainText)
        '                    End Using
        '        encrypted = msEncrypt.ToArray()
        '                End Using
        '            End Using
        '        End Using

        '        ' Return the encrypted bytes from the memory stream.
        '        Return encrypted

        '    End Function 'EncryptStringToBytes

        '    Shared Function DecryptStringFromBytes(ByVal cipherText() As Byte, ByVal Key() As Byte, ByVal IV() As Byte) As String
        '        ' Check arguments.
        '        If cipherText Is Nothing OrElse cipherText.Length <= 0 Then
        '            Throw New ArgumentNullException("cipherText")
        '        End If
        '        If Key Is Nothing OrElse Key.Length <= 0 Then
        '            Throw New ArgumentNullException("Key")
        '        End If
        '        If IV Is Nothing OrElse IV.Length <= 0 Then
        '            Throw New ArgumentNullException("IV")
        '        End If
        '        ' Declare the string used to hold
        '        ' the decrypted text.
        '        Dim plaintext As String = Nothing

        '        ' Create an Rijndael object
        '        ' with the specified key and IV.
        '        Using(rijAlg = Rijndael.Create())
        '        rijAlg.Key = Key
        '        rijAlg.IV = IV

        '        ' Create a decryptor to perform the stream transform.
        '        Dim decryptor As ICryptoTransform = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV)

        '        ' Create the streams used for decryption.
        '            Using msDecrypt As New MemoryStream(cipherText)

        '                Using csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)

        '                    Using srDecrypt As New StreamReader(csDecrypt)


        '        ' Read the decrypted bytes from the decrypting stream
        '        ' and place them in a string.
        '        plaintext = srDecrypt.ReadToEnd()
        '                    End Using
        '                End Using
        '            End Using
        '        End Using

        '        Return plaintext

        '    End Function 'DecryptStringFromBytes 
        'End Class


End Namespace