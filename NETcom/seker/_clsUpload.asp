<%
Class clsUpload

' This class is produced
' by 20/20 Applications.
' Copyright (c) 2002 by
' 20/20 Applications of
' www.2020applications.com.

' This code is modelled
' after similar scripts found
' on the internet (one of
' earliest authors we are
' aware of is Lewis Moten).
' However, our version contains
' improvements such as the
' RestrictFileExtension option,
' improved support for 
' multi-lingual environments,
' a superior multi-byte string
' to binary conversion function
' (which uses ADO rather than
' VBScript multi-byte functions),
' and a superior method to
' stream data to disk (which
' uses ADO rather than
' FileSystemObject.

Public TotalBytes
Public Files
Public Form
Public SafeFileExtensions
Public RestrictFileExtentions

Private binData

	Private SUB Class_Initialize
		set Files = Server.CreateObject("Scripting.Dictionary")
		set Form = Server.CreateObject("Scripting.Dictionary")
		RestrictFileExtentions = False
		SafeFileExtensions = ""
		TotalBytes = 0
		server.scripttimeout = 600
		on error resume next
		Response.CodePage = 1255
		Response.CharSet = "windows-1255"
		on error GoTo 0
	END SUB

	Private SUB Class_Terminate
		set Files = nothing
		set Form = nothing
		on error resume next
		response.CodePage = 1255
		Response.CharSet = "windows-1255"
		on error GoTo 0
	END SUB

	Private FUNCTION IsSafe(strFileName)

		IF RestrictFileExtentions THEN

			IF Instr(SafeFileExtensions,LCase(Right(strFileName,Len(strFileName) - InStrRev(strFileName,".")))) <> 0 THEN
			IsSafe = True
			ELSE
			IsSafe = False
			END IF

		ELSE
		IsSafe = True
		END IF

	END FUNCTION

	Public Default FUNCTION Upload(ByVal strUploadPath)

	Upload = False

	dim tokenNewLine
	dim tokenTerminate
	dim tokenContentDisposition
	dim tokenName
	dim tokenDoubleQuotes
	dim tokenFileName
	dim tokenContentType
	dim lngBeginningPos
	dim lngEndPos
	dim strSeparator
	dim lngSeparatorPos
	dim lngCurrentPos
	dim strInputName
	dim lngBOF
	dim lngEOF

	TotalBytes = request.totalbytes

	tokenNewLine = ConvertStringToMultiByteString(Chr(13))
	tokenTerminate = ConvertStringToMultiByteString("--")
	tokenContentDisposition = ConvertStringToMultiByteString("Content-Disposition")
	tokenName = ConvertStringToMultiByteString("name=")
	tokenDoubleQuotes = ConvertStringToMultiByteString(Chr(34))
	tokenFileName = ConvertStringToMultiByteString("filename=")
	tokenContentType = ConvertStringToMultiByteString("Content-Type:")

	on error resume next
	binData = request.binaryread(TotalBytes)

		IF err.Number <> 0 THEN
		Upload = False
		EXIT FUNCTION
		END IF

	on error GoTo 0

	lngBeginningPos = 1
	lngEndPos = InstrB(lngBeginningPos,binData,tokenNewLine)

		IF (lngEndPos-lngBeginningPos) <= 0 THEN
		EXIT FUNCTION
		END IF

	'strSeparator is a separator like -----------------------------21763138716045

	strSeparator = MidB(binData,lngBeginningPos,lngEndPos-lngBeginningPos)
	lngSeparatorPos = InstrB(lngBeginningPos,binData,strSeparator)

		DO UNTIL (lngSeparatorPos >= InstrB(binData,strSeparator & tokenTerminate))

		lngCurrentPos = InstrB(lngSeparatorPos,binData,tokenContentDisposition) 'skips the contentdisposition token
		lngCurrentPos = InstrB(lngCurrentPos,binData,tokenName) 'skips the name token
		lngBeginningPos = lngCurrentPos + 6 'forwards to the quotes
		lngEndPos = InstrB(lngBeginningPos,binData,tokenDoubleQuotes) 'finds the end of the quotes (in between quotes is the name of the input/or/file)
		strInputName = LCase(ConvertBinaryToString(MidB(binData,lngBeginningPos,lngEndPos-lngBeginningPos)))

		lngBOF = InstrB(lngSeparatorPos,binData,tokenFileName)
		lngEOF = InstrB(lngEndPos,binData,strSeparator)

			IF ((lngBOF <> 0) AND (lngBOF < lngEOF)) THEN

			dim objFileBlob
			dim strFileName
			dim strContentType
			dim strData

			lngBeginningPos = lngBOF + 10
			lngEndPos =  InstrB(lngBeginningPos,binData,tokenDoubleQuotes)

			strFileName = Replace(ConvertBinaryToString(MidB(binData,lngBeginningPos,lngEndPos-lngBeginningPos)),"/","\") 'UNIX paths might contain "/"
			strFileName = Right(strFileName,len(strFileName)-InstrRev(strFileName,"\"))

			lngCurrentPos = InstrB(lngEndPos,binData,tokenContentType)
			lngBeginningPos = lngCurrentPos + 14
			lngEndPos = InstrB(lngBeginningPos,binData,tokenNewLine)

			strContentType = ConvertBinaryToString(MidB(binData,lngBeginningPos,lngEndPos-lngBeginningPos))

			lngBeginningPos = lngEndPos+4
			lngEndPos = InstrB(lngBeginningPos,binData,strSeparator) - 2

			strData = MidB(binData,lngBeginningPos,lngEndPos-lngBeginningPos)

				IF IsSafe(strFileName) THEN

				set objFileBlob = New clsUploadedFile
				    objFileBlob.UploadPath = strUploadPath
				    objFileBlob.FileName = strFileName
				    objFileBlob.ContentType = strContentType
				    objFileBlob.Data = strData

					IF objFileBlob.Size > 0 THEN
					Files.Add LCase(strInputName),objFileBlob
					END IF

				END IF

			ELSE

			lngCurrentPos = InstrB(lngCurrentPos,binData,tokenNewLine)
			lngBeginningPos = lngCurrentPos + 4
			lngEndPos = InstrB(lngBeginningPos,binData,strSeparator) - 2

				IF NOT Form.Exists(LCase(strInputName)) THEN
				Form.Add LCase(strInputName),ConvertBinaryToString(MidB(binData,lngBeginningPos,lngEndPos-lngBeginningPos))
				END IF

			END IF

		lngSeparatorPos = InstrB(lngSeparatorPos + LenB(lngSeparatorPos),binData,strSeparator)

		LOOP

	Upload = True

	END FUNCTION

	Private FUNCTION ConvertStringToMultiByteString(str)

	dim lngIndex
	dim strNew

		FOR lngIndex = 1 to len(str)
		strNew = strNew & ChrB(AscB(Mid(str,lngIndex,1)))
		NEXT

	ConvertStringToMultiByteString = strNew

	END FUNCTION

	Private FUNCTION ConvertBinaryToString(bin)

	dim lngIndex
	dim strNew
	    strNew = ""

		FOR lngIndex = 1 TO lenB(bin)
		strNew = strNew & Chr(AscB(MidB(bin,lngIndex,1))) 
		NEXT

	ConvertBinaryToString = strNew

	END FUNCTION
End Class

Class clsUploadedFile

Public ContentType
Public Data
Public FileName
Public UploadPath

	Private SUB Class_Initialize
		UploadPath = Server.MapPath("_clsUpload.asp")
	END SUB

	Private SUB Class_Terminate
	END SUB

	Public Property GET Size
	Size = lenB(Data)
	END Property

	Public FUNCTION BuildPath
	dim objFSO
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	BuildPath = objFSO.BuildPath(UploadPath,FileName)
	set objFSO = nothing
	END FUNCTION

	Public Property GET FileExists
	dim objFSO
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	FileExists = objFSO.FileExists(BuildPath)
	set objFSO = nothing
	END Property

	Public Property GET FolderExists
	dim objFSO
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	FolderExists = objFSO.FolderExists(UploadPath)
	set objFSO = nothing
	END Property
	
	Public FUNCTION Save

	Save = False

	dim strPath
	    strPath = BuildPath

	dim objADOStream
	set objADOStream = Server.CreateObject("ADODB.Stream")
	    objADOStream.Type = 1 'adTypeBinary
	    objADOStream.Open
	    objADOStream.Write ConvertMultiByteStringToBinary(Data)
	    objADOStream.SavetoFile strPath, 2 'adSaveCreateOverWrite
	set objADOStream = nothing

	Save = True

	END FUNCTION

	Private FUNCTION ConvertMultiByteStringToBinary(bStr)

	dim binData
	dim lngMByteString
	    lngMByteString = LenB(bStr)

		IF lngMByteString > 0 THEN

		dim objRs
		set objRs = Server.CreateObject("ADODB.Recordset")
		    objRs.Fields.Append "binData",205,lngMByteString 'adLongVarBinary,length
		    objRs.Open
		    objRs.AddNew
		    objRs("binData").AppendChunk bStr & ChrB(0)
		    objRs.Update
		binData = objRs("binData").GetChunk(lngMByteString)
		    objRs.Close
		set objRs = nothing

		END IF

	ConvertMultiByteStringToBinary = binData

	END FUNCTION

END Class%>