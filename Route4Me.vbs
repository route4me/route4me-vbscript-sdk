Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
CONST WHR_EnableRedirects = 6

Class Route4Me

	Private m_fileName

	Public Property Get OutputFile

		OutputFile = m_fileName

	End Property

	Public Property Let OutputFile(ByVal value)

		m_fileName = value

	End Property

	Public Sub Write2File(result)
		Dim fileName
		Dim spFile
		If m_fileName = Empty Then
			'MsgBox("File not defined")
			fileName="file1.txt"
		Else
			'MsgBox("File Defined")
			fileName=OutputFile
		End If
		

		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(fileName) Then
			Set spFile = fso.OpenTextFile(fileName,2,True)
		Else
			Set spFile= fso.CreateTextFile(fileName,True)
		End If

		spFile.WriteLine(result)
		spFile.Close
		Set fso = nothing
	End Sub
	
	Public Function File2Json(jFile)
		Dim spFile
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(jFile) Then
			Set spFile = fso.OpenTextFile(jFile,1,True)
			File2Json = spFile.ReadAll()
			File2Json=Trim(File2Json)
		Else
			WScript.Echo "File " & fileName &" doesn't exists..."
			File2Json = ""
		End If
	End Function
	
	Public Sub HttpGetRequest(url)
		Set WshShell = WScript.CreateObject("WScript.Shell")
		'Set http = CreateObject("Microsoft.XmlHttp")
		Set http = CreateObject("MSXML2.ServerXMLHTTP")
		
		On Error Resume Next

		http.open "GET",url,False
		http.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
		http.send ""
		
		If Err.Number = 0 Then
			Write2File(http.responseText)
		Else
			WScript.Echo "error " & Err.Number& ":" & Err.Description
		End If
		
		Set WshShell = Nothing
		Set http = Nothing
	End Sub
	
	Public Sub HttpGetRequest2(url,jFile)
		Set WshShell = WScript.CreateObject("WScript.Shell")
		'Set http = CreateObject("Microsoft.XmlHttp")
		Set http = CreateObject("MSXML2.ServerXMLHTTP")
		
		On Error Resume Next

		http.open "GET",url,False
		http.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
		
		On Error Resume Next
		If jFile="" Then
			http.send ""
		Else
			jText = File2Json(jFile)
			http.setRequestHeader "Content-Length", Len(jText)
			http.send jText
		End If
		
		If Err.Number = 0 Then
			Write2File(http.responseText)
		Else
			WScript.Echo "error " & Err.Number& ":" & Err.Description
		End If
		
		Set WshShell = Nothing
		Set http = Nothing
	End Sub
	
	Public Sub HttpPostRequest(url,jFile)
		Dim jText
		Set WshShell = WScript.CreateObject("WScript.Shell")
		'Set http = CreateObject("Microsoft.XmlHttp")
		Set http = CreateObject("MSXML2.ServerXMLHTTP")
		http.open "POST", url, False
		http.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
		http.setRequestHeader "Content-Type", "application/json"
		
		On Error Resume Next
		
		jText = File2Json(jFile)
		http.setRequestHeader "Content-Length", Len(jText)
		http.send jText

		'http.waitForResponse(20)
		If http.Status >= 400 And http.Status <= 599 Then
        	WScript.Echo "Error Occurred : " & http.status & " - " & http.statusText
        End If
		
		If Err.Number = 0 Then
			Write2File(http.responseText)
		Else
			WScript.Echo "error " & Err.Number& ":" & Err.Description
		End If
		
		Set WshShell = Nothing
		Set http = Nothing
	End Sub
	
	Public Sub HttpPostRequest2(url,jFile)
		Dim jText
		Set WshShell = WScript.CreateObject("WScript.Shell")
		'Set http = CreateObject("MSXML2.ServerXMLHTTP")
		Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
		http.open "POST", url, False
		http.Option(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
		http.Option(WHR_EnableRedirects) = False
		
		http.setRequestHeader "Content-Type", "application/json"
		
		On Error Resume Next
		
		jText = File2Json(jFile)
		http.setRequestHeader "Content-Length", Len(jText)

		http.send jText

		If http.Status >= 400 And http.Status <= 599 Then
        	WScript.Echo "Error Occurred : " & http.status & " - " & http.statusText
        End If
		
		If Err.Number = 0 Then
			Write2File(http.responseText)
		Else
			WScript.Echo "error " & Err.Number& ":" & Err.Description
		End If
		
		Set WshShell = Nothing
		Set http = Nothing
	End Sub
	
	Sub Upload(strUploadUrl, strFilePath, strFileField, strDataPairs)
	'Uses POST to upload a file and miscellaneous form data
	'strUploadUrl is the URL (https://www.route4me.com/actions/upload/upload.php)
	'strFilePath is the file to upload
	'strFileField is the web page equivalent form field name for the file (strFilename)
	'strDataPairs are pipe-delimited form data pairs (foo=bar|strFilename=filename)
	Const MULTIPART_BOUNDARY = "---------------------------0123456789012"
	Dim ado, rs
	Dim lngCount
	Dim bytFormData, bytFormStart, bytFormEnd, bytFile
	Dim strFormStart, strFormEnd, strDataPair
	Dim web
	Const adLongVarBinary = 205
		'Read the file into a byte array
		Set ado = CreateObject("ADODB.Stream")
		ado.Type = 1
		ado.Open
		ado.LoadFromFile strFilePath
		bytFile = ado.Read
		ado.Close
		'Create the multipart form data. 
		'Define the end of form
		strFormEnd = vbCrLf & "--" & MULTIPART_BOUNDARY & "--" & vbCrLf
		'First add any ordinary form data pairs
		strFormStart = ""
		For Each strDataPair In Split(strDataPairs, "|")
			strFormStart = strFormStart & "--" & MULTIPART_BOUNDARY & vbCrLf
			strFormStart = strFormStart & "Content-Disposition: form-data; "
			strFormStart = strFormStart & "name=""" & Split(strDataPair, "=")(0) & """"
			strFormStart = strFormStart & vbCrLf & vbCrLf
			strFormStart = strFormStart & Split(strDataPair, "=")(1)
			strFormStart = strFormStart & vbCrLf
		Next
		'Now add the header for the uploaded file
		strFormStart = strFormStart & "--" & MULTIPART_BOUNDARY & vbCrLf
		strFormStart = strFormStart & "Content-Disposition: form-data; "
		strFormStart = strFormStart & "name=""" & strFileField & """; "
		strFormStart = strFormStart & "filename=""" & Mid(strFilePath, InStrRev(strFilePath, "\") + 1) & """"
		strFormStart = strFormStart & vbCrLf
		strFormStart = strFormStart & "Content-Type: application/upload" 'bogus, but it works
		strFormStart = strFormStart & vbCrLf & vbCrLf

		'Create a recordset large enough to hold everything
		Set rs = CreateObject("ADODB.Recordset")
		rs.Fields.Append "FormData", adLongVarBinary, Len(strFormStart) + LenB(bytFile) + Len(strFormEnd)
		rs.Open
		rs.AddNew
		'Convert form data so far to zero-terminated byte array
		For lngCount = 1 To Len(strFormStart)
			bytFormStart = bytFormStart & ChrB(Asc(Mid(strFormStart, lngCount, 1)))
		Next
		rs("FormData").AppendChunk bytFormStart & ChrB(0)
		bytFormStart = rs("formData").GetChunk(Len(strFormStart))
		rs("FormData") = ""
		'Get the end boundary as a zero-terminated byte array
		For lngCount = 1 To Len(strFormEnd)
			bytFormEnd = bytFormEnd & ChrB(Asc(Mid(strFormEnd, lngCount, 1)))
		Next
		
		rs("FormData").AppendChunk bytFormEnd & ChrB(0)
		bytFormEnd = rs("formData").GetChunk(Len(strFormEnd))
		rs("FormData") = ""
		'Now merge it all
		rs("FormData").AppendChunk bytFormStart
		rs("FormData").AppendChunk bytFile
		rs("FormData").AppendChunk bytFormEnd
		bytFormData = rs("FormData")
		rs.Close
		'Upload it
		Set web = CreateObject("WinHttp.WinHttpRequest.5.1")
		web.Open "POST", strUploadUrl, False
		web.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & MULTIPART_BOUNDARY
		web.Send bytFormData
		
		
		If Err.Number = 0 Then
			Write2File(web.responseText)
		Else
			WScript.Echo "error " & Err.Number& ":" & Err.Description
		End If
	End Sub
	
	Public Sub HttpPutRequest(url,jFile)
		Dim jText
		Set WshShell = WScript.CreateObject("WScript.Shell")
		'Set http = CreateObject("Microsoft.XmlHttp")
		Set http = CreateObject("MSXML2.ServerXMLHTTP")
		http.open "PUT", url, False
		http.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
		http.setRequestHeader "Content-Type", "application/json"
		
		On Error Resume Next
		
		jText = File2Json(jFile)
		http.setRequestHeader "Content-Length", Len(jText)
		http.send jText

		If Err.Number = 0 Then
			Write2File(http.responseText)
		Else
			WScript.Echo "error " & Err.Number& ":" & Err.Description
		End If
		
		Set WshShell = Nothing
		Set http = Nothing
	End Sub
	
	Public Sub HttpDeleteRequest(url,jFile)
		Dim jText
		Set WshShell = WScript.CreateObject("WScript.Shell")
		'Set http = CreateObject("Microsoft.XmlHttp")
		Set http = CreateObject("MSXML2.ServerXMLHTTP")
		http.open "DELETE", url, False
		http.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
		http.setRequestHeader "Content-Type", "application/json"
		
		On Error Resume Next
		If jFile="" Then
			http.send ""
		Else
			jText = File2Json(jFile)
			http.setRequestHeader "Content-Length", Len(jText)
			http.send jText
		End If
		
		If Err.Number = 0 Then
			Write2File(http.responseText)
		Else
			WScript.Echo "error " & Err.Number& ":" & Err.Description
		End If
		
		Set WshShell = Nothing
		Set http = Nothing
	End Sub
End Class