<%
Dim strBugSnagAccessToken
Dim strBugSnagPersonUserId
Dim strBugSnagPersonUserName

strBugSnagAccessToken    = ""
strBugSnagPersonUserId   = ""
strBugSnagPersonUserName = ""

Function BugSnagASPError()
    Dim objError
    Set objError = Server.GetLastError()
    BugSnagASPError = BugSnag("error", "", "", objError)
    Set objError = Nothing
End Function

Function BugSnagError(strMessage, strExtraPayload)
    BugSnagError = BugSnag("error", strMessage, strExtraPayload, NULL)
End Function

Function BugSnagWarning(strMessage, strExtraPayload)
    BugSnagWarning = BugSnag("warning", strMessage, strExtraPayload, NULL)
End Function

Function BugSnagInfo(strMessage, strExtraPayload)
    BugSnagInfo = BugSnag("info", strMessage, strExtraPayload, NULL)
End Function

Function BugSnag(strLevel, strMessage, strExtraPayload, objError)
    Dim strPayload, strURL, intResponseCode

    BugSnag = False
    
    If strBugSnagAccessToken = "" Then
        Exit Function
    End If

    On Error Resume Next
    strPayload = GetBugSnagPayload(strLevel, strMessage, strExtraPayload, objError, True)
    Call PostToBugSnag("https://notify.bugsnag.com/", 1, strPayload, intResponseCode)
    If intResponseCode = 200 Then
        BugSnag = True
    End If
    If NOT BugSnag Then
        strPayload = GetBugSnagPayload(strLevel, strMessage, strExtraPayload, objError, False)
        Call PostToUrl("https://notify.bugsnag.com/", 1, strPayload, "", "", intResponseCode)
        If intResponseCode = 200 Then
            BugSnag = True
        End If
    End If
    On Error Goto 0

    If strLevel = "error" Then
        Set objError = Nothing
    End If
End Function

Function GetBugSnagPayload(strLevel, strMessage, strExtraPayload, objError, blnIncludeSession)
    Dim strPayload

    On Error Resume Next
    If strLevel = "error" Then
        If IsObject(objError) Then
            strMessage = objError.Description
        End If
    End If

    If Request.ServerVariables("HTTPS") = "ON" Then
        strURL = "https://"
    Else
        strURL = "http://"
    End If
    strURL = strURL & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
    If Request.QueryString <> "" Then
        strURL = strURL & Request.QueryString
    End If

    strPayload = "{"
    strPayload = strPayload & """apiKey"": """&strBugSnagAccessToken&""","
    strPayload = strPayload & """notifier"": "
    strPayload = strPayload & "{"
    strPayload = strPayload & "     ""name"": ""Bugsnag ASP"","
    strPayload = strPayload & "     ""version"": ""1.0.0"","
    strPayload = strPayload & "     ""url"": ""https://github.com/hellokoan/bugsnag-classic-asp-notifier"""
    strPayload = strPayload & "},"
    strPayload = strPayload & """events"": "
    strPayload = strPayload & "[{"
    strPayload = strPayload & "     ""payloadVersion"": 2,"
    strPayload = strPayload & "     ""severity"": """&strLevel&""", "
    strPayload = strPayload & "     ""exceptions"": "
    strPayload = strPayload & "     [{"
    If strLevel = "error" AND IsObject(objError) Then
    strPayload = strPayload & "         ""errorClass"": """&PrepareForBugSnag(objError.Category)&""","
    strPayload = strPayload & "         ""message"": """&PrepareForBugSnag(objError.Description)&""","
    strPayload = strPayload & "         ""stacktrace"": "
    strPayload = strPayload & "         [{"
    strPayload = strPayload & "             ""inProject"": true,"
    strPayload = strPayload & "             ""file"": """&PrepareForBugSnag(objError.File)&""","
    strPayload = strPayload & "             ""lineNumber"": """&PrepareForBugSnag(objError.Line)&""","
    strPayload = strPayload & "             ""columnNumber"": """&PrepareForBugSnag(objError.Column)&""""
    If objError.Source <> "" Then
    strPayload = strPayload & "             ,""code"": {"""&PrepareForBugSnag(objError.Line)&""": """&PrepareForBugSnag(objError.Source)&"""}"
    End If
    strPayload = strPayload & "         }]"
    Else
    strPayload = strPayload & "         ""errorClass"": """&PrepareForBugSnag(strMessage)&""","
    strPayload = strPayload & "         ""stacktrace"": [{}]"
    End If
    strPayload = strPayload & "     }]"

    If strBugSnagPersonUserId <> "" OR strBugSnagPersonUserName <> "" Then
    strPayload = strPayload & "     ,""user"": "
    strPayload = strPayload & "     {"
    strPayload = strPayload & "        ""id"": """&PrepareForBugSnag(strBugSnagPersonUserId)&""","
    strPayload = strPayload & "        ""email"": """&PrepareForBugSnag(strBugSnagPersonUserName)&""""
    strPayload = strPayload & "     }"
    End If
    strPayload = strPayload & "     ,""metaData"": "
    strPayload = strPayload & "     {"
    strPayload = strPayload & "         ""method"": """&PrepareForBugSnag(Request.ServerVariables("HTTP_METHOD"))&""","
    strPayload = strPayload & "         ""url"": """&PrepareForBugSnag(strUrl)&""","
    strPayload = strPayload & "         ""query_string"": """&PrepareForBugSnag(Request.QueryString)&""","
    strPayload = strPayload & "         ""form"": """&PrepareForBugSnag(Replace(Request.Form, "&", VbCrLf))&""","
    strPayload = strPayload & "         ""ip"": """&PrepareForBugSnag(Request.ServerVariables("REMOTE_ADDR"))&""""
    If blnIncludeSession Then
    strPayload = strPayload & "         ,""session"": "&GetSessionAsString()
    End If
    If strExtraPayload <> "" Then
        strPayload = strPayload & ","
        strPayload = strPayload & strExtraPayload
    End If
    strPayload = strPayload & "     }"
    strPayload = strPayload & "}]"
    strPayload = strPayload & "}"
    On Error Goto 0

    GetBugSnagPayload = strPayload
End Function

Function PrepareForBugSnag(strData)
    strData = EnsureIsTrimmedString(strData)
    strData = Replace(strData, "\", "\\")
    strData = Replace(strData, """", "\""")
    strData = Replace(strData, VbCrLf, "\n")
    PrepareForBugSnag = strData
End Function

Function EnsureIsTrimmedString(ByVal strString)
    strString = strString & ""
    If NOT IsNull(strString) Then
        strString = Trim(strString)
    End If
    EnsureIsTrimmedString = strString 
End Function

Function GetSessionAsString()
    On Error Resume Next
    Dim sessionItem, strSession
    strSession = "{"
    For Each sessionItem in Session.Contents
        If IsArray(Session(sessionItem)) Then
            strSession = strSession & """" & PrepareForBugSnag(sessionItem) & """" & ":" & """" & PrepareForBugSnag(BugSnagPrintArray(Session(sessionItem))) & ""","
        ElseIf Left(Session(sessionItem), 1) = "{" Then
            strSession = strSession & """" & PrepareForBugSnag(sessionItem) & """" & ":" & Session(sessionItem) & ","
        Else
            strSession = strSession & """" & PrepareForBugSnag(sessionItem) & """" & ":" & """" & PrepareForBugSnag(Session(sessionItem)) & ""","
        End If
    Next
    If Len(strSession) > 1 Then
        ' Remove the trailing comma
        strSession = Left(strSession, Len(strSession)-1)
    End If
    strSession = strSession & "}"
    GetSessionAsString = strSession
End Function

Function BugSnagPrintArray(aryArray)
	On Error Resume Next
	Dim i, j, k, strOut, strElement, aryDimensions(10), strDimensions
	i=0
	strDimensions = ""
	For Each strElement in aryArray
		i = i + 1
	Next
	j = 0
	k=1
	Do While j >= 0 AND k < 10
		j = UBound(aryArray, k)
		If j > 0 Then
			strDimensions = strDimensions & j & ","
			aryDimensions(k-1) = j
		End If
		j = 0
		k = k + 1
	Loop
	If strDimensions <> "" Then
		strDimensions = Left(strDimensions, Len(strDimensions)-1)
	End If
	strOut = "Array ("&strDimensions&"): " & VbCrLf
	strOut = strOut & "----------" & VbCrLf
	j=0
	For Each strElement in aryArray
		strOut = strOut & strElement
		j = j + 1
		If j = aryDimensions(0)+1 Then
			strOut = strOut & VbCrLf
			j=0
		Else
			strOut = strOut & ","
		End If
	Next	
    strOut = strOut & "----------"
	BugSnagPrintArray = strOut
End Function

Function PostToUrl(strUrl, lTotal, strData, strUserName, strPassword, intResponseCode)
    Dim strAuthorizationHeader
    strAuthorizationHeader = "Basic " & base64_encode(strUserName & ":" & strPassword) & "=="
    PostToUrl = GetURL(strUrl, lTotal, "POST", strData, "application/json", "", "", strAuthorizationHeader, intResponseCode)
End Function

Function GetURL(strUrl, lTotal, strMethod, strData, strContentType, strUserName, strPassword, strAuthorizationHeader, intResponseCode)
    Dim objHttp, GotResponse, intSecondsWait
    On Error Resume Next
    GetURL = ""
        
    Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    If lTotal = 0 Then
        objHttp.open strMethod, StripHTML(Replace(strUrl, "&amp;", "&")), True, strUserName, strPassword
    Else
        objHttp.setTimeouts lTotal*1000/4, lTotal*1000/4, lTotal*1000/4, lTotal*1000
        objHttp.open strMethod, StripHTML(Replace(strUrl, "&amp;", "&")), False, strUserName, strPassword
    End If
    If strContentType <> "" Then
        objHttp.setRequestHeader "Content-Type", strContentType
    End If
    If strAuthorizationHeader <> "" Then
        objHttp.setRequestHeader "Authorization", strAuthorizationHeader
    End If
    objHttp.send strData
    If lTotal = 0 Then
        Set objHttp = Nothing
        Exit Function
    End If
    
    intSecondsWait  = 0
    GotResponse     = False
    Do While objHttp.readyState <> 4
        If Err.Number <> 0 Then
            Exit Do
        End If
        
        objHttp.waitForResponse 1
        intSecondsWait = intSecondsWait + 1
        
        If objHttp.readyState = 4 Then
            GotResponse = True
            Exit Do
        End If
        If intSecondsWait > lTotal Then
            GotResponse = False
            Exit Do
        End If
    Loop
    If objHttp.readyState = 4 Then
        GotResponse = True
    End If
    
    If GotResponse AND Err.Number = 0 Then
        intResponseCode = objHttp.status
        If objHttp.status >= 200 AND objHttp.status <= 299 Then
            If InStr(strURL, ".jpg") > 0 OR InStr(strURL, ".gif") > 0 OR InStr(strURL, ".png") > 0 OR InStr(strURL, ".jpeg") > 0 Then
                GetURL = objHTTP.ResponseBody
            Else
                GetURL = strOutDB(objHTTP.ResponseText)
            End If
        Else
            GotResponse = False
        End If
        
    ElseIf Err.Number <> 0 Then
        Err.Clear
    End If
    
    Set objHttp = Nothing
    
    On Error Goto 0
End Function
%>
