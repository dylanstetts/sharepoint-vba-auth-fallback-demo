Attribute VB_Name = "SharePointSoapRepro"
Option Explicit

#If Mac Then
    ' MSXML ActiveX is unavailable in Excel for Mac.
#ElseIf VBA7 Then
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef value As SYSTEMTIME)
#Else
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef value As SYSTEMTIME)
#End If

Private Type SYSTEMTIME
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type

Private Const RESULT_SHEET As String = "SharePoint Repro"
Private Const SOAP_NAMESPACE As String = _
    "http://schemas.microsoft.com/sharepoint/soap/"
Private lastRequestError As String
Private sessionAccessToken As String

Public Sub RunSharePointReadOnlyRepro()
    Dim results As Worksheet
    Dim soapBody As String
    Dim actionUri As String
    Dim restStatus As Long

    On Error GoTo ReproFailed

    ValidateRuntime
    ValidateConfiguration
    Set results = PrepareResultSheet()
    soapBody = BuildGetListItemsBody(SharePointAuthListName())
    actionUri = SOAP_NAMESPACE & "GetListItems"

    sessionAccessToken = vbNullString
    restStatus = ExecuteRequest(results, "REST GET control - ambient", "GET", _
        SharePointAuthRestProbeUrl(), "text/xml; charset=UTF-8", vbNullString, _
        vbNullString, "Ambient", False)

    If restStatus <> 200 Then
        sessionAccessToken = GetSharePointAccessToken()
        restStatus = ExecuteRequest(results, "REST GET control - OAuth", "GET", _
            SharePointAuthRestProbeUrl(), "text/xml; charset=UTF-8", _
            vbNullString, vbNullString, "OAuth fallback", True)
    End If

    If restStatus <> 200 Then
        results.Columns.AutoFit
        results.Activate
        If restStatus = 0 Then
            MsgBox "The REST control failed before an HTTP response was " & _
                "received. SOAP comparison was skipped." & vbCrLf & vbCrLf & _
                lastRequestError, vbCritical, "SharePoint request failed"
        Else
            MsgBox "The REST authentication control returned HTTP " & _
                CStr(restStatus) & ". SOAP comparison was skipped." & _
                vbCrLf & vbCrLf & _
                "Review the response and VBA error columns on the " & _
                RESULT_SHEET & " worksheet.", vbExclamation, _
                "SharePoint authentication required"
        End If
        Exit Sub
    End If

    Call ExecuteRequest(results, "SOAP GetListItems - unquoted", "POST", _
        SharePointAuthSoapUrl(), "text/xml; charset=UTF-8", actionUri, _
        soapBody, "Unquoted", Len(sessionAccessToken) > 0)

    Call ExecuteRequest(results, "SOAP GetListItems - quoted", "POST", _
        SharePointAuthSoapUrl(), "text/xml; charset=UTF-8", _
        Chr$(34) & actionUri & Chr$(34), soapBody, "Quoted", _
        Len(sessionAccessToken) > 0)

    results.Columns.AutoFit
    results.Activate
    Exit Sub

ReproFailed:
    If Not results Is Nothing Then
        results.Columns.AutoFit
        results.Activate
    End If
    MsgBox "The authentication fallback could not complete." & _
        vbCrLf & vbCrLf & "VBA error: " & CStr(Err.Number) & _
        vbCrLf & "Source: " & Err.Source & _
        vbCrLf & "Description: " & Err.Description, _
        vbCritical, "SharePoint OAuth fallback failed"
End Sub

Private Function ExecuteRequest(ByVal results As Worksheet, _
                                ByVal testName As String, _
                                ByVal method As String, _
                                ByVal requestUrl As String, _
                                ByVal contentType As String, _
                                ByVal soapAction As String, _
                                ByVal requestBody As String, _
                                ByVal actionVariant As String, _
                                ByVal useOAuth As Boolean) As Long
    Dim http As Object
    Dim startedAt As Double
    Dim durationMilliseconds As Long
    Dim startUtc As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String
    Dim failureStage As String

    On Error GoTo RequestFailed

    lastRequestError = vbNullString
    startUtc = UtcNowText()
    startedAt = Timer

    If useOAuth Then
        failureStage = "validating the SharePoint access token"
        If Len(sessionAccessToken) = 0 Then
            Err.Raise vbObjectError + 2605, "ExecuteRequest", _
                "OAuth was requested but no access token is available."
        End If
    End If

    failureStage = "creating MSXML2.XMLHTTP.6.0"
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    failureStage = "opening the " & method & " request"
    http.Open method, requestUrl, False

    failureStage = "setting request headers"
    If useOAuth Then
        http.setRequestHeader "Authorization", _
            "Bearer " & sessionAccessToken
    End If
    http.setRequestHeader "Content-Type", contentType
    If Len(soapAction) > 0 Then
        http.setRequestHeader "SOAPAction", soapAction
    End If

    failureStage = "sending the request"
    http.Send requestBody

    failureStage = "reading the HTTP response"

    durationMilliseconds = ElapsedMilliseconds(startedAt)
    AppendResult results, testName, startUtc, method, requestUrl, _
        actionVariant, http.Status, http.statusText, durationMilliseconds, _
        SafeResponseHeader(http, "SPRequestGuid"), _
        SafeResponseHeader(http, "request-id"), _
        SafeResponseHeader(http, "Date"), _
        SafeResponseHeader(http, "x-msdavext_error"), _
        SafeResponseHeader(http, "x-forms_based_auth_required"), _
        SafeResponseHeader(http, "Content-Length"), _
        Left$(CStr(http.responseText), 2000), vbNullString
    ExecuteRequest = CLng(http.Status)
    Exit Function

RequestFailed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    lastRequestError = "Stage: " & failureStage & vbCrLf & _
        "VBA error: " & CStr(failureNumber) & vbCrLf & _
        "Source: " & failureSource & vbCrLf & _
        "Description: " & failureDescription
    If startedAt > 0 Then
        durationMilliseconds = ElapsedMilliseconds(startedAt)
    End If
    AppendResult results, testName, startUtc, method, requestUrl, _
        actionVariant, vbNullString, vbNullString, durationMilliseconds, _
        vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, _
        vbNullString, vbNullString, Replace(lastRequestError, vbCrLf, " | ")
    ExecuteRequest = 0
End Function

Private Function PrepareResultSheet() As Worksheet
    Dim results As Worksheet
    Dim headers As Variant
    Dim columnIndex As Long

    On Error Resume Next
    Set results = ThisWorkbook.Worksheets(RESULT_SHEET)
    On Error GoTo 0

    If results Is Nothing Then
        Set results = ThisWorkbook.Worksheets.Add
        results.Name = RESULT_SHEET
    Else
        results.Cells.Clear
    End If

    headers = Array("Test", "Start UTC", "Method", "URL", _
        "SOAPAction variant", "HTTP status", "Status text", "Duration ms", _
        "SPRequestGuid", "request-id", "Server date", "x-msdavext_error", _
        "x-forms_based_auth_required", "Content-Length", _
        "Response excerpt", "VBA error")

    For columnIndex = LBound(headers) To UBound(headers)
        results.Cells(1, columnIndex + 1).Value = headers(columnIndex)
    Next columnIndex
    results.Rows(1).Font.Bold = True

    Set PrepareResultSheet = results
End Function

Private Sub AppendResult(ByVal results As Worksheet, _
                         ByVal testName As String, _
                         ByVal startUtc As String, _
                         ByVal method As String, _
                         ByVal requestUrl As String, _
                         ByVal actionVariant As String, _
                         ByVal httpStatus As Variant, _
                         ByVal statusText As String, _
                         ByVal durationMilliseconds As Long, _
                         ByVal spRequestGuid As String, _
                         ByVal requestId As String, _
                         ByVal serverDate As String, _
                         ByVal msDavError As String, _
                         ByVal formsAuthRequired As String, _
                         ByVal contentLength As String, _
                         ByVal responseExcerpt As String, _
                         ByVal vbaError As String)
    Dim rowNumber As Long
    Dim values As Variant
    Dim columnIndex As Long

    rowNumber = results.Cells(results.Rows.Count, 1).End(xlUp).Row + 1
    values = Array(testName, startUtc, method, SanitizeUrl(requestUrl), _
        actionVariant, httpStatus, statusText, durationMilliseconds, _
        spRequestGuid, requestId, serverDate, msDavError, formsAuthRequired, _
        contentLength, responseExcerpt, vbaError)

    For columnIndex = LBound(values) To UBound(values)
        results.Cells(rowNumber, columnIndex + 1).Value = values(columnIndex)
    Next columnIndex
End Sub

Private Function BuildGetListItemsBody(ByVal listName As String) As String
    BuildGetListItemsBody = _
        "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " & _
        "xmlns:xsd='http://www.w3.org/2001/XMLSchema' " & _
        "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
        "<soap:Body><GetListItems xmlns='" & SOAP_NAMESPACE & "'>" & _
        "<listName>" & XmlEncode(listName) & "</listName>" & _
        "<rowLimit>1</rowLimit></GetListItems></soap:Body></soap:Envelope>"
End Function

Private Function XmlEncode(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, Chr$(34), "&quot;")
    XmlEncode = Replace(value, "'", "&apos;")
End Function

Private Function SafeResponseHeader(ByVal http As Object, _
                                    ByVal headerName As String) As String
    On Error Resume Next
    SafeResponseHeader = CStr(http.getResponseHeader(headerName))
    On Error GoTo 0
End Function

Private Function SanitizeUrl(ByVal requestUrl As String) As String
    Dim queryPosition As Long

    queryPosition = InStr(1, requestUrl, "?", vbBinaryCompare)
    If queryPosition > 0 Then
        SanitizeUrl = Left$(requestUrl, queryPosition - 1)
    Else
        SanitizeUrl = requestUrl
    End If
End Function

Private Function ElapsedMilliseconds(ByVal startedAt As Double) As Long
    Dim elapsedSeconds As Double

    elapsedSeconds = Timer - startedAt
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
    ElapsedMilliseconds = CLng(elapsedSeconds * 1000#)
End Function

Private Function UtcNowText() As String
#If Mac Then
    Err.Raise vbObjectError + 2603, "UtcNowText", _
        "This reproducer requires Windows desktop Excel."
#Else
    Dim value As SYSTEMTIME

    GetSystemTime value
    UtcNowText = Format$(value.Year, "0000") & "-" & _
        Format$(value.Month, "00") & "-" & Format$(value.Day, "00") & "T" & _
        Format$(value.Hour, "00") & ":" & Format$(value.Minute, "00") & ":" & _
        Format$(value.Second, "00") & "." & _
        Format$(value.Milliseconds, "000") & "Z"
#End If
End Function

Private Sub ValidateRuntime()
#If Mac Then
    Err.Raise vbObjectError + 2603, "RunSharePointReadOnlyRepro", _
        "This reproducer requires Windows desktop Excel because the customer " & _
        "path uses MSXML2.XMLHTTP.6.0. ActiveX/COM is unavailable in Excel for Mac."
#Else
    Dim probe As Object
    Dim createError As String

    On Error Resume Next
    Set probe = CreateObject("MSXML2.XMLHTTP.6.0")
    createError = CStr(Err.Number) & ": " & Err.Description
    On Error GoTo 0

    If probe Is Nothing Then
        Err.Raise vbObjectError + 2604, "RunSharePointReadOnlyRepro", _
            "MSXML2.XMLHTTP.6.0 could not be created on this Windows host (" & _
            createError & "). Confirm MSXML 6 registration and Office policy."
    End If

    Set probe = Nothing
#End If
End Sub

Private Sub ValidateConfiguration()
    If SharePointAuthListName() Like "REPLACE-*" Then
        Err.Raise vbObjectError + 2600, "RunSharePointReadOnlyRepro", _
            "Configure SharePointAuthConfig.bas before running the demo."
    End If

    If LCase$(Left$(SharePointAuthRestProbeUrl(), 8)) <> "https://" Or _
       LCase$(Left$(SharePointAuthSoapUrl(), 8)) <> "https://" Then
        Err.Raise vbObjectError + 2601, "RunSharePointReadOnlyRepro", _
            "Only HTTPS endpoints are allowed."
    End If

    If InStr(1, SharePointAuthRestProbeUrl(), ".sharepoint.com/", _
             vbTextCompare) = 0 Or _
       InStr(1, SharePointAuthSoapUrl(), ".sharepoint.com/", _
             vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 2602, "RunSharePointReadOnlyRepro", _
            "Both endpoints must be SharePoint Online URLs."
    End If
End Sub
