Attribute VB_Name = "SharePointDeviceAuth"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If

Private Const FORM_CONTENT_TYPE As String = _
    "application/x-www-form-urlencoded"

Private cachedAccessToken As String
Private cachedRefreshToken As String
Private accessTokenExpiresAt As Date

Public Function GetSharePointAccessToken() As String
    ValidateAuthConfiguration

    If Len(cachedAccessToken) > 0 Then
        If DateDiff("s", Now, accessTokenExpiresAt) > 300 Then
            GetSharePointAccessToken = cachedAccessToken
            Exit Function
        End If
    End If

    If Len(cachedRefreshToken) > 0 Then
        If TryRefreshSharePointToken() Then
            GetSharePointAccessToken = cachedAccessToken
            Exit Function
        End If
    End If

    AcquireSharePointTokenByDeviceCode
    GetSharePointAccessToken = cachedAccessToken
End Function

Public Sub ResetSharePointAuthentication()
    cachedAccessToken = vbNullString
    cachedRefreshToken = vbNullString
    accessTokenExpiresAt = 0
End Sub

Private Sub AcquireSharePointTokenByDeviceCode()
    Dim deviceResponse As Object
    Dim tokenResponse As Object
    Dim deviceCode As String
    Dim userCode As String
    Dim verificationUri As String
    Dim signInMessage As String
    Dim tokenError As String
    Dim tokenBody As String
    Dim pollInterval As Long
    Dim expiresIn As Long
    Dim deadline As Date

    Set deviceResponse = AuthHttpPostForm(AuthorityUrl("devicecode"), _
        "client_id=" & FormUrlEncode(SharePointAuthClientId()) & _
        "&scope=" & FormUrlEncode(TokenScope()))

    If deviceResponse.Status <> 200 Then
        RaiseOAuthError "Device authorization failed", deviceResponse
    End If

    deviceCode = JsonString(deviceResponse.responseText, "device_code")
    userCode = JsonString(deviceResponse.responseText, "user_code")
    verificationUri = JsonString(deviceResponse.responseText, _
        "verification_uri")
    signInMessage = JsonString(deviceResponse.responseText, "message")
    pollInterval = JsonLong(deviceResponse.responseText, "interval", 5)
    expiresIn = JsonLong(deviceResponse.responseText, "expires_in", 900)
    deadline = DateAdd("s", expiresIn, Now)

    If Len(deviceCode) = 0 Or Len(userCode) = 0 Or _
       Len(verificationUri) = 0 Then
        Err.Raise vbObjectError + 2700, _
            "AcquireSharePointTokenByDeviceCode", _
            "The device authorization response was incomplete."
    End If

    If Len(signInMessage) = 0 Then
        signInMessage = "Open " & verificationUri & _
            " and enter code " & userCode & "."
    End If

    MsgBox signInMessage, vbInformation, "Sign in to SharePoint"
    ThisWorkbook.FollowHyperlink Address:=verificationUri, NewWindow:=True

    tokenBody = "grant_type=" & _
        FormUrlEncode("urn:ietf:params:oauth:grant-type:device_code") & _
        "&client_id=" & FormUrlEncode(SharePointAuthClientId()) & _
        "&device_code=" & FormUrlEncode(deviceCode)

    Do While Now < deadline
        WaitSeconds pollInterval
        Set tokenResponse = AuthHttpPostForm(AuthorityUrl("token"), tokenBody)

        If tokenResponse.Status = 200 Then
            CacheTokenResponse tokenResponse.responseText
            Exit Sub
        End If

        tokenError = JsonString(tokenResponse.responseText, "error")
        Select Case tokenError
            Case "authorization_pending"
            Case "slow_down"
                pollInterval = pollInterval + 5
            Case "authorization_declined"
                Err.Raise vbObjectError + 2701, _
                    "AcquireSharePointTokenByDeviceCode", _
                    "The user declined the SharePoint sign-in request."
            Case "expired_token", "bad_verification_code"
                Err.Raise vbObjectError + 2702, _
                    "AcquireSharePointTokenByDeviceCode", _
                    "The device sign-in code expired. Start the request again."
            Case Else
                RaiseOAuthError "SharePoint token acquisition failed", _
                    tokenResponse
        End Select
    Loop

    Err.Raise vbObjectError + 2703, _
        "AcquireSharePointTokenByDeviceCode", _
        "The device sign-in request expired before authentication completed."
End Sub

Private Function TryRefreshSharePointToken() As Boolean
    Dim response As Object
    Dim body As String

    body = "client_id=" & FormUrlEncode(SharePointAuthClientId()) & _
        "&grant_type=refresh_token" & _
        "&refresh_token=" & FormUrlEncode(cachedRefreshToken) & _
        "&scope=" & FormUrlEncode(TokenScope())

    Set response = AuthHttpPostForm(AuthorityUrl("token"), body)
    If response.Status = 200 Then
        CacheTokenResponse response.responseText
        TryRefreshSharePointToken = True
    Else
        ResetSharePointAuthentication
    End If
End Function

Private Sub CacheTokenResponse(ByVal json As String)
    Dim expiresIn As Long
    Dim newRefreshToken As String

    cachedAccessToken = JsonString(json, "access_token")
    newRefreshToken = JsonString(json, "refresh_token")
    expiresIn = JsonLong(json, "expires_in", 3600)

    If Len(cachedAccessToken) = 0 Then
        Err.Raise vbObjectError + 2710, "CacheTokenResponse", _
            "The token response did not contain an access token."
    End If

    If Len(newRefreshToken) > 0 Then cachedRefreshToken = newRefreshToken
    accessTokenExpiresAt = DateAdd("s", expiresIn, Now)
End Sub

Private Function AuthHttpPostForm(ByVal url As String, _
                                  ByVal body As String) As Object
    Dim http As Object

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", FORM_CONTENT_TYPE
    http.setRequestHeader "Accept", "application/json"
    http.Send body
    Set AuthHttpPostForm = http
End Function

Private Function AuthorityUrl(ByVal endpointName As String) As String
    AuthorityUrl = "https://login.microsoftonline.com/" & _
        SharePointAuthTenantId() & _
        "/oauth2/v2.0/" & endpointName
End Function

Private Function TokenScope() As String
    TokenScope = SharePointAuthResource() & _
        "/.default offline_access openid profile"
End Function

Private Sub ValidateAuthConfiguration()
    If SharePointAuthTenantId() Like "REPLACE-*" Or _
       SharePointAuthClientId() Like "REPLACE-*" Or _
       SharePointAuthResource() Like "*contoso.sharepoint.com*" Then
        Err.Raise vbObjectError + 2720, "GetSharePointAccessToken", _
            "Configure SharePointAuthConfig.bas before requesting a token."
    End If
End Sub

Private Sub RaiseOAuthError(ByVal context As String, _
                            ByVal response As Object)
    Dim errorCode As String
    Dim description As String
    Dim correlationId As String

    errorCode = JsonString(response.responseText, "error")
    description = JsonString(response.responseText, "error_description")
    correlationId = JsonString(response.responseText, "correlation_id")

    Err.Raise vbObjectError + 2730, "SharePointDeviceAuth", _
        context & ": HTTP " & response.Status & _
        "; error=" & errorCode & _
        "; correlation_id=" & correlationId & _
        "; description=" & description
End Sub

Private Sub WaitSeconds(ByVal seconds As Long)
    Dim finishAt As Date

    finishAt = DateAdd("s", seconds, Now)
    Do While Now < finishAt
        DoEvents
        Sleep 100
    Loop
End Sub

Private Function FormUrlEncode(ByVal value As String) As String
    Dim index As Long
    Dim characterCode As Long
    Dim character As String
    Dim result As String

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        characterCode = AscW(character)

        If (characterCode >= 48 And characterCode <= 57) Or _
           (characterCode >= 65 And characterCode <= 90) Or _
           (characterCode >= 97 And characterCode <= 122) Or _
           character = "-" Or character = "." Or _
           character = "_" Or character = "~" Then
            result = result & character
        ElseIf characterCode >= 0 And characterCode <= 127 Then
            result = result & "%" & Right$("0" & Hex$(characterCode), 2)
        Else
            Err.Raise vbObjectError + 2740, "FormUrlEncode", _
                "Only ASCII OAuth parameter values are supported."
        End If
    Next index

    FormUrlEncode = result
End Function

Private Function JsonString(ByVal json As String, _
                            ByVal key As String) As String
    Dim keyPosition As Long
    Dim valuePosition As Long
    Dim currentCharacter As String
    Dim escaped As Boolean
    Dim result As String

    keyPosition = InStr(1, json, Chr$(34) & key & Chr$(34), vbBinaryCompare)
    If keyPosition = 0 Then Exit Function

    valuePosition = InStr(keyPosition + Len(key) + 2, json, ":", _
        vbBinaryCompare)
    If valuePosition = 0 Then Exit Function
    valuePosition = valuePosition + 1

    Do While valuePosition <= Len(json) And _
             InStr(1, " " & vbTab & vbCr & vbLf, _
                   Mid$(json, valuePosition, 1), vbBinaryCompare) > 0
        valuePosition = valuePosition + 1
    Loop

    If Mid$(json, valuePosition, 1) <> Chr$(34) Then Exit Function
    valuePosition = valuePosition + 1

    Do While valuePosition <= Len(json)
        currentCharacter = Mid$(json, valuePosition, 1)
        If escaped Then
            Select Case currentCharacter
                Case Chr$(34), "\", "/"
                    result = result & currentCharacter
                Case "b"
                    result = result & Chr$(8)
                Case "f"
                    result = result & Chr$(12)
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case "u"
                    If valuePosition + 4 <= Len(json) Then
                        result = result & ChrW$(CLng("&H" & _
                            Mid$(json, valuePosition + 1, 4)))
                        valuePosition = valuePosition + 4
                    End If
                Case Else
                    result = result & currentCharacter
            End Select
            escaped = False
        ElseIf currentCharacter = "\" Then
            escaped = True
        ElseIf currentCharacter = Chr$(34) Then
            JsonString = result
            Exit Function
        Else
            result = result & currentCharacter
        End If
        valuePosition = valuePosition + 1
    Loop
End Function

Private Function JsonLong(ByVal json As String, ByVal key As String, _
                          ByVal defaultValue As Long) As Long
    Dim keyPosition As Long
    Dim valuePosition As Long
    Dim endPosition As Long
    Dim currentCharacter As String
    Dim numberText As String

    JsonLong = defaultValue
    keyPosition = InStr(1, json, Chr$(34) & key & Chr$(34), vbBinaryCompare)
    If keyPosition = 0 Then Exit Function

    valuePosition = InStr(keyPosition + Len(key) + 2, json, ":", _
        vbBinaryCompare)
    If valuePosition = 0 Then Exit Function
    valuePosition = valuePosition + 1

    Do While valuePosition <= Len(json) And _
             InStr(1, " " & vbTab & vbCr & vbLf, _
                   Mid$(json, valuePosition, 1), vbBinaryCompare) > 0
        valuePosition = valuePosition + 1
    Loop

    endPosition = valuePosition
    Do While endPosition <= Len(json)
        currentCharacter = Mid$(json, endPosition, 1)
        If InStr(1, "-0123456789", currentCharacter, _
                 vbBinaryCompare) = 0 Then Exit Do
        numberText = numberText & currentCharacter
        endPosition = endPosition + 1
    Loop

    If Len(numberText) > 0 Then JsonLong = CLng(numberText)
End Function
