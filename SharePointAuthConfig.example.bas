Attribute VB_Name = "SharePointAuthConfig"
Option Explicit

Public Function SharePointAuthTenantId() As String
    SharePointAuthTenantId = "REPLACE-WITH-TENANT-ID"
End Function

Public Function SharePointAuthClientId() As String
    SharePointAuthClientId = "REPLACE-WITH-PUBLIC-CLIENT-ID"
End Function

Public Function SharePointAuthResource() As String
    SharePointAuthResource = "https://contoso.sharepoint.com"
End Function

Public Function SharePointAuthSiteUrl() As String
    SharePointAuthSiteUrl = SharePointAuthResource() & _
        "/sites/REPLACE-WITH-SITE-PATH"
End Function

Public Function SharePointAuthListName() As String
    SharePointAuthListName = "REPLACE-WITH-LIST-TITLE"
End Function

Public Function SharePointAuthRestProbeUrl() As String
    SharePointAuthRestProbeUrl = SharePointAuthSiteUrl() & _
        "/_api/Web/Lists/getbytitle('" & _
        UrlEncodePathValue(SharePointAuthListName()) & "')/items?$top=1"
End Function

Public Function SharePointAuthSoapUrl() As String
    SharePointAuthSoapUrl = SharePointAuthSiteUrl() & _
        "/_vti_bin/Lists.asmx"
End Function

Private Function UrlEncodePathValue(ByVal value As String) As String
    UrlEncodePathValue = Replace(value, " ", "%20")
End Function
