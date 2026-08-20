# SharePoint VBA authentication fallback demo

This repository demonstrates how a Windows Excel VBA integration can preserve
an existing ambient `MSXML2.XMLHTTP60` SharePoint request path while using
explicit delegated OAuth authentication when the ambient session is unavailable.

## Disclaimer

This sample is provided **as-is, without warranty of any kind**, for
demonstration and diagnostic purposes only. It is not a supported Microsoft
product, production library, or substitute for your organization's security,
privacy, compliance, deployment, and code-review processes. Test in a disposable
workbook and non-production SharePoint site before adapting it.

The sample never guarantees that legacy ambient authentication will remain
available. SharePoint ASP.NET `.asmx` web services are deprecated but retained
for backward compatibility; do not create new architectural dependencies on
them. Prefer Microsoft Graph or supported SharePoint REST APIs for modernization.

## What the demo does

1. Sends a read-only SharePoint REST probe using the original ambient
   `XMLHTTP60` behavior, with no bearer header.
2. If the probe succeeds, subsequent demo requests preserve that ambient path.
3. If the probe fails, the user completes Microsoft Entra device authorization.
4. The sample keeps delegated SharePoint access and refresh tokens in VBA module
   memory and adds `Authorization: Bearer <token>` to subsequent requests.
5. It records HTTP status, `SPRequestGuid`, `request-id`, server date, response
   excerpt, duration, and VBA errors. It never logs the bearer token.

The authentication decision happens before the SOAP operation. The sample does
not blindly replay a failed write, which could duplicate a successfully processed
operation whose response was lost.

## Files

- `SharePointDeviceAuth.bas`: OAuth device-code acquisition, in-memory caching,
  silent refresh, and reset functions.
- `SharePointSoapRepro.bas`: read-only ambient-first/OAuth-fallback demonstration
  using `GetListItems`.
- `SharePointAuthConfig.example.bas`: placeholder-only tenant, app, site, and list
  configuration. Rename it to `SharePointAuthConfig.bas` before importing.

## Requirements

- Windows desktop Excel with VBA enabled.
- MSXML 6 (`MSXML2.XMLHTTP.6.0`).
- A single-tenant Microsoft Entra app registration configured as a public client.
- **Allow public client flows** enabled.
- Delegated **Office 365 SharePoint Online** API permissions appropriate for the
  operations and tenant consent where required.
- The signed-in user must independently have access to the target site and list.

Do not add a client secret, certificate private key, username/password, or
persisted refresh token to a distributed workbook.

## Configure

1. Rename `SharePointAuthConfig.example.bas` to `SharePointAuthConfig.bas`.
2. Replace every `REPLACE-WITH-*` value and the `contoso.sharepoint.com` host.
3. Import all three `.bas` modules into a disposable macro-enabled workbook.
4. Run `RunSharePointReadOnlyRepro`.
5. If prompted, complete Microsoft device sign-in in the browser.

Tokens remain in memory only and are cleared when Excel closes. Persistent
cross-session single sign-on should be implemented in a signed MSAL-based helper
with an operating-system-protected token cache, not by writing tokens from VBA.

## Adapting existing requests

After a read-only probe determines that OAuth fallback is required, add the
bearer header after `Open` and before `Send`:

```vb
request.setRequestHeader "Authorization", _
    "Bearer " & GetSharePointAccessToken()
```

Do not add this header when preserving a valid ambient request path. Do not log
the token. Surface non-success status, response body, `SPRequestGuid`, and the
VBA error to users and diagnostics instead of silently continuing.

## Known boundaries

- Windows only; ActiveX/COM is unavailable in Excel for Mac.
- Device code may be blocked by Conditional Access policy.
- The simple JSON parser is intentionally limited to OAuth responses.
- The sample is synchronous to match common legacy VBA request patterns.
- Tokens are not persisted across Excel sessions.
- Production write operations require idempotency and retry review.

## Troubleshooting

- `HTTP 0` means VBA failed before receiving an HTTP response. Use the captured
  stage, VBA error number, source, and description.
- HTTP 401/403 after OAuth requires checking token tenant/audience, delegated
  consent, Conditional Access, and the user's SharePoint permissions.
- For a SharePoint service trace, preserve UTC time and `SPRequestGuid`; query
  RequestUsage by `correlationId`, then RawULS by `CorrelationId` through the
  approved support environment.
