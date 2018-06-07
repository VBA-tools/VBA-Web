# 4.1.0

- Update `UrlEncode` behavior to target different encoding RFCs based on `UrlEncodingMode`
- Add `UrlEncodingMode` with support for `Strict`, `Form`, `Query`, `Cookie`, and `Path`
- `UrlEncodingMode.StrictUrlEncoding` uses [RFC 3986](https://tools.ietf.org/html/rfc3986) and is the default
- `UrlEncodingMode.FormUrlEncoding` uses HTML5 form url-encoding and is used with `WebFormat.FormUrlEncoded`
- `UrlEncodingMode.QueryUrlEncoding` uses subset of `Strict` and `Form` for default querystring encoding
- `UrlEncodingMode.CookieUrlEncoding` uses [RFC 6265](https://tools.ietf.org/html/rfc6265)
- `UrlEncodingMode.PathUrlEncoding` uses "pchar" from [RFC 3986](https://tools.ietf.org/html/rfc3986) and is the default
- Update VBA-JSON to v2.2.2
- __4.1.1__ Adjust `CookieUrlEncoding` mode to match value encoding in RFC 6265 (rather than name encoding)
- __4.1.2__ Compatibility with 64-bit Mac
- __4.1.3__ Mac bugfix for % encoding
- __4.1.4__ Fix compilation issues for 64-bit Mac
- __4.1.5__ Update VBA-JSON to v2.3.0 (fixes JSON slowdown on Windows)

# 4.0.0

Major Changes:

- __Mac Support!__
- General VBA support (no Excel-specific code)
- Custom formatters
- Automatic proxy detection
- Windows authentication
- Switch to `WinHttpRequest` (Windows' modern web library)
- General API cleanup and bugfixes
- __4.0.1__ cURL escape parameters in authenticators, url-encode UrlSegments, and add `SetHeader`
- __4.0.2__ Add `Base64Decode`
- __4.0.3__ Fix out-of-date workbooks
- __4.0.4__ Move `Application.OnTime` to `WebAsyncWrapper` and add dislaimer that it's Excel-only
- __4.0.5__ Fix incorrect regional guard in ParseNumber in VBA-JSON (upgrade to v1.0.1)
- __4.0.6__ Resolve 64-bit compilation issues in VBA-JSON
- __4.0.7__ Handle resolve error when offline as timeout and add `EnableCustomFormatting` flag for `Application.Run` issues
- __4.0.8__ Fix a critical bug that caused Excel to crash if an error was thrown with AutoProxy
- __4.0.9__ Add VBA-Dictionary to installer and fix `AddBodyParameter` bug in 64-bit Excel
- __4.0.10__ Fix `SysAllocString` bug for 32-bit Excel
- __4.0.11__ Fix extract header bug in `DigestAuthenticator`
- __4.0.12__ Fix revocation check bug for `Insecure`, `Split` bug in credentials, and fix Example workbook
- __4.0.13__ Fix 100 Continue bug for Mac, fix regional issues for url-encoded numbers and dates
- __4.0.14__ Fix cached Body issue with AddBodyParameter
- __4.0.15__ Fix cookie decoding issue
- __4.0.16__ Add Access support to installer and fix installer bugs (long paths on Mac, check if files exist, password protected)
- __4.0.17__ Add `FollowRedirects` and follow redirects by default, convert stored body to Variant, fix multiple 100 Continue bug
- __4.0.18__ Add `VBA.Randomize` to `CreateNonce` and add `TodoistAuthenticator`
- __4.0.19__ Fix installer and update VBA-JSON to v1.0.3
- __4.0.20__ Update VBA-JSON to v2.0.1 (Note: Breaking change in VBA-JSON, Solidus is no longer escaped by default)
- __4.0.21__ Fix `vbCrLf` issue in Excel for Mac 2016 and namespace internal method calls
- __4.0.22__ Add `Request.UserAgent`, sort OAuth1 querystring parameters, fix `UrlEncode` issues, and `WebClient` tweaks

Breaking Changes:

- Excel-REST is now VBA-Web and all classes/modules have been renamed
- `ExecuteAsync` is Window-only and has been moved to `WebAsyncWrapper`
- `{format}` UrlSegment is no longer automatically replaced
- Many methods removed, renamed, or moved (see the [Upgrade Guide](https://github.com/VBA-tools/VBA-Web/wiki/Upgrading-from-v3.*-to-v4.*#6-replaceremove-vba-web-incompatibilities) for a detailed breakdown)

# 3.1.0

- Add `Request.RequestFormat`, `Request.ResponseFormat`, and `Request.Accept` for setting separate request and response formats (e.g. form-urlencoded request with json response)
- Add `LogRequest` and `LogResponse` for better logging detail (enable with `RestHelpers.EnableLogging = True`)
- Allow headers and content-type to be set in authenticator `BeforeExecute`
- __3.1.1__ Fix importing class incorrectly as module bug
- __3.1.2__ Add XML and plain text formats
- __3.1.3__ Fix hard dependency for XML
- __3.1.4__ Fix logging in `PrepareProxyForHttpRequest`

# 3.0.0

- Add `Client.GetJSON` and `Client.PostJSON` helpers to GET and POST JSON without setting up request
- Add `AfterExecute` to `IAuthenticator` (Breaking change, all IAuthenticators must implement this new method)
- __3.0.1__ Add `DigestAuthenticator`, new helpers, and cleanup
- __3.0.2__ Switch timeout to `Long` and remove `RestClientBase` (out of sync with v3)
- __3.0.3__ Update OAuth1, deprecate `IncludeCacheBreaker`, update True/False formatting to lowercase, add LinkedIn example
- __3.0.4__ Fix formatting of parameters with spaces for OAuth1 and add logging
- __3.0.5__ Allow Array and Collection for Body in `Request.AddBody` and `Client.PostJSON`
- __3.0.6__ Convert Empty to `null` for json
- __3.0.7__ Add `install.bat` script for easy installation and upgrade

# 2.3.0

- Add `form-urlencoded` format and helpers
- Combine Body + Parameters and Querystring + Parameters with priority given to Body or Querystring, respectively

# 2.2.0

- Add cookies support with `Request.AddCookie(key, value)` and `Response.Cookies`
- __2.2.1__ Add `Response.Headers` collection of response headers

# 2.1.0

- Add Microsoft Scripting Runtime dependency (for Dictionary support)
- Add `RestClient.SetProxy` for use in proxy environments
- __2.1.1__ Use `Val` for number parsing in locale-dependent settings
- __2.1.2__ Add raw binary `Body` to `RestResponse` for handling files (thanks [@berkus](https://github.com/berkus))
- __2.1.3__ Bugfixes and refactor

# 2.0.0

- Remove JSONLib dependency (merged with RestHelpers)
- Add RestClientBase for future use with extension for single-client applications
- Add build scripts for import/export
- New specs and bugfixes
- __2.0.1__ Handle duplicate keys when parsing json
- __2.0.2__ Add Content-Length header and 408 status code for timeout

# 1.1.0

- Integrate Excel-TDD to fully test Excel-REST library
- Handle timeouts for sync and async requests
- Remove reference dependencies and use CreateObject instead
- Add cachebreaker as querystring param only
- Add Join helpers to resolve double-slash issue between base and resource url
- Only add "?" for querystring if querystring will be created and "?" isn't present
- Only put parameters in body if there are parameters

# 0.2.0

- Add async support
