// Shared Graph helpers for SharePoint sequences (Rhino compatible).

function gs_isBlank(value) {
  return value === null || value === undefined || String(value).trim().length === 0;
}

function gs_safeString(value) {
  return value === null || value === undefined ? "" : String(value);
}

function gs_defaultString(value, defaultValue) {
  return gs_isBlank(value) ? defaultValue : String(value);
}

function gs_require(name, value) {
  if (gs_isBlank(value)) {
    throw new java.lang.IllegalArgumentException("Missing required input: " + name);
  }
  return String(value);
}

function gs_errorClass(error) {
  try {
    if (error && typeof error.errorClass === "string") {
      return error.errorClass;
    }
    if (error && error.javaException) {
      return String(error.javaException.getClass().getName());
    }
    if (error && error.getClass) {
      return String(error.getClass().getName());
    }
  } catch (ignoreErrorClass) {
  }
  return "";
}

function gs_responseOk(graphOperation, tokenMode, requiredPermission, data) {
  return {
    ok: true,
    status: "OK",
    graphOperation: graphOperation,
    tokenMode: tokenMode,
    requiredPermission: requiredPermission,
    data: data
  };
}

function gs_responseError(graphOperation, tokenMode, requiredPermission, error) {
  var http = gs_httpStatusOf(error);
  var message = "";
  try {
    message = String(error.message);
  } catch (ignoreMessage) {
    message = String(error);
  }
  return {
    ok: false,
    status: "ERROR",
    graphOperation: graphOperation,
    tokenMode: tokenMode,
    requiredPermission: requiredPermission,
    errorClass: gs_errorClass(error),
    errorMessage: message,
    httpStatus: http.status,
    retryAfter: http.retryAfter
  };
}

// The error of an HTTP answer that is not a success: the message as before, its status and its Retry-After kept for
// gs_responseError (httpStatus, retryAfter), so a caller can wait when Graph throttles (429, 503) or sign in again (401).
function gs_httpFailure(label, result) {
  var error = new Error(label + " (" + result.statusCode + "): " + gs_safeString(result.body));
  error.errorClass = "java.lang.RuntimeException";
  error.httpStatus = Number(result.statusCode);
  error.retryAfter = gs_safeString(result.retryAfter);
  return error;
}

// The HTTP status and the Retry-After of an error: one of gs_httpFailure, or an ApiException of the Graph SDK.
// {status: null, retryAfter: ""} for an error that is not an HTTP answer (no network, a wrong input...).
function gs_httpStatusOf(error) {
  var found = { status: null, retryAfter: "" };
  try {
    if (error && error.httpStatus !== undefined && error.httpStatus !== null) {
      found.status = Number(error.httpStatus);
      found.retryAfter = gs_safeString(error.retryAfter);
      return found;
    }
    var javaError = error && error.javaException ? error.javaException : null;
    if (javaError !== null && javaError instanceof Packages.com.microsoft.kiota.ApiException) {
      found.status = Number(javaError.getResponseStatusCode());
      var headers = javaError.getResponseHeaders();
      if (headers !== null) {
        var names = headers.keySet().iterator();
        while (names.hasNext()) {
          var name = names.next();
          var values = headers.get(name);
          if (String(name).toLowerCase() === "retry-after" && values !== null && !values.isEmpty()) {
            found.retryAfter = String(values.iterator().next());
          }
        }
      }
    }
  } catch (ignoreStatus) {
  }
  return found;
}

function gs_readAll(inputStream) {
  if (inputStream === null) {
    return "";
  }
  var scanner = new java.util.Scanner(inputStream, "UTF-8").useDelimiter("\\A");
  try {
    return scanner.hasNext() ? String(scanner.next()) : "";
  } finally {
    try {
      scanner.close();
    } catch (ignoreScanner) {
    }
  }
}

function gs_urlEncode(value) {
  return java.net.URLEncoder.encode(String(value), "UTF-8");
}

function gs_encodePathSegment(value) {
  return String(gs_urlEncode(value)).replace(/\+/g, "%20");
}

function gs_appendQueryParam(url, name, value) {
  if (gs_isBlank(value)) {
    return url;
  }
  var separator = url.indexOf("?") >= 0 ? "&" : "?";
  return url + separator + gs_encodePathSegment(name) + "=" + gs_encodePathSegment(value);
}

function gs_toBoolean(value, defaultValue) {
  if (value === null || value === undefined || String(value).trim().length === 0) {
    return defaultValue === true;
  }
  var normalized = String(value).toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "y";
}

function gs_parseInteger(name, value, defaultValue) {
  if (gs_isBlank(value)) {
    return defaultValue;
  }
  var parsed = parseInt(String(value), 10);
  if (isNaN(parsed)) {
    throw new java.lang.IllegalArgumentException("Invalid integer for " + name + ": " + value);
  }
  return parsed;
}

var gs_graphHttpClient = null;
var gs_graphClientByToken = {};

function gs_requireGraphJarClasses() {
  if (typeof Packages === "undefined" || typeof Packages.okhttp3 === "undefined") {
    throw new java.lang.IllegalStateException("Missing okhttp classes. Build/load libs/microsoft-graph-sharepoint-flat.jar before Graph calls.");
  }
}

function gs_getGraphHttpClient() {
  if (gs_graphHttpClient !== null) {
    return gs_graphHttpClient;
  }
  gs_requireGraphJarClasses();
  var OkHttpClient = Packages.okhttp3.OkHttpClient;
  var TimeUnit = java.util.concurrent.TimeUnit;
  gs_graphHttpClient = new OkHttpClient.Builder()
    .proxySelector(gs_engineProxySelector())
    .proxyAuthenticator(gs_engineProxyAuthenticator())
    .followRedirects(true)
    .followSslRedirects(false)
    .connectTimeout(30, TimeUnit.SECONDS)
    .readTimeout(180, TimeUnit.SECONDS)
    .writeTimeout(180, TimeUnit.SECONDS)
    .callTimeout(300, TimeUnit.SECONDS)
    .build();
  return gs_graphHttpClient;
}

// The proxy the settings of the engine give for an URI (manual or automatic, its domains without proxy), null for none, so
// Graph is reached as the other connections of the engine. Its user and password only for a basic authentication.
function gs_engineProxy(uri) {
  try {
    var manager = com.twinsoft.convertigo.engine.Engine.theApp.proxyManager;
    if (manager === null || !manager.isEnabled()) {
      return null;
    }
    var hostConfiguration = new org.apache.commons.httpclient.HostConfiguration();
    var state = new org.apache.commons.httpclient.HttpState();
    manager.setProxy(hostConfiguration, state, uri.toURL());
    var host = gs_safeString(hostConfiguration.getProxyHost());
    if (gs_isBlank(host)) {
      return null;
    }
    var proxy = { host: host, port: hostConfiguration.getProxyPort(), user: "", password: "" };
    var credentials = state.getProxyCredentials(org.apache.commons.httpclient.auth.AuthScope.ANY);
    if (credentials !== null && String(manager.proxyMethod) === "basic" &&
        credentials instanceof org.apache.commons.httpclient.UsernamePasswordCredentials) {
      proxy.user = gs_safeString(credentials.getUserName());
      proxy.password = gs_safeString(credentials.getPassword());
    }
    return proxy;
  } catch (ignoreProxy) {
    return null;
  }
}

function gs_engineProxySelector() {
  return new JavaAdapter(java.net.ProxySelector, {
    select: function(uri) {
      var proxy = gs_engineProxy(uri);
      return java.util.Collections.singletonList(proxy === null ? java.net.Proxy.NO_PROXY :
        new java.net.Proxy(java.net.Proxy.Type.HTTP, java.net.InetSocketAddress.createUnresolved(proxy.host, proxy.port)));
    },
    connectFailed: function(uri, address, failure) {
    }
  });
}

function gs_engineProxyAuthenticator() {
  return new JavaAdapter(Packages.okhttp3.Authenticator, {
    authenticate: function(route, response) {
      if (route === null || response.request().header("Proxy-Authorization") !== null) {
        return null;
      }
      var proxy = gs_engineProxy(route.address().url().uri());
      if (proxy === null || gs_isBlank(proxy.user)) {
        return null;
      }
      return response.request().newBuilder()
        .header("Proxy-Authorization", Packages.okhttp3.Credentials.basic(proxy.user, proxy.password))
        .build();
    }
  });
}

function gs_newAuthenticationProvider(token) {
  var AuthenticationProvider = Packages.com.microsoft.kiota.authentication.AuthenticationProvider;
  var normalizedToken = gs_safeString(token);
  return new JavaAdapter(AuthenticationProvider, {
    authenticateRequest: function(requestInfo, additionalAuthenticationContext) {
      if (!gs_isBlank(normalizedToken)) {
        requestInfo.headers.add("Authorization", "Bearer " + normalizedToken);
      }
    }
  });
}

function gs_getGraphClientContext(token) {
  gs_requireGraphJarClasses();
  var key = gs_safeString(token);
  var context = gs_graphClientByToken[key];
  if (context !== undefined && context !== null) {
    return context;
  }

  var GraphServiceClient = Packages.com.microsoft.graph.serviceclient.GraphServiceClient;
  var authProvider = gs_newAuthenticationProvider(key);
  var graphClient = new GraphServiceClient(authProvider, gs_getGraphHttpClient());
  context = {
    client: graphClient,
    adapter: graphClient.getRequestAdapter()
  };
  gs_graphClientByToken[key] = context;
  return context;
}

function gs_httpMethodEnum(methodUpper) {
  var HttpMethod = Packages.com.microsoft.kiota.HttpMethod;
  try {
    return HttpMethod.valueOf(String(methodUpper));
  } catch (unknownMethod) {
    throw new java.lang.IllegalArgumentException("Unsupported HTTP method: " + methodUpper);
  }
}

function gs_setRequestHeaders(requestInfo, accept, contentType, headers) {
  if (!gs_isBlank(accept)) {
    requestInfo.headers.add("Accept", String(accept));
  }
  if (!gs_isBlank(contentType)) {
    requestInfo.headers.add("Content-Type", String(contentType));
  }
  if (headers !== null && headers !== undefined) {
    for (var headerName in headers) {
      if (Object.prototype.hasOwnProperty.call(headers, headerName)) {
        var headerValue = headers[headerName];
        if (!gs_isBlank(headerValue)) {
          requestInfo.headers.add(String(headerName), String(headerValue));
        }
      }
    }
  }
}

function gs_setRequestBody(requestInfo, methodUpper, payloadString, payloadBytes, contentType) {
  var hasString = payloadString !== null && payloadString !== undefined;
  var hasBytes = payloadBytes !== null && payloadBytes !== undefined;
  if (hasString && hasBytes) {
    throw new java.lang.IllegalArgumentException("Provide either string payload or binary payload, not both");
  }
  var resolvedContentType = gs_defaultString(contentType, "application/octet-stream");
  var bodyBytes = null;

  if (!hasString && !hasBytes) {
    if (methodUpper !== "POST" && methodUpper !== "PUT" && methodUpper !== "PATCH") {
      return;
    }
    bodyBytes = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 0);
  } else if (hasBytes) {
    bodyBytes = payloadBytes;
  } else {
    bodyBytes = new java.lang.String(String(payloadString)).getBytes("UTF-8");
  }

  var bodyStream = new java.io.ByteArrayInputStream(bodyBytes);
  requestInfo.setStreamContent(bodyStream, resolvedContentType);
}

function gs_buildNativeRequest(methodUpper, endpoint, token, payloadString, payloadBytes, contentType, accept, headers) {
  var RequestInformation = Packages.com.microsoft.kiota.RequestInformation;
  var URI = java.net.URI;
  var context = gs_getGraphClientContext(token);
  var requestInfo = new RequestInformation();

  requestInfo.httpMethod = gs_httpMethodEnum(methodUpper);
  requestInfo.setUri(new URI(String(endpoint)));
  gs_setRequestHeaders(requestInfo, accept, contentType, headers);
  gs_setRequestBody(requestInfo, methodUpper, payloadString, payloadBytes, contentType);

  return context.adapter.convertToNativeRequest(requestInfo);
}

function gs_httpRequestInternal(method, endpoint, bearerToken, payloadString, payloadBytes, contentType, accept, headers, responseAsBytes) {
  gs_requireGraphJarClasses();
  var normalizedMethod = gs_defaultString(method, "GET").toUpperCase();
  var nativeRequest = gs_buildNativeRequest(
    normalizedMethod,
    endpoint,
    bearerToken,
    payloadString,
    payloadBytes,
    contentType,
    accept,
    headers
  );
  return gs_executeNativeRequest(nativeRequest, responseAsBytes);
}

function gs_executeNativeRequest(nativeRequest, responseAsBytes) {
  var response = gs_getGraphHttpClient().newCall(nativeRequest).execute();
  try {
    var statusCode = response.code();
    var responseBody = response.body();
    var body = "";
    var bytes = null;
    var errorBody = "";

    if (responseBody !== null) {
      if (responseAsBytes) {
        var rawBytes = responseBody.bytes();
        if (statusCode >= 200 && statusCode < 300) {
          bytes = rawBytes;
        } else {
          errorBody = new java.lang.String(rawBytes, "UTF-8");
        }
      } else {
        body = responseBody.string();
      }
    }

    return {
      statusCode: statusCode,
      body: body,
      bytes: bytes,
      errorBody: errorBody,
      location: gs_safeString(response.header("Location")),
      retryAfter: gs_safeString(response.header("Retry-After")),
      requestId: gs_safeString(response.header("request-id")),
      clientRequestId: gs_safeString(response.header("client-request-id")),
      etag: gs_safeString(response.header("ETag")),
      contentType: gs_safeString(response.header("Content-Type")),
      contentLengthHeader: gs_safeString(response.header("Content-Length")),
      contentDisposition: gs_safeString(response.header("Content-Disposition")),
      contentRange: gs_safeString(response.header("Content-Range"))
    };
  } finally {
    response.close();
  }
}

function gs_httpRequest(method, endpoint, bearerToken, payload, contentType, accept, headers) {
  return gs_httpRequestInternal(method, endpoint, bearerToken, payload, null, contentType, accept, headers, false);
}

function gs_httpRequestBinary(method, endpoint, bearerToken, payloadBytes, contentType, accept, headers) {
  var result = gs_httpRequestInternal(method, endpoint, bearerToken, null, payloadBytes, contentType, accept, headers, false);
  return {
    statusCode: result.statusCode,
    body: result.body,
    location: result.location,
    retryAfter: result.retryAfter,
    requestId: result.requestId,
    clientRequestId: result.clientRequestId,
    etag: result.etag,
    contentRange: result.contentRange
  };
}

function gs_httpRequestBytes(method, endpoint, bearerToken, accept, headers) {
  var result = gs_httpRequestInternal(method, endpoint, bearerToken, null, null, "", accept, headers, true);
  return {
    statusCode: result.statusCode,
    bytes: result.bytes,
    errorBody: result.errorBody,
    contentType: result.contentType,
    contentLengthHeader: result.contentLengthHeader,
    contentDisposition: result.contentDisposition
  };
}

function gs_graphGetJson(endpoint, token, headers) {
  var result = gs_httpRequest("GET", endpoint, token, null, "", "application/json", headers);
  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw gs_httpFailure("Graph request failed", result);
  }
  if (gs_isBlank(result.body)) {
    return {};
  }
  return JSON.parse(result.body);
}

function gs_graphPostJson(endpoint, token, payloadObject, headers) {
  var payload = payloadObject === null || payloadObject === undefined ? "{}" : JSON.stringify(payloadObject);
  var result = gs_httpRequest("POST", endpoint, token, payload, "application/json", "application/json", headers);
  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw gs_httpFailure("Graph request failed", result);
  }
  if (gs_isBlank(result.body)) {
    return {};
  }
  return JSON.parse(result.body);
}

function gs_graphPatchJson(endpoint, token, payloadObject, headers) {
  var payload = payloadObject === null || payloadObject === undefined ? "{}" : JSON.stringify(payloadObject);
  var result = gs_httpRequest("PATCH", endpoint, token, payload, "application/json", "application/json", headers);
  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw gs_httpFailure("Graph patch failed", result);
  }
  if (gs_isBlank(result.body)) {
    return {};
  }
  return JSON.parse(result.body);
}

function gs_graphDelete(endpoint, token, headers) {
  var result = gs_httpRequest("DELETE", endpoint, token, null, "", "application/json", headers);
  if (result.statusCode !== 404 && (result.statusCode < 200 || result.statusCode >= 300)) {
    throw gs_httpFailure("Graph delete failed", result);
  }
  return result;
}

function gs_acquireAppToken(tenantIdValue, clientIdValue, clientSecretValue, scopeValue) {
  var endpoint = "https://login.microsoftonline.com/" + gs_require("tenantId", tenantIdValue) + "/oauth2/v2.0/token";
  var payload =
    "client_id=" + gs_urlEncode(gs_require("clientId", clientIdValue)) +
    "&client_secret=" + gs_urlEncode(gs_require("clientSecret", clientSecretValue)) +
    "&scope=" + gs_urlEncode(gs_defaultString(scopeValue, "https://graph.microsoft.com/.default")) +
    "&grant_type=client_credentials";

  var tokenResponse = gs_httpRequest("POST", endpoint, "", payload, "application/x-www-form-urlencoded", "application/json", null);
  if (tokenResponse.statusCode < 200 || tokenResponse.statusCode >= 300) {
    throw gs_httpFailure("Token endpoint failed", tokenResponse);
  }

  var parsed = JSON.parse(tokenResponse.body);
  if (!parsed || gs_isBlank(parsed.access_token)) {
    throw new java.lang.RuntimeException("Token endpoint returned no access_token");
  }

  return {
    mode: "application",
    token: String(parsed.access_token),
    tokenType: gs_safeString(parsed.token_type),
    expiresIn: parsed.expires_in === undefined ? 0 : parsed.expires_in,
    scope: gs_safeString(parsed.scope)
  };
}

function gs_resolveToken(accessTokenValue, tenantIdValue, clientIdValue, clientSecretValue, scopeValue) {
  if (!gs_isBlank(accessTokenValue)) {
    return {
      mode: "delegated",
      token: String(accessTokenValue),
      tokenType: "Bearer",
      expiresIn: 0,
      scope: ""
    };
  }
  return gs_acquireAppToken(tenantIdValue, clientIdValue, clientSecretValue, gs_defaultString(scopeValue, "https://graph.microsoft.com/.default"));
}

function gs_normalizeSitePath(sitePathValue) {
  var rawPath = gs_require("sitePath", sitePathValue).trim();
  if (rawPath.length === 0) {
    throw new java.lang.IllegalArgumentException("sitePath cannot be empty");
  }
  if (rawPath.charAt(0) !== "/") {
    rawPath = "/" + rawPath;
  }
  if (rawPath.length > 1 && rawPath.charAt(rawPath.length - 1) === "/") {
    rawPath = rawPath.substring(0, rawPath.length - 1);
  }
  return rawPath;
}

function gs_encodeSitePath(sitePathValue) {
  if (sitePathValue === "/") {
    return "/";
  }
  var parts = String(sitePathValue).split("/");
  var encodedParts = [];
  var i;
  for (i = 0; i < parts.length; i++) {
    if (parts[i].length > 0) {
      encodedParts.push(gs_encodePathSegment(parts[i]));
    }
  }
  return "/" + encodedParts.join("/");
}

function gs_odataEscapeString(value) {
  return String(value).split("'").join("''");
}

function gs_splitCsv(csvValue) {
  if (gs_isBlank(csvValue)) {
    return [];
  }
  var rawParts = String(csvValue).split(",");
  var normalized = [];
  var i;
  for (i = 0; i < rawParts.length; i++) {
    var trimmed = String(rawParts[i]).trim();
    if (trimmed.length > 0) {
      normalized.push(trimmed);
    }
  }
  return normalized;
}

function gs_decodeBase64(name, rawValue) {
  var source = gs_require(name, rawValue);
  var compact = String(source).replace(/\s/g, "");
  try {
    return java.util.Base64.getDecoder().decode(compact);
  } catch (decodeError) {
    throw new java.lang.IllegalArgumentException("Invalid base64 for " + name + ": " + gs_safeString(decodeError.message));
  }
}

function gs_parseJsonObject(name, rawJson) {
  var source = gs_require(name, rawJson);
  var parsed = null;
  try {
    parsed = JSON.parse(String(source));
  } catch (parseError) {
    throw new java.lang.IllegalArgumentException("Invalid JSON for " + name + ": " + gs_safeString(parseError.message));
  }
  if (parsed === null || parsed === undefined || parsed instanceof Array || typeof parsed !== "object") {
    throw new java.lang.IllegalArgumentException(name + " must be a JSON object");
  }
  return parsed;
}

function gs_graphSerializeParsable(value) {
  if (value === null || value === undefined) {
    return {};
  }
  var JsonSerializationWriter = Packages.com.microsoft.kiota.serialization.JsonSerializationWriter;
  var writer = new JsonSerializationWriter();
  try {
    writer.writeObjectValue(null, value);
    var rawJson = gs_readAll(writer.getSerializedContent());
    if (gs_isBlank(rawJson)) {
      return {};
    }
    return JSON.parse(rawJson);
  } finally {
    try {
      writer.close();
    } catch (ignoreWriterClose) {
    }
  }
}

function gs_graphToObject(value) {
  if (value === null || value === undefined) {
    return {};
  }
  if (value instanceof Array) {
    return value;
  }
  if (value instanceof java.lang.String || value instanceof java.lang.Number || value instanceof java.lang.Boolean) {
    return value;
  }
  if (typeof value !== "object") {
    return value;
  }
  try {
    if (value instanceof Packages.com.microsoft.kiota.serialization.Parsable) {
      return gs_graphSerializeParsable(value);
    }
  } catch (ignoreParsableCheck) {
  }
  return value;
}

function gs_safeObject(value) {
  if (value === null || value === undefined) {
    return {};
  }
  var normalized = gs_graphToObject(value);
  if (normalized === null || normalized === undefined || normalized instanceof Array || typeof normalized !== "object") {
    return {};
  }
  return normalized;
}

function gs_firstDefined(objectValue, names) {
  var i;
  if (objectValue === null || objectValue === undefined) {
    return null;
  }
  for (i = 0; i < names.length; i++) {
    if (objectValue[names[i]] !== undefined && objectValue[names[i]] !== null) {
      return objectValue[names[i]];
    }
  }
  return null;
}

function gs_normalizeSite(siteObject, siteHostnameValue, sitePathValue) {
  var mapped = gs_safeObject(siteObject);
  return {
    id: gs_safeString(mapped.id),
    siteId: gs_safeString(mapped.id),
    name: gs_safeString(mapped.name),
    displayName: gs_safeString(mapped.displayName),
    webUrl: gs_safeString(mapped.webUrl),
    siteHostname: gs_safeString(siteHostnameValue),
    sitePath: gs_safeString(sitePathValue),
    siteCollection: gs_safeObject(mapped.siteCollection),
    sharepointIds: gs_safeObject(gs_firstDefined(mapped, ["sharepointIds", "sharePointIds"]))
  };
}

function gs_normalizeList(listObject) {
  var mapped = gs_safeObject(listObject);
  return {
    id: gs_safeString(mapped.id),
    listId: gs_safeString(mapped.id),
    name: gs_safeString(mapped.name),
    listName: gs_safeString(mapped.name),
    displayName: gs_safeString(mapped.displayName),
    listDisplayName: gs_safeString(mapped.displayName),
    description: gs_safeString(mapped.description),
    webUrl: gs_safeString(mapped.webUrl),
    listWebUrl: gs_safeString(mapped.webUrl),
    createdDateTime: gs_safeString(mapped.createdDateTime),
    lastModifiedDateTime: gs_safeString(mapped.lastModifiedDateTime),
    list: gs_safeObject(mapped.list),
    listInfo: gs_safeObject(mapped.list),
    sharepointIds: gs_safeObject(gs_firstDefined(mapped, ["sharepointIds", "sharePointIds"]))
  };
}

function gs_normalizeDrive(driveObject) {
  var mapped = gs_safeObject(driveObject);
  return {
    id: gs_safeString(mapped.id),
    driveId: gs_safeString(mapped.id),
    name: gs_safeString(mapped.name),
    driveName: gs_safeString(mapped.name),
    webUrl: gs_safeString(mapped.webUrl),
    driveWebUrl: gs_safeString(mapped.webUrl),
    driveType: gs_safeString(mapped.driveType),
    createdDateTime: gs_safeString(mapped.createdDateTime),
    lastModifiedDateTime: gs_safeString(mapped.lastModifiedDateTime),
    sharepointIds: gs_safeObject(gs_firstDefined(mapped, ["sharepointIds", "sharePointIds"])),
    quota: gs_safeObject(mapped.quota),
    owner: gs_safeObject(mapped.owner)
  };
}

function gs_normalizeDriveItem(itemObject, includeRaw) {
  var mapped = gs_safeObject(itemObject);
  var normalized = {
    id: gs_safeString(mapped.id),
    itemId: gs_safeString(mapped.id),
    name: gs_safeString(mapped.name),
    webUrl: gs_safeString(mapped.webUrl),
    size: mapped.size === undefined || mapped.size === null ? 0 : mapped.size,
    eTag: gs_safeString(mapped.eTag),
    cTag: gs_safeString(mapped.cTag),
    createdDateTime: gs_safeString(mapped.createdDateTime),
    lastModifiedDateTime: gs_safeString(mapped.lastModifiedDateTime),
    file: gs_safeObject(mapped.file),
    folder: gs_safeObject(mapped.folder),
    image: gs_safeObject(mapped.image),
    shared: gs_safeObject(mapped.shared),
    parentReference: gs_safeObject(mapped.parentReference),
    createdBy: gs_safeObject(mapped.createdBy),
    lastModifiedBy: gs_safeObject(mapped.lastModifiedBy),
    fileSystemInfo: gs_safeObject(mapped.fileSystemInfo),
    sharepointIds: gs_safeObject(gs_firstDefined(mapped, ["sharepointIds", "sharePointIds"])),
    removed: gs_safeObject(mapped["@removed"])
  };
  if (includeRaw) {
    normalized.raw = mapped;
  }
  return normalized;
}

function gs_normalizeItem(itemObject, includeRaw) {
  var mapped = gs_safeObject(itemObject);
  var normalized = {
    id: gs_safeString(mapped.id),
    itemId: gs_safeString(mapped.id),
    eTag: gs_safeString(mapped.eTag),
    createdDateTime: gs_safeString(mapped.createdDateTime),
    lastModifiedDateTime: gs_safeString(mapped.lastModifiedDateTime),
    webUrl: gs_safeString(mapped.webUrl),
    contentType: gs_safeObject(mapped.contentType),
    fields: gs_safeObject(mapped.fields),
    sharepointIds: gs_safeObject(gs_firstDefined(mapped, ["sharepointIds", "sharePointIds"]))
  };
  // an item of a document library expanded with its driveItem: the file (or the folder) it is
  if (mapped.driveItem !== undefined && mapped.driveItem !== null) {
    var drive = gs_safeObject(mapped.driveItem);
    var parent = gs_safeObject(drive.parentReference);
    normalized.driveItem = {
      id: gs_safeString(drive.id),
      name: gs_safeString(drive.name),
      size: drive.size === undefined || drive.size === null ? null : Number(drive.size),
      folder: drive.folder !== undefined && drive.folder !== null,
      mimeType: drive.file ? gs_safeString(gs_safeObject(drive.file).mimeType) : "",
      driveId: gs_safeString(parent.driveId),
      path: gs_safeString(parent.path),
      webUrl: gs_safeString(drive.webUrl)
    };
  }
  if (includeRaw) {
    normalized.raw = mapped;
  }
  return normalized;
}

function gs_normalizePermission(permissionObject, includeRaw) {
  var mapped = gs_safeObject(permissionObject);
  var normalized = {
    permissionId: gs_safeString(mapped.id),
    shareId: gs_safeString(mapped.shareId),
    roles: mapped.roles instanceof Array ? mapped.roles : [],
    expirationDateTime: gs_safeString(mapped.expirationDateTime),
    hasPassword: !gs_isBlank(mapped.hasPassword) ? String(mapped.hasPassword) : "",
    link: gs_safeObject(mapped.link),
    inheritedFrom: gs_safeObject(mapped.inheritedFrom),
    invitation: gs_safeObject(mapped.invitation),
    grantedToV2: gs_safeObject(mapped.grantedToV2),
    grantedToIdentitiesV2: mapped.grantedToIdentitiesV2 instanceof Array ? mapped.grantedToIdentitiesV2 : [],
    application: gs_safeObject(mapped.application)
  };
  if (includeRaw) {
    normalized.raw = mapped;
  }
  return normalized;
}

function gs_normalizeDriveItemVersion(versionObject, includeRaw) {
  var mapped = gs_safeObject(versionObject);
  var normalized = {
    versionId: gs_safeString(mapped.id),
    lastModifiedDateTime: gs_safeString(mapped.lastModifiedDateTime),
    size: mapped.size === undefined || mapped.size === null ? 0 : mapped.size,
    publication: gs_safeObject(mapped.publication),
    lastModifiedBy: gs_safeObject(mapped.lastModifiedBy)
  };
  if (includeRaw) {
    normalized.raw = mapped;
  }
  return normalized;
}

function gs_graphSiteBuilder(token, resolvedSiteId) {
  return gs_getGraphClientContext(token).client.sites().bySiteId(gs_require("siteId", resolvedSiteId));
}

function gs_graphCollectionFind(collectionBuilder, predicate) {
  var response = collectionBuilder.get();
  while (response !== null && response !== undefined) {
    var values = response.getValue();
    if (values !== null && values !== undefined) {
      var size = values.size();
      var i;
      for (i = 0; i < size; i++) {
        var current = values.get(i);
        if (predicate(current)) {
          return current;
        }
      }
    }
    var nextLink = gs_safeString(response.getOdataNextLink());
    if (gs_isBlank(nextLink)) {
      break;
    }
    response = collectionBuilder.withUrl(nextLink).get();
  }
  return null;
}

function gs_resolveSite(token, siteIdValue, siteHostnameValue, sitePathValue) {
  var resolvedSiteId = gs_safeString(siteIdValue);
  if (!gs_isBlank(resolvedSiteId)) {
    return {
      id: resolvedSiteId,
      siteId: resolvedSiteId,
      name: "",
      displayName: "",
      webUrl: "",
      siteHostname: "",
      sitePath: "",
      siteCollection: {},
      sharepointIds: {}
    };
  }

  var resolvedSiteHostname = gs_require("siteHostname", siteHostnameValue).toLowerCase();
  var resolvedSitePath = gs_normalizeSitePath(sitePathValue);
  var compositeSiteId = resolvedSiteHostname + ":" + resolvedSitePath + ":";
  var siteModel = gs_getGraphClientContext(token)
    .client
    .sites()
    .bySiteId(compositeSiteId)
    .get();
  var normalized = gs_normalizeSite(siteModel, resolvedSiteHostname, resolvedSitePath);
  normalized.siteId = gs_require("siteId", normalized.siteId);
  normalized.id = normalized.siteId;
  return normalized;
}

function gs_resolveList(token, resolvedSiteId, listIdValue, listNameValue) {
  var explicitListId = gs_safeString(listIdValue);
  var requestedListName = gs_safeString(listNameValue);
  if (gs_isBlank(explicitListId) && gs_isBlank(requestedListName)) {
    throw new java.lang.IllegalArgumentException("Provide listId or listName");
  }

  var siteBuilder = gs_graphSiteBuilder(token, resolvedSiteId);
  var listModel = null;
  if (!gs_isBlank(explicitListId)) {
    listModel = siteBuilder.lists().byListId(explicitListId).get();
  } else {
    var exactLookup = requestedListName;
    var wantedName = requestedListName.toLowerCase();
    listModel = gs_graphCollectionFind(siteBuilder.lists(), function(current) {
      var currentName = gs_safeString(current.getName());
      var currentDisplayName = gs_safeString(current.getDisplayName());
      return currentName === exactLookup ||
        currentDisplayName === exactLookup ||
        currentName.toLowerCase() === wantedName ||
        currentDisplayName.toLowerCase() === wantedName;
    });
    if (listModel === null) {
      throw new java.lang.RuntimeException("List not found: " + requestedListName);
    }
  }

  var normalized = gs_normalizeList(listModel);
  normalized.listId = gs_require("listId", normalized.listId);
  normalized.id = normalized.listId;
  return normalized;
}

function gs_findDriveByName(token, resolvedSiteId, driveNameValue) {
  var requestedDriveName = gs_require("driveName", driveNameValue);
  var wantedName = requestedDriveName.toLowerCase();
  var driveModel = gs_graphCollectionFind(gs_graphSiteBuilder(token, resolvedSiteId).drives(), function(current) {
    return gs_safeString(current.getName()).toLowerCase() === wantedName;
  });
  if (driveModel === null) {
    throw new java.lang.RuntimeException("Drive not found: " + requestedDriveName);
  }
  return gs_normalizeDrive(driveModel);
}

function gs_resolveDrive(token, siteObject, driveIdValue, driveNameValue, listIdValue, listNameValue) {
  var explicitDriveId = gs_safeString(driveIdValue);
  var requestedDriveName = gs_safeString(driveNameValue);
  var explicitListId = gs_safeString(listIdValue);
  var requestedListName = gs_safeString(listNameValue);
  if (gs_isBlank(explicitDriveId) && gs_isBlank(requestedDriveName) && gs_isBlank(explicitListId) && gs_isBlank(requestedListName)) {
    throw new java.lang.IllegalArgumentException("Provide driveId, driveName, listId or listName");
  }

  var resolvedList = null;
  var resolvedDrive = null;
  var siteBuilder = gs_graphSiteBuilder(token, siteObject.siteId);
  if (!gs_isBlank(explicitDriveId)) {
    resolvedDrive = gs_normalizeDrive(siteBuilder.drives().byDriveId(explicitDriveId).get());
  } else if (!gs_isBlank(explicitListId) || !gs_isBlank(requestedListName)) {
    resolvedList = gs_resolveList(token, siteObject.siteId, explicitListId, requestedListName);
    resolvedDrive = gs_normalizeDrive(siteBuilder.lists().byListId(resolvedList.listId).drive().get());
  } else {
    resolvedDrive = gs_findDriveByName(token, siteObject.siteId, requestedDriveName);
  }

  resolvedDrive.driveId = gs_require("driveId", resolvedDrive.driveId);
  resolvedDrive.id = resolvedDrive.driveId;
  return {
    drive: resolvedDrive,
    list: resolvedList
  };
}

function gs_resolveSiteAndList(token, siteIdValue, siteHostnameValue, sitePathValue, listIdValue, listNameValue) {
  var resolvedSite = gs_resolveSite(token, siteIdValue, siteHostnameValue, sitePathValue);
  var resolvedList = gs_resolveList(token, resolvedSite.siteId, listIdValue, listNameValue);
  return {
    siteId: resolvedSite.siteId,
    siteHostname: resolvedSite.siteHostname,
    sitePath: resolvedSite.sitePath,
    listId: resolvedList.listId,
    listName: resolvedList.listName,
    listDisplayName: resolvedList.listDisplayName,
    listWebUrl: resolvedList.listWebUrl,
    listInfo: resolvedList.listInfo,
    sharepointIds: resolvedList.sharepointIds
  };
}

function gs_resolveRootItemId(token, resolvedDriveId) {
  var rootItem = gs_getGraphClientContext(token).client.drives().byDriveId(gs_require("driveId", resolvedDriveId)).root().get();
  return gs_require("rootItemId", rootItem.getId());
}

function gs_graphQueryName(name) {
  var normalized = gs_safeString(name);
  if (normalized.indexOf("$") === 0) {
    return "%24" + normalized.substring(1);
  }
  return normalized;
}

function gs_graphAddQueryParameter(requestInfo, name, value) {
  if (requestInfo === null || requestInfo === undefined || gs_isBlank(value)) {
    return requestInfo;
  }
  requestInfo.addQueryParameter(gs_graphQueryName(name), value);
  return requestInfo;
}

function gs_graphAddQueryParameters(requestInfo, values) {
  if (values === null || values === undefined) {
    return requestInfo;
  }
  for (var key in values) {
    if (Object.prototype.hasOwnProperty.call(values, key)) {
      gs_graphAddQueryParameter(requestInfo, key, values[key]);
    }
  }
  return requestInfo;
}

function gs_graphPathParameters(token, extraParameters) {
  var params = new java.util.HashMap();
  params.put("baseurl", gs_getGraphClientContext(token).adapter.getBaseUrl());
  if (extraParameters !== null && extraParameters !== undefined) {
    for (var key in extraParameters) {
      if (Object.prototype.hasOwnProperty.call(extraParameters, key)) {
        params.put(String(key), extraParameters[key]);
      }
    }
  }
  return params;
}

function gs_newGraphRequestInformation(token, methodUpper, urlTemplate, extraParameters) {
  var RequestInformation = Packages.com.microsoft.kiota.RequestInformation;
  return new RequestInformation(
    gs_httpMethodEnum(methodUpper),
    String(urlTemplate),
    gs_graphPathParameters(token, extraParameters)
  );
}

function gs_httpRequestInfoInternal(token, requestInfo, payloadString, payloadBytes, contentType, accept, headers, responseAsBytes) {
  if (requestInfo === null || requestInfo === undefined) {
    throw new java.lang.IllegalArgumentException("requestInfo is required");
  }
  gs_setRequestHeaders(requestInfo, accept, contentType, headers);
  gs_setRequestBody(
    requestInfo,
    requestInfo.httpMethod === null || requestInfo.httpMethod === undefined ? "GET" : String(requestInfo.httpMethod),
    payloadString,
    payloadBytes,
    contentType
  );
  var nativeRequest = gs_getGraphClientContext(token).adapter.convertToNativeRequest(requestInfo);
  return gs_executeNativeRequest(nativeRequest, responseAsBytes === true);
}

function gs_graphGetJsonRequestInfo(requestInfo, token, headers) {
  var result = gs_httpRequestInfoInternal(token, requestInfo, null, null, "", "application/json", headers, false);
  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw gs_httpFailure("Graph request failed", result);
  }
  if (gs_isBlank(result.body)) {
    return {};
  }
  return JSON.parse(result.body);
}

function gs_graphPostJsonRequestInfo(requestInfo, token, payloadObject, headers) {
  var payload = payloadObject === null || payloadObject === undefined ? "{}" : JSON.stringify(payloadObject);
  var result = gs_httpRequestInfoInternal(token, requestInfo, payload, null, "application/json", "application/json", headers, false);
  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw gs_httpFailure("Graph request failed", result);
  }
  if (gs_isBlank(result.body)) {
    return {};
  }
  return JSON.parse(result.body);
}

function gs_graphPatchJsonRequestInfo(requestInfo, token, payloadObject, headers) {
  var payload = payloadObject === null || payloadObject === undefined ? "{}" : JSON.stringify(payloadObject);
  var result = gs_httpRequestInfoInternal(token, requestInfo, payload, null, "application/json", "application/json", headers, false);
  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw gs_httpFailure("Graph patch failed", result);
  }
  if (gs_isBlank(result.body)) {
    return {};
  }
  return JSON.parse(result.body);
}

function gs_graphDeleteRequestInfo(requestInfo, token, headers) {
  var result = gs_httpRequestInfoInternal(token, requestInfo, null, null, "", "application/json", headers, false);
  if (result.statusCode !== 404 && (result.statusCode < 200 || result.statusCode >= 300)) {
    throw gs_httpFailure("Graph delete failed", result);
  }
  return result;
}

function gs_graphPutBinaryRequestInfo(requestInfo, token, payloadBytes, contentType, headers) {
  var result = gs_httpRequestInfoInternal(token, requestInfo, null, payloadBytes, contentType, "application/json", headers, false);
  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw gs_httpFailure("Graph upload failed", result);
  }
  if (gs_isBlank(result.body)) {
    return {};
  }
  return JSON.parse(result.body);
}

function gs_graphGetBytesRequestInfo(requestInfo, token, headers) {
  return gs_httpRequestInfoInternal(token, requestInfo, null, null, "", "*/*", headers, true);
}

function gs_graphPostRequestInfo(requestInfo, token, payloadObject, headers) {
  var payload = payloadObject === null || payloadObject === undefined ? "{}" : JSON.stringify(payloadObject);
  return gs_httpRequestInfoInternal(token, requestInfo, payload, null, "application/json", "application/json", headers, false);
}

function gs_requestInfoUrl(requestInfo) {
  if (requestInfo === null || requestInfo === undefined) {
    return "";
  }
  return gs_safeString(requestInfo.getUri());
}

function gs_graphCollectPagedValues(initialUrl, token, maxPageCount) {
  var values = [];
  var nextLink = gs_safeString(initialUrl);
  var pageCount = 0;
  var latestDeltaLink = "";
  while (!gs_isBlank(nextLink) && pageCount < maxPageCount) {
    if (pageCount > 0 && !gs_isGraphLink(nextLink)) {
      throw new java.lang.IllegalStateException("Next link refused: not an URL of Microsoft Graph");
    }
    var pageObject = gs_graphGetJson(nextLink, token);
    var pageValues = pageObject.value instanceof Array ? pageObject.value : [];
    var i;
    for (i = 0; i < pageValues.length; i++) {
      values.push(pageValues[i]);
    }
    pageCount++;
    nextLink = gs_safeString(pageObject["@odata.nextLink"]);
    var pageDeltaLink = gs_safeString(pageObject["@odata.deltaLink"]);
    if (!gs_isBlank(pageDeltaLink)) {
      latestDeltaLink = pageDeltaLink;
    }
  }
  return {
    values: values,
    pagesFetched: pageCount,
    nextLink: nextLink,
    deltaLink: latestDeltaLink
  };
}

// The type facets of a column definition of Microsoft Graph.
var GS_COLUMN_FACETS = ["boolean", "calculated", "choice", "contentApprovalStatus", "currency", "dateTime", "geolocation",
  "hyperlinkOrPicture", "lookup", "number", "personOrGroup", "term", "text", "thumbnail"];

function gs_columnHasType(column) {
  if (column === null || typeof column !== "object") {
    return true;
  }
  if (!gs_isBlank(column.type)) {
    return true;
  }
  var i;
  for (i = 0; i < GS_COLUMN_FACETS.length; i++) {
    var facet = column[GS_COLUMN_FACETS[i]];
    if (facet !== null && facet !== undefined && typeof facet === "object") {
      return true;
    }
  }
  return false;
}

// The type a value of an item of Graph shows, for a column without type facet: "url" for a hyperlink ({Url,
// Description}), "thumbnail" for an image (the JSON text of the image: type thumbnail, fileName, serverRelativeUrl), ""
// when the value tells nothing.
function gs_columnTypeOfValue(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "object" && !(value instanceof Array) && typeof value.Url === "string") {
    return "url";
  }
  if (typeof value === "string" && value.charAt(0) === "{") {
    try {
      var image = JSON.parse(value);
      if (image !== null && typeof image === "object" && (image.type === "thumbnail" || (typeof image.serverRelativeUrl === "string" && typeof image.fileName === "string"))) {
        return "thumbnail";
      }
    } catch (ignoreJson) {
      return "";
    }
  }
  return "";
}

// Graph (v1.0 and beta) gives some columns of a list without any type facet: hyperlink, image (thumbnail),
// attachments (seen on a real tenant, 2026-11-10; only the SharePoint API v2.1 types them). Those columns get a "type"
// as that API names it: "attachment" for the Attachments column, "url" or "thumbnail" when the values of the first
// items show it (GET .../items, only those columns, maxPageCount pages of 200 at most, until each is known). Changes
// the columns in place; reading the items failing leaves them as they are.
// Answers {missing, completed, error}: the visible columns without type, the ones completed, the HTTP status of a failure.
function gs_completeColumnTypes(token, columnsUrl, columns, maxPageCount) {
  var found = { missing: 0, completed: 0, error: "" };
  var lacking = [];
  var i;
  for (i = 0; i < columns.length; i++) {
    var candidate = columns[i];
    if (!gs_columnHasType(candidate) && candidate.hidden !== true && candidate.readOnly !== true && /^[A-Za-z0-9_]+$/.test(gs_safeString(candidate.name))) {
      lacking.push(candidate);
    }
  }
  found.missing = lacking.length;
  var unknown = [];
  for (i = 0; i < lacking.length; i++) {
    if (lacking[i].name === "Attachments") {
      lacking[i].type = "attachment";
      found.completed++;
    } else {
      unknown.push(lacking[i]);
    }
  }
  var url = gs_safeString(columnsUrl);
  if (unknown.length === 0 || !gs_isGraphLink(url) || !/\/columns$/.test(url)) {
    return found;
  }
  var names = [];
  for (i = 0; i < unknown.length; i++) {
    names.push(unknown[i].name);
  }
  var next = url.replace(/\/columns$/, "/items") + "?$top=200&$expand=fields($select=" + names.join(",") + ")";
  var pages = 0;
  try {
    while (!gs_isBlank(next) && pages < maxPageCount && unknown.length > 0) {
      if (!gs_isGraphLink(next)) {
        break;
      }
      var page = gs_graphGetJson(next, token);
      pages++;
      var items = page.value instanceof Array ? page.value : [];
      var j;
      for (j = 0; j < items.length && unknown.length > 0; j++) {
        var fields = items[j] && items[j].fields ? items[j].fields : {};
        var k;
        for (k = unknown.length - 1; k >= 0; k--) {
          var type = gs_columnTypeOfValue(fields[unknown[k].name]);
          if (type !== "") {
            unknown[k].type = type;
            found.completed++;
            unknown.splice(k, 1);
          }
        }
      }
      next = gs_safeString(page["@odata.nextLink"]);
    }
  } catch (e) {
    var status = gs_httpStatusOf(e).status;
    found.error = status === null || status === undefined ? "items_failed" : String(status);
  }
  return found;
}

function gs_graphMonitorOperation(token, monitorUrl, includeMonitorBody, failOnError) {
  var requestedMonitorUrl = gs_require("monitorUrl", monitorUrl);
  var monitorResult = gs_httpRequest("GET", requestedMonitorUrl, token, null, "", "application/json", null);
  var shouldFailOnError = failOnError === undefined ? true : gs_toBoolean(failOnError, true);
  if (shouldFailOnError && monitorResult.statusCode >= 400) {
    throw gs_httpFailure("Copy monitor request failed", monitorResult);
  }

  var monitorBody = {};
  var monitorBodyText = "";
  if (gs_toBoolean(includeMonitorBody, true) && !gs_isBlank(monitorResult.body)) {
    try {
      monitorBody = JSON.parse(monitorResult.body);
    } catch (parseError) {
      monitorBodyText = monitorResult.body;
    }
  }

  var operationState = "";
  if (monitorBody && monitorBody.status !== undefined && monitorBody.status !== null) {
    operationState = gs_safeString(monitorBody.status);
  }

  return {
    monitorUrl: requestedMonitorUrl,
    statusCode: monitorResult.statusCode,
    inProgress: monitorResult.statusCode === 202,
    completed: monitorResult.statusCode >= 200 && monitorResult.statusCode < 300 && monitorResult.statusCode !== 202,
    operationState: operationState,
    retryAfter: monitorResult.retryAfter,
    location: monitorResult.location,
    requestId: monitorResult.requestId,
    clientRequestId: monitorResult.clientRequestId,
    body: monitorBody,
    bodyText: monitorBodyText
  };
}

function gs_graphUploadChunk(uploadUrl, payloadBytes, contentType, headers) {
  var result = gs_httpRequestBinary("PUT", gs_require("uploadUrl", uploadUrl), "", payloadBytes, contentType, "application/json", headers);
  var parsed = {};
  if (!gs_isBlank(result.body)) {
    parsed = JSON.parse(result.body);
  }
  return {
    statusCode: result.statusCode,
    body: result.body,
    bodyObject: parsed,
    contentRange: result.contentRange,
    eTag: result.etag,
    location: result.location
  };
}

function gs_graphClient(token) {
  return gs_getGraphClientContext(token).client;
}

function gs_graphDriveBuilder(token, driveId) {
  return gs_graphClient(token).drives().byDriveId(gs_require("driveId", driveId));
}

function gs_graphDriveItemBuilder(token, driveId, itemId) {
  var resolvedItemId = gs_safeString(itemId);
  if (gs_isBlank(resolvedItemId) || resolvedItemId.toLowerCase() === "root") {
    return null;
  }
  return gs_graphDriveBuilder(token, driveId).items().byDriveItemId(resolvedItemId);
}

function gs_graphDriveItemGetRequestInfo(token, driveId, itemId) {
  var itemBuilder = gs_graphDriveItemBuilder(token, driveId, itemId);
  if (itemBuilder === null) {
    return gs_graphDriveBuilder(token, driveId).root().toGetRequestInformation();
  }
  return itemBuilder.toGetRequestInformation();
}

function gs_graphDriveItemDeleteRequestInfo(token, driveId, itemId) {
  var itemBuilder = gs_graphDriveItemBuilder(token, driveId, itemId);
  if (itemBuilder === null) {
    throw new java.lang.IllegalArgumentException("Deleting drive root is not allowed");
  }
  return itemBuilder.toDeleteRequestInformation();
}

function gs_graphDriveItemPatchRequestInfo(token, driveId, itemId) {
  var DriveItem = Packages.com.microsoft.graph.models.DriveItem;
  var itemBuilder = gs_graphDriveItemBuilder(token, driveId, itemId);
  if (itemBuilder === null) {
    throw new java.lang.IllegalArgumentException("Patching drive root is not supported");
  }
  return itemBuilder.toPatchRequestInformation(new DriveItem());
}

function gs_graphDriveItemContentGetRequestInfo(token, driveId, itemId) {
  var itemBuilder = gs_graphDriveItemBuilder(token, driveId, itemId);
  if (itemBuilder === null) {
    return gs_graphDriveBuilder(token, driveId).root().content().toGetRequestInformation();
  }
  return itemBuilder.content().toGetRequestInformation();
}

function gs_graphDriveItemContentPutRequestInfo(token, driveId, itemId, payloadBytes) {
  var itemBuilder = gs_graphDriveItemBuilder(token, driveId, itemId);
  if (itemBuilder === null) {
    throw new java.lang.IllegalArgumentException("Uploading drive root content is not allowed");
  }
  return itemBuilder.content().toPutRequestInformation(new java.io.ByteArrayInputStream(payloadBytes));
}

function gs_graphDriveChildrenGetRequestInfo(token, driveId, parentItemId) {
  var resolvedParentId = gs_safeString(parentItemId);
  if (gs_isBlank(resolvedParentId) || resolvedParentId.toLowerCase() === "root") {
    return gs_newGraphRequestInformation(token, "GET", "{+baseurl}/drives/{drive%2Did}/root/children", {
      "drive%2Did": gs_require("driveId", driveId)
    });
  }
  return gs_graphDriveBuilder(token, driveId).items().byDriveItemId(resolvedParentId).children().toGetRequestInformation();
}

function gs_graphDriveChildrenPostRequestInfo(token, driveId, parentItemId) {
  var resolvedParentId = gs_safeString(parentItemId);
  if (gs_isBlank(resolvedParentId) || resolvedParentId.toLowerCase() === "root") {
    return gs_newGraphRequestInformation(token, "POST", "{+baseurl}/drives/{drive%2Did}/root/children", {
      "drive%2Did": gs_require("driveId", driveId)
    });
  }
  var DriveItem = Packages.com.microsoft.graph.models.DriveItem;
  return gs_graphDriveBuilder(token, driveId).items().byDriveItemId(resolvedParentId).children().toPostRequestInformation(new DriveItem());
}

function gs_graphDrivePathContentPutRequestInfo(token, driveId, parentItemId, fileName, payloadBytes) {
  var requestInfo = null;
  if (gs_isBlank(parentItemId) || String(parentItemId).toLowerCase() === "root") {
    requestInfo = gs_newGraphRequestInformation(token, "PUT", "{+baseurl}/drives/{drive%2Did}/root:/{filePath}:/content", {
      "drive%2Did": gs_require("driveId", driveId),
      "filePath": gs_require("fileName", fileName)
    });
  } else {
    requestInfo = gs_newGraphRequestInformation(token, "PUT", "{+baseurl}/drives/{drive%2Did}/items/{driveItem%2Did}:/{filePath}:/content", {
      "drive%2Did": gs_require("driveId", driveId),
      "driveItem%2Did": gs_require("parentItemId", parentItemId),
      "filePath": gs_require("fileName", fileName)
    });
  }
  requestInfo.setStreamContent(new java.io.ByteArrayInputStream(payloadBytes), "application/octet-stream");
  return requestInfo;
}

function gs_graphDriveCreateUploadSessionRequestInfo(token, driveId, itemId, parentItemId, fileName) {
  var CreateUploadSessionPostRequestBody = Packages.com.microsoft.graph.drives.item.items.item.createuploadsession.CreateUploadSessionPostRequestBody;
  var explicitItemId = gs_safeString(itemId);
  if (!gs_isBlank(explicitItemId)) {
    var itemBuilder = gs_graphDriveItemBuilder(token, driveId, explicitItemId);
    if (itemBuilder === null) {
      throw new java.lang.IllegalArgumentException("Uploading drive root is not allowed");
    }
    return itemBuilder.createUploadSession().toPostRequestInformation(new CreateUploadSessionPostRequestBody());
  }
  if (gs_isBlank(parentItemId) || String(parentItemId).toLowerCase() === "root") {
    return gs_newGraphRequestInformation(token, "POST", "{+baseurl}/drives/{drive%2Did}/root:/{filePath}:/createUploadSession", {
      "drive%2Did": gs_require("driveId", driveId),
      "filePath": gs_require("fileName", fileName)
    });
  }
  return gs_newGraphRequestInformation(token, "POST", "{+baseurl}/drives/{drive%2Did}/items/{driveItem%2Did}:/{filePath}:/createUploadSession", {
    "drive%2Did": gs_require("driveId", driveId),
    "driveItem%2Did": gs_require("parentItemId", parentItemId),
    "filePath": gs_require("fileName", fileName)
  });
}

function gs_graphDriveItemVersionsGetRequestInfo(token, driveId, itemId) {
  return gs_graphDriveBuilder(token, driveId).items().byDriveItemId(gs_require("itemId", itemId)).versions().toGetRequestInformation();
}

function gs_graphDriveItemVersionRestoreRequestInfo(token, driveId, itemId, versionId) {
  return gs_graphDriveBuilder(token, driveId)
    .items()
    .byDriveItemId(gs_require("itemId", itemId))
    .versions()
    .byDriveItemVersionId(gs_require("versionId", versionId))
    .restoreVersion()
    .toPostRequestInformation();
}

function gs_graphDriveItemPermissionsGetRequestInfo(token, driveId, itemId) {
  return gs_graphDriveBuilder(token, driveId).items().byDriveItemId(gs_require("itemId", itemId)).permissions().toGetRequestInformation();
}

function gs_graphDriveItemPermissionGetRequestInfo(token, driveId, itemId, permissionId) {
  return gs_graphDriveBuilder(token, driveId)
    .items()
    .byDriveItemId(gs_require("itemId", itemId))
    .permissions()
    .byPermissionId(gs_require("permissionId", permissionId))
    .toGetRequestInformation();
}

function gs_graphDriveItemPermissionDeleteRequestInfo(token, driveId, itemId, permissionId) {
  return gs_graphDriveBuilder(token, driveId)
    .items()
    .byDriveItemId(gs_require("itemId", itemId))
    .permissions()
    .byPermissionId(gs_require("permissionId", permissionId))
    .toDeleteRequestInformation();
}

function gs_graphDriveItemCreateLinkRequestInfo(token, driveId, itemId) {
  var CreateLinkPostRequestBody = Packages.com.microsoft.graph.drives.item.items.item.createlink.CreateLinkPostRequestBody;
  return gs_graphDriveBuilder(token, driveId)
    .items()
    .byDriveItemId(gs_require("itemId", itemId))
    .createLink()
    .toPostRequestInformation(new CreateLinkPostRequestBody());
}

function gs_graphDriveItemInviteRequestInfo(token, driveId, itemId) {
  var InvitePostRequestBody = Packages.com.microsoft.graph.drives.item.items.item.invite.InvitePostRequestBody;
  return gs_graphDriveBuilder(token, driveId)
    .items()
    .byDriveItemId(gs_require("itemId", itemId))
    .invite()
    .toPostRequestInformation(new InvitePostRequestBody());
}

function gs_graphDriveItemCopyRequestInfo(token, driveId, itemId) {
  var CopyPostRequestBody = Packages.com.microsoft.graph.drives.item.items.item.copy.CopyPostRequestBody;
  return gs_graphDriveBuilder(token, driveId)
    .items()
    .byDriveItemId(gs_require("itemId", itemId))
    .copy()
    .toPostRequestInformation(new CopyPostRequestBody());
}

function gs_graphDriveItemDeltaRequestInfo(token, driveId, parentItemId) {
  var resolvedParentId = gs_safeString(parentItemId);
  if (gs_isBlank(resolvedParentId) || resolvedParentId.toLowerCase() === "root") {
    return gs_newGraphRequestInformation(token, "GET", "{+baseurl}/drives/{drive%2Did}/root/delta", {
      "drive%2Did": gs_require("driveId", driveId)
    });
  }
  return gs_newGraphRequestInformation(token, "GET", "{+baseurl}/drives/{drive%2Did}/items/{driveItem%2Did}/delta", {
    "drive%2Did": gs_require("driveId", driveId),
    "driveItem%2Did": resolvedParentId
  });
}

function gs_graphSiteListsGetRequestInfo(token, siteId) {
  return gs_graphSiteBuilder(token, siteId).lists().toGetRequestInformation();
}

function gs_graphSiteDrivesGetRequestInfo(token, siteId) {
  return gs_graphSiteBuilder(token, siteId).drives().toGetRequestInformation();
}

function gs_graphSiteListItemsGetRequestInfo(token, siteId, listId) {
  return gs_graphSiteBuilder(token, siteId).lists().byListId(gs_require("listId", listId)).items().toGetRequestInformation();
}

function gs_graphSiteListItemsPostRequestInfo(token, siteId, listId) {
  var ListItem = Packages.com.microsoft.graph.models.ListItem;
  return gs_graphSiteBuilder(token, siteId).lists().byListId(gs_require("listId", listId)).items().toPostRequestInformation(new ListItem());
}

function gs_graphSiteListItemsDeltaRequestInfo(token, siteId, listId) {
  return gs_newGraphRequestInformation(token, "GET", "{+baseurl}/sites/{site%2Did}/lists/{list%2Did}/items/delta", {
    "site%2Did": gs_require("siteId", siteId),
    "list%2Did": gs_require("listId", listId)
  });
}

function gs_graphSiteListItemGetRequestInfo(token, siteId, listId, itemId) {
  return gs_graphSiteBuilder(token, siteId).lists().byListId(gs_require("listId", listId)).items().byListItemId(gs_require("itemId", itemId)).toGetRequestInformation();
}

function gs_graphSiteListItemDeleteRequestInfo(token, siteId, listId, itemId) {
  return gs_graphSiteBuilder(token, siteId).lists().byListId(gs_require("listId", listId)).items().byListItemId(gs_require("itemId", itemId)).toDeleteRequestInformation();
}

function gs_graphSiteListItemFieldsGetRequestInfo(token, siteId, listId, itemId) {
  return gs_graphSiteBuilder(token, siteId).lists().byListId(gs_require("listId", listId)).items().byListItemId(gs_require("itemId", itemId)).fields().toGetRequestInformation();
}

function gs_graphSiteListItemFieldsPatchRequestInfo(token, siteId, listId, itemId) {
  var FieldValueSet = Packages.com.microsoft.graph.models.FieldValueSet;
  return gs_graphSiteBuilder(token, siteId).lists().byListId(gs_require("listId", listId)).items().byListItemId(gs_require("itemId", itemId)).fields().toPatchRequestInformation(new FieldValueSet());
}

function gs_graphSiteOperationGetRequestInfo(token, siteId, operationId) {
  return gs_graphSiteBuilder(token, siteId).operations().byRichLongRunningOperationId(gs_require("operationId", operationId)).toGetRequestInformation();
}

function gs_graphSubscriptionsPostRequestInfo(token) {
  var Subscription = Packages.com.microsoft.graph.models.Subscription;
  return gs_graphClient(token).subscriptions().toPostRequestInformation(new Subscription());
}

function gs_graphSubscriptionDeleteRequestInfo(token, subscriptionId) {
  return gs_graphClient(token).subscriptions().bySubscriptionId(gs_require("subscriptionId", subscriptionId)).toDeleteRequestInformation();
}

function gs_graphBatchPostRequestInfo(token) {
  return gs_newGraphRequestInformation(token, "POST", "{+baseurl}/$batch", null);
}

// ---- reads of an import (C8Oforms): a page at a time, a site or a list by id without a lookup ---------------------

// An URL of Microsoft Graph (https://graph.microsoft.com/...), the only host a next link may send the token to.
function gs_isGraphLink(link) {
  try {
    var uri = new java.net.URI(String(link));
    return String(uri.getScheme()).toLowerCase() === "https" &&
      String(uri.getHost()).toLowerCase() === "graph.microsoft.com" &&
      uri.getPort() === -1 && uri.getRawUserInfo() === null;
  } catch (ignoreLink) {
    return false;
  }
}

// The $skiptoken of a next link of Graph: "" when there is none, or when the link is not an URL of Graph.
function gs_skipTokenOf(nextLink) {
  if (gs_isBlank(nextLink) || !gs_isGraphLink(nextLink)) {
    return "";
  }
  var query = String(nextLink).split("?")[1] || "";
  var params = query.split("&");
  var i;
  for (i = 0; i < params.length; i++) {
    var at = params[i].indexOf("=");
    var name = at === -1 ? params[i] : params[i].substring(0, at);
    if (at !== -1 && (name === "$skiptoken" || name.toLowerCase() === "%24skiptoken")) {
      return String(java.net.URLDecoder.decode(params[i].substring(at + 1), "UTF-8"));
    }
  }
  return "";
}

// The URL of a request with the $skiptoken of a next page (the templates of the SDK have no $skiptoken).
function gs_withSkipToken(url, skipToken) {
  if (gs_isBlank(skipToken)) {
    return url;
  }
  if (!/^[!-~]{1,4096}$/.test(String(skipToken))) {
    throw new java.lang.IllegalArgumentException("Invalid skipToken");
  }
  return gs_appendQueryParam(url, "$skiptoken", String(skipToken));
}

// A site by its id (GET /sites/{id}).
function gs_getSite(token, siteIdValue) {
  var requestInfo = gs_graphSiteBuilder(token, gs_require("siteId", siteIdValue)).toGetRequestInformation();
  gs_graphAddQueryParameter(requestInfo, "$select", "id,name,displayName,webUrl,siteCollection,sharepointIds");
  var normalized = gs_normalizeSite(gs_graphGetJsonRequestInfo(requestInfo, token), "", "");
  normalized.siteId = gs_require("siteId", normalized.siteId);
  normalized.id = normalized.siteId;
  return normalized;
}

// The site and the list of a call. Given by their ids with describe false, they are used as they are, with no call to
// Graph (a page of an import); else as gs_resolveSiteAndList.
function gs_siteAndList(token, siteIdValue, siteHostnameValue, sitePathValue, listIdValue, listNameValue, describe) {
  if (!describe && !gs_isBlank(siteIdValue) && !gs_isBlank(listIdValue)) {
    return {
      siteId: String(siteIdValue),
      siteHostname: "",
      sitePath: "",
      listId: String(listIdValue),
      listName: "",
      listDisplayName: "",
      listWebUrl: "",
      listInfo: {},
      sharepointIds: {}
    };
  }
  return gs_resolveSiteAndList(token, siteIdValue, siteHostnameValue, sitePathValue, listIdValue, listNameValue);
}

// The sites whose name matches (GET /sites?search=), "*" for all the sites the token reaches.
function gs_graphSitesSearchUrl(token, searchValue, selectFields) {
  var search = gs_defaultString(searchValue, "*").replace(/[\u0000-\u001f]/g, " ");
  var url = String(gs_getGraphClientContext(token).adapter.getBaseUrl()) + "/sites?search=" + gs_encodePathSegment(search);
  return gs_appendQueryParam(url, "$select", selectFields.join(","));
}

// The sites a user can open without searching, then the ones the search finds: the sites the user follows (GET
// /me/followedSites, a delegated token), the root site of the tenant (GET /sites/root), then GET /sites?search=. The
// search of Graph relies on Microsoft Search and fails on some tenants (500 generalException): its failure is answered
// (searchFailed: its HTTP status), the other sites still are, those whose name has the text searched (any for * or "").
// {items: [normalized sites, each once], followed: n, searchFailed}.
function gs_sitesOfUser(token, searchText, fields, maxPageCount) {
  var base = String(gs_getGraphClientContext(token).adapter.getBaseUrl());
  var wanted = gs_safeString(searchText).trim();
  var all = wanted === "" || wanted === "*";
  var lower = wanted.toLowerCase();
  var found = { items: [], followed: 0, searchFailed: "" };
  var seen = {};
  var add = function (site, filter) {
    var normalized = gs_normalizeSite(site, "", "");
    var id = gs_safeString(normalized.siteId || normalized.id);
    if (id === "" || seen[id]) {
      return false;
    }
    if (filter && !all && (gs_safeString(normalized.displayName) + " " + gs_safeString(normalized.name) + " " + gs_safeString(normalized.webUrl)).toLowerCase().indexOf(lower) === -1) {
      return false;
    }
    seen[id] = true;
    found.items.push(normalized);
    return true;
  };
  var select = "$select=" + fields.join(",");
  var i;
  try {
    var followed = gs_graphCollectPagedValues(base + "/me/followedSites?" + select, token, maxPageCount).values;
    for (i = 0; i < followed.length; i++) {
      if (add(followed[i], true)) {
        found.followed++;
      }
    }
  } catch (ignoreFollowed) {
    // an application token has no user: no followed sites
  }
  try {
    add(gs_graphGetJson(base + "/sites/root?" + select, token), true);
  } catch (ignoreRoot) {
  }
  try {
    var searched = gs_graphCollectPagedValues(gs_graphSitesSearchUrl(token, all ? "*" : wanted, fields), token, maxPageCount).values;
    for (i = 0; i < searched.length; i++) {
      add(searched[i], false);
    }
  } catch (e) {
    var status = gs_httpStatusOf(e).status;
    found.searchFailed = status === null || status === undefined ? "failed" : String(status);
  }
  return found;
}

function gs_graphSiteListColumnsGetRequestInfo(token, siteId, listId) {
  return gs_graphSiteBuilder(token, siteId).lists().byListId(gs_require("listId", listId)).columns().toGetRequestInformation();
}

// GET /me as an URL: the request of the SDK for /me is /users/TokenToReplace, which only a middleware of the SDK client
// rewrites, and the HTTP client of this library has none.
function gs_graphMeUrl(token, selectFields) {
  var url = String(gs_getGraphClientContext(token).adapter.getBaseUrl()) + "/me";
  return gs_appendQueryParam(url, "$select", selectFields);
}

// ---- SharePoint REST of SharePoint Online: the attachments of list items and the files of a site, read with a
// delegated token of the host of the site (audience https://<tenant>.sharepoint.com, AllSites.Read) ----------------

// The files a project hands over to another one (lib_BaseRow uploads from there only): <java.io.tmpdir>/c8o-file-transfer.
var GS_TRANSFER_DIR = "c8o-file-transfer";
var GS_MAX_FILE_BYTES = 52428800;
var GS_GUID = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

// The URL of a site for its REST API: https, no user, port, query nor fragment, no trailing slash.
function gs_restSiteUrl(siteUrlValue) {
  var value = gs_require("siteUrl", siteUrlValue).trim().replace(/\/+$/, "");
  var uri = null;
  try {
    uri = new java.net.URI(value);
  } catch (ignoreUri) {
    throw new java.lang.IllegalArgumentException("Invalid siteUrl");
  }
  if (String(uri.getScheme()).toLowerCase() !== "https" || gs_isBlank(uri.getHost()) || uri.getRawUserInfo() !== null ||
      uri.getPort() !== -1 || uri.getRawQuery() !== null || uri.getRawFragment() !== null ||
      /(^|\/)\.\.?(\/|$)/.test(gs_safeString(uri.getPath()))) {
    throw new java.lang.IllegalArgumentException("siteUrl must be the https URL of a site");
  }
  return value;
}

// GET of the REST API with Accept application/json;odata=nometadata; an error keeps the status and Retry-After.
function gs_restGetJson(url, token) {
  var result = gs_httpRequest("GET", url, token, null, "", "application/json;odata=nometadata", null);
  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw gs_httpFailure("SharePoint request failed", result);
  }
  return gs_isBlank(result.body) ? {} : JSON.parse(result.body);
}

// The attachments of an item: /_api/web/lists(guid'<list>')/items(<id>)/AttachmentFiles.
function gs_restAttachmentsUrl(siteUrl, listIdValue, itemIdValue) {
  var listId = gs_require("listId", listIdValue).replace(/^\{|\}$/g, "");
  var itemId = gs_require("itemId", itemIdValue);
  if (!GS_GUID.test(listId) || !/^[1-9][0-9]{0,9}$/.test(itemId)) {
    throw new java.lang.IllegalArgumentException("Invalid listId or itemId");
  }
  return siteUrl + "/_api/web/lists(guid'" + listId.toLowerCase() + "')/items(" + itemId + ")/AttachmentFiles?$select=FileName,ServerRelativeUrl";
}

// The content of a file by its server relative path (/sites/x/Lists/y/Attachments/1/a b.pdf): each segment encoded,
// its single quotes doubled for the OData literal.
function gs_restFileUrl(siteUrl, pathValue) {
  var path = gs_require("serverRelativeUrl", pathValue);
  if (path.charAt(0) !== "/" || path.length > 2000 || /[\u0000-\u001f]/.test(path) || /(^|\/)\.\.?(\/|$)/.test(path)) {
    throw new java.lang.IllegalArgumentException("Invalid serverRelativeUrl");
  }
  var encoded = path.split("/").map(function(segment) {
    return encodeURIComponent(segment.split("'").join("''"));
  }).join("/");
  return siteUrl + "/_api/web/GetFileByServerRelativePath(decodedurl='" + encoded + "')/$value";
}

// The transfer folder, created when missing.
function gs_transferDir() {
  var dir = new java.io.File(java.lang.System.getProperty("java.io.tmpdir"), GS_TRANSFER_DIR);
  if (!dir.isDirectory() && !dir.mkdirs() && !dir.isDirectory()) {
    throw new java.lang.IllegalStateException("The transfer folder cannot be created");
  }
  return dir.getCanonicalFile();
}

// A file name that is a name only: no folder, no character a file system refuses, 150 characters at most.
function gs_safeFileName(nameValue) {
  var name = gs_safeString(nameValue).replace(/[\\\/:*?"<>|\u0000-\u001f]/g, "_").replace(/^[.\s]+/, "").trim();
  if (name.length > 150) {
    var dot = name.lastIndexOf(".");
    var extension = dot > 0 && name.length - dot <= 10 ? name.substring(dot) : "";
    name = name.substring(0, 150 - extension.length) + extension;
  }
  return name.length ? name : "file";
}

// Downloads an URL with a token into a new file of the transfer folder (in a folder of its own, under its own name):
// {path, fileName, size, contentType}. More than maxBytes: an error of status 413, nothing kept. The caller deletes the
// file (and its folder) once used.
function gs_downloadToTransferFile(url, token, fileNameValue, maxBytesValue) {
  var maxBytes = maxBytesValue > 0 ? maxBytesValue : GS_MAX_FILE_BYTES;
  var response = gs_getGraphHttpClient().newCall(gs_buildNativeRequest("GET", url, token, null, null, "", "*/*", null)).execute();
  var folder = null;
  var file = null;
  var kept = false;
  try {
    var statusCode = response.code();
    if (statusCode < 200 || statusCode >= 300) {
      var errorBody = "";
      try {
        errorBody = response.body() === null ? "" : String(response.body().string()).substring(0, 2000);
      } catch (ignoreBody) {
      }
      throw gs_httpFailure("SharePoint download failed", { statusCode: statusCode, body: errorBody, retryAfter: gs_safeString(response.header("Retry-After")) });
    }
    if (response.body() === null || response.body().contentLength() > maxBytes) {
      throw gs_httpFailure("File too large", { statusCode: 413, body: "", retryAfter: "" });
    }
    folder = new java.io.File(gs_transferDir(), String(java.util.UUID.randomUUID()));
    if (!folder.mkdir()) {
      throw new java.lang.IllegalStateException("The transfer folder cannot be created");
    }
    file = new java.io.File(folder, gs_safeFileName(fileNameValue));
    var input = response.body().byteStream();
    var output = new java.io.FileOutputStream(file);
    var size = 0;
    try {
      var buffer = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 65536);
      var read = input.read(buffer);
      while (read !== -1) {
        size += read;
        if (size > maxBytes) {
          throw gs_httpFailure("File too large", { statusCode: 413, body: "", retryAfter: "" });
        }
        output.write(buffer, 0, read);
        read = input.read(buffer);
      }
    } finally {
      output.close();
    }
    kept = true;
    return { path: String(file.getCanonicalPath()), fileName: String(file.getName()), size: size, contentType: gs_safeString(response.header("Content-Type")) };
  } finally {
    response.close();
    if (!kept) {
      try {
        if (file !== null) {
          file["delete"]();
        }
        if (folder !== null) {
          folder["delete"]();
        }
      } catch (ignoreDelete) {
      }
    }
  }
}

// The content of a drive item (a document of a library) by Graph: /drives/{driveId}/items/{itemId}/content, which
// redirects to its download URL (followed, without the token).
function gs_graphDriveItemContentUrl(token, driveIdValue, itemIdValue) {
  var driveId = gs_require("driveId", driveIdValue);
  var itemId = gs_require("itemId", itemIdValue);
  if (!/^[A-Za-z0-9!_\-]{1,200}$/.test(driveId) || !/^[A-Za-z0-9!_\-]{1,200}$/.test(itemId)) {
    throw new java.lang.IllegalArgumentException("Invalid driveId or itemId");
  }
  return String(gs_getGraphClientContext(token).adapter.getBaseUrl()) + "/drives/" + driveId + "/items/" + itemId + "/content";
}
