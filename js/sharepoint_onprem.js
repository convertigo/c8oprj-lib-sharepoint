// Shared helpers for SharePoint on-prem REST sequences (Rhino compatible).

function spop_isBlank(value) {
  return value === null || value === undefined || String(value).trim().length === 0;
}

function spop_safeString(value) {
  return value === null || value === undefined ? "" : String(value);
}

function spop_defaultString(value, defaultValue) {
  return spop_isBlank(value) ? defaultValue : String(value);
}

function spop_isUnresolvedSymbol(value) {
  if (value === null || value === undefined) {
    return false;
  }
  return /^\$\{[^}]+\}$/.test(String(value).trim());
}

function spop_isConfiguredValue(value) {
  return !spop_isBlank(value) && !spop_isUnresolvedSymbol(value);
}

function spop_require(name, value) {
  if (spop_isBlank(value)) {
    throw new java.lang.IllegalArgumentException("Missing required input: " + name);
  }
  return String(value);
}

function spop_toBoolean(value, defaultValue) {
  if (value === null || value === undefined || String(value).trim().length === 0) {
    return defaultValue === true;
  }
  var normalized = String(value).toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "y" || normalized === "on";
}

function spop_parseInteger(name, value, defaultValue) {
  if (spop_isBlank(value)) {
    return defaultValue;
  }
  var parsed = parseInt(String(value), 10);
  if (isNaN(parsed)) {
    throw new java.lang.IllegalArgumentException("Invalid integer for " + name + ": " + value);
  }
  return parsed;
}

function spop_errorClass(error) {
  try {
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

function spop_responseOk(operation, tokenMode, requiredPermission, data) {
  return {
    ok: true,
    status: "OK",
    graphOperation: operation,
    tokenMode: tokenMode,
    requiredPermission: requiredPermission,
    data: data
  };
}

function spop_responseError(operation, tokenMode, requiredPermission, error) {
  var message = "";
  try {
    message = String(error.message);
  } catch (ignoreMessage) {
    message = String(error);
  }
  return {
    ok: false,
    status: "ERROR",
    graphOperation: operation,
    tokenMode: tokenMode,
    requiredPermission: requiredPermission,
    errorClass: spop_errorClass(error),
    errorMessage: message
  };
}

function spop_readAll(inputStream) {
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

function spop_readAllBytes(inputStream) {
  if (inputStream === null) {
    return null;
  }
  var baos = new java.io.ByteArrayOutputStream();
  var buffer = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 8192);
  var read = 0;
  try {
    while ((read = inputStream.read(buffer)) !== -1) {
      baos.write(buffer, 0, read);
    }
    return baos.toByteArray();
  } finally {
    try {
      inputStream.close();
    } catch (ignoreInputClose) {
    }
    try {
      baos.close();
    } catch (ignoreBaosClose) {
    }
  }
}

function spop_urlEncode(value) {
  return java.net.URLEncoder.encode(String(value), "UTF-8");
}

function spop_encodeQueryValue(value) {
  return String(spop_urlEncode(value)).split("+").join("%20");
}

function spop_encodePathSegment(value) {
  return String(spop_urlEncode(value)).split("+").join("%20");
}

function spop_ensureLeadingSlash(value) {
  var normalized = spop_defaultString(value, "").trim();
  if (normalized.length === 0) {
    return "/";
  }
  if (normalized.charAt(0) !== "/") {
    normalized = "/" + normalized;
  }
  return normalized;
}

function spop_trimTrailingSlash(value) {
  var text = spop_defaultString(value, "").trim();
  while (text.length > 1 && text.charAt(text.length - 1) === "/") {
    text = text.substring(0, text.length - 1);
  }
  return text;
}

function spop_normalizeSitePath(value) {
  var normalized = spop_ensureLeadingSlash(value);
  return spop_trimTrailingSlash(normalized);
}

function spop_encodePathKeepSlash(value) {
  var normalized = spop_ensureLeadingSlash(value);
  var parts = normalized.split("/");
  var encoded = [];
  var i;
  for (i = 0; i < parts.length; i++) {
    if (parts[i].length === 0) {
      encoded.push("");
    } else {
      encoded.push(spop_encodePathSegment(parts[i]));
    }
  }
  return encoded.join("/");
}

function spop_encodeODataStringLiteral(value) {
  return String(value).split("'").join("''");
}

function spop_escapeListTitle(value) {
  return spop_encodeODataStringLiteral(spop_require("listName", value));
}

function spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, protocol) {
  if (spop_isConfiguredValue(siteBaseUrl)) {
    return spop_trimTrailingSlash(siteBaseUrl);
  }

  var host = spop_safeString(siteHostname).trim();
  if (!spop_isConfiguredValue(host)) {
    host = spop_getConnectorServer();
  }
  host = spop_require("siteHostname", host).trim();

  var scheme = spop_normalizeProtocol(protocol, "");
  if (spop_isBlank(scheme)) {
    scheme = spop_getConnectorScheme();
  }
  if (spop_isBlank(scheme)) {
    scheme = "https";
  }

  var port = spop_getConnectorPort();
  var authority = host;
  if (port > 0 && host.indexOf(":") < 0) {
    if ((scheme === "http" && port !== 80) || (scheme === "https" && port !== 443)) {
      authority = host + ":" + port;
    }
  }

  var resolvedSitePath = sitePath;
  if (!spop_isConfiguredValue(resolvedSitePath)) {
    resolvedSitePath = spop_getConnectorBaseDir();
  }
  var path = spop_normalizeSitePath(spop_defaultString(resolvedSitePath, "/"));
  if (path === "/") {
    return scheme + "://" + authority;
  }
  return scheme + "://" + authority + path;
}

function spop_normalizeConnectorAuthenticationType(value) {
  if (!spop_isConfiguredValue(value)) {
    return "none";
  }
  var normalized = String(value).trim().toLowerCase();
  if (normalized === "basic" || normalized === "basicpreemptive" || normalized === "ntlm" || normalized === "anonymous") {
    return normalized;
  }
  return "none";
}

function spop_getConnectorAuthenticationType() {
  try {
    if (typeof context !== "undefined" && context !== null && context.requestedObject && context.requestedObject.getProject) {
      var projectObject = context.requestedObject.getProject();
      if (projectObject !== null && projectObject !== undefined && projectObject.getConnectorByName) {
        var connector = projectObject.getConnectorByName("sharepointOnPremHttp");
        if (connector !== null && connector !== undefined && connector.getAuthenticationType) {
          return spop_normalizeConnectorAuthenticationType(spop_safeString(connector.getAuthenticationType()));
        }
      }
    }
  } catch (ignoreConnectorAuthType) {
  }
  return "none";
}

function spop_getOnPremConnector() {
  try {
    if (typeof context !== "undefined" && context !== null && context.requestedObject && context.requestedObject.getProject) {
      var projectObject = context.requestedObject.getProject();
      if (projectObject !== null && projectObject !== undefined && projectObject.getConnectorByName) {
        return projectObject.getConnectorByName("sharepointOnPremHttp");
      }
    }
  } catch (ignoreConnectorLookup) {
  }
  return null;
}

function spop_compileSymbolValue(expression) {
  try {
    var Engine = Packages.com.twinsoft.convertigo.engine.Engine;
    if (Engine !== null && Engine !== undefined && Engine.theApp !== null && Engine.theApp !== undefined
      && Engine.theApp.databaseObjectsManager !== null && Engine.theApp.databaseObjectsManager !== undefined) {
      var compiled = Engine.theApp.databaseObjectsManager.getCompiledValue(String(expression));
      return spop_safeString(compiled);
    }
  } catch (ignoreCompileSymbolValue) {
  }
  return "";
}

function spop_getOnPremUsernameSymbol() {
  return spop_compileSymbolValue("${lib_Microsoft_Sharepoint.onPrem.username}");
}

function spop_getOnPremPasswordSymbol() {
  return spop_compileSymbolValue("${lib_Microsoft_Sharepoint.onPrem.password.secret}");
}

function spop_getConnectorServer() {
  var connector = spop_getOnPremConnector();
  if (connector !== null && connector !== undefined && connector.getServer) {
    return spop_safeString(connector.getServer()).trim();
  }
  return "";
}

function spop_getConnectorScheme() {
  var connector = spop_getOnPremConnector();
  if (connector !== null && connector !== undefined && connector.isHttps) {
    return connector.isHttps() ? "https" : "http";
  }
  return "";
}

function spop_getConnectorPort() {
  var connector = spop_getOnPremConnector();
  if (connector !== null && connector !== undefined && connector.getPort) {
    return connector.getPort();
  }
  return -1;
}

function spop_getConnectorBaseDir() {
  var connector = spop_getOnPremConnector();
  if (connector !== null && connector !== undefined && connector.getBaseDir) {
    return spop_safeString(connector.getBaseDir()).trim();
  }
  return "";
}

function spop_normalizeProtocol(value, defaultValue) {
  if (!spop_isConfiguredValue(value)) {
    return spop_safeString(defaultValue).trim().toLowerCase();
  }
  var rawScheme = spop_safeString(value).trim().toLowerCase();
  if (rawScheme === "true" || rawScheme === "1") {
    return "https";
  }
  if (rawScheme === "false" || rawScheme === "0") {
    return "http";
  }
  if (rawScheme === "http" || rawScheme === "https") {
    return rawScheme;
  }
  return spop_safeString(defaultValue).trim().toLowerCase();
}

function spop_normalizeNtlmTransportMode(value) {
  var normalized = spop_defaultString(value, "").trim().toLowerCase();
  if (normalized === "connector" || normalized === "connectoronly" || normalized === "connector_only") {
    return "connector_only";
  }
  if (normalized === "httpclient" || normalized === "httpclientonly" || normalized === "httpclient_only") {
    return "httpclient_only";
  }
  if (normalized === "httpclientthenconnector" || normalized === "httpclient_then_connector") {
    return "httpclient_then_connector";
  }
  if (normalized === "connectorthenhttpclient" || normalized === "connector_then_httpclient" || normalized === "auto" || normalized === "hybrid" || normalized.length === 0) {
    return "connector_then_httpclient";
  }
  return "connector_then_httpclient";
}

function spop_shouldPreferHttpClientNtlm(endpoint) {
  var normalizedEndpoint = spop_safeString(endpoint).toLowerCase();
  if (spop_isBlank(normalizedEndpoint)) {
    return false;
  }
  return normalizedEndpoint.indexOf("/_api/web/getfolderbyserverrelativeurl(") >= 0
    || normalizedEndpoint.indexOf("/_api/web/getfilebyserverrelativeurl(") >= 0
    || normalizedEndpoint.indexOf("/_api/web/folders") >= 0
    || normalizedEndpoint.indexOf("/files/") >= 0
    || normalizedEndpoint.indexOf("/versions") >= 0
    || normalizedEndpoint.indexOf("/getlistitemchangessincetoken") >= 0
    || normalizedEndpoint.indexOf("/$value") >= 0
    || normalizedEndpoint.indexOf("/moveto(") >= 0
    || normalizedEndpoint.indexOf("/copyto(") >= 0;
}

function spop_getNtlmTransportMode() {
  var defaultMode = "connector_then_httpclient";
  var connector = spop_getOnPremConnector();
  if (connector === null || connector === undefined || !connector.getComment) {
    return defaultMode;
  }

  var comment = spop_safeString(connector.getComment());
  if (spop_isBlank(comment)) {
    return defaultMode;
  }

  var marker = "ntlmTransportMode=";
  var markerIndex = comment.indexOf(marker);
  if (markerIndex < 0) {
    return defaultMode;
  }

  var rawMode = comment.substring(markerIndex + marker.length);
  var endIndex = rawMode.length;
  var i;
  for (i = 0; i < rawMode.length; i++) {
    var current = rawMode.charAt(i);
    if (current === " " || current === "\t" || current === "\r" || current === "\n" || current === ";" || current === "," || current === ")") {
      endIndex = i;
      break;
    }
  }
  rawMode = rawMode.substring(0, endIndex).trim();
  if (spop_isBlank(rawMode) || spop_isUnresolvedSymbol(rawMode)) {
    return defaultMode;
  }

  return spop_normalizeNtlmTransportMode(rawMode);
}

function spop_splitDomainQualifiedUser(rawUser) {
  var userValue = spop_safeString(rawUser).trim();
  var result = {
    hasDomain: false,
    domain: "",
    username: userValue
  };
  if (spop_isBlank(userValue)) {
    return result;
  }

  var separatorIndex = userValue.indexOf("\\");
  if (separatorIndex < 0) {
    separatorIndex = userValue.indexOf("/");
  }

  if (separatorIndex > 0 && separatorIndex < userValue.length - 1) {
    result.hasDomain = true;
    result.domain = userValue.substring(0, separatorIndex);
    result.username = userValue.substring(separatorIndex + 1);
  }
  return result;
}

function spop_buildAuth(accessToken, username, password, cookieHeader) {
  var auth = {
    mode: "onprem_anonymous",
    headers: {},
    connectorGivenUser: null,
    connectorGivenPassword: null
  };
  var connectorAuthenticationType = spop_getConnectorAuthenticationType();
  var resolvedUsername = spop_isConfiguredValue(username) ? String(username) : spop_getOnPremUsernameSymbol();
  var resolvedPassword = spop_isConfiguredValue(password) ? String(password) : spop_getOnPremPasswordSymbol();

  if (spop_isConfiguredValue(accessToken)) {
    auth.mode = "onprem_bearer";
    auth.headers.Authorization = "Bearer " + String(accessToken);
  }

  if (auth.mode !== "onprem_bearer" && spop_isConfiguredValue(cookieHeader)) {
    auth.mode = "onprem_cookie";
    auth.headers.Cookie = String(cookieHeader);
  } else if (auth.mode === "onprem_anonymous" && spop_isConfiguredValue(resolvedUsername) && spop_isConfiguredValue(resolvedPassword)) {
    if (connectorAuthenticationType === "none" || connectorAuthenticationType === "anonymous") {
      auth.mode = "onprem_basic";
      var raw = String(resolvedUsername) + ":" + String(resolvedPassword);
      var encoded = java.util.Base64.getEncoder().encodeToString(new java.lang.String(raw).getBytes("UTF-8"));
      auth.headers.Authorization = "Basic " + String(encoded);
    } else if (connectorAuthenticationType === "ntlm") {
      auth.mode = "onprem_connector_ntlm";
      auth.connectorGivenUser = String(resolvedUsername);
      auth.connectorGivenPassword = String(resolvedPassword);
    } else {
      auth.mode = "onprem_connector_basic";
      auth.connectorGivenUser = String(resolvedUsername);
      auth.connectorGivenPassword = String(resolvedPassword);
    }
  } else if (auth.mode === "onprem_anonymous") {
    if (connectorAuthenticationType === "ntlm") {
      auth.mode = "onprem_connector_ntlm";
    } else if (connectorAuthenticationType === "basic" || connectorAuthenticationType === "basicpreemptive") {
      auth.mode = "onprem_connector_basic";
    }
  }

  return auth;
}

function spop_applyHeaders(connection, auth, headers, accept, contentType) {
  if (!spop_isBlank(accept)) {
    connection.setRequestProperty("Accept", String(accept));
  }
  if (!spop_isBlank(contentType)) {
    connection.setRequestProperty("Content-Type", String(contentType));
  }

  var key;
  if (auth !== null && auth !== undefined && auth.headers !== null && auth.headers !== undefined) {
    for (key in auth.headers) {
      if (Object.prototype.hasOwnProperty.call(auth.headers, key)) {
        connection.setRequestProperty(String(key), String(auth.headers[key]));
      }
    }
  }

  if (headers !== null && headers !== undefined) {
    for (key in headers) {
      if (Object.prototype.hasOwnProperty.call(headers, key)) {
        connection.setRequestProperty(String(key), String(headers[key]));
      }
    }
  }
}

function spop_currentProjectName() {
  try {
    if (typeof context !== "undefined" && context !== null && context.requestedObject && context.requestedObject.getProject) {
      var projectObject = context.requestedObject.getProject();
      if (projectObject !== null && projectObject !== undefined && projectObject.getName) {
        return String(projectObject.getName());
      }
    }
  } catch (ignoreRequestedObject) {
  }
  return "";
}

function spop_collectHeaders(auth, headers, accept, contentType) {
  var merged = {};
  var key;

  if (!spop_isBlank(accept)) {
    merged.Accept = String(accept);
  }
  if (!spop_isBlank(contentType)) {
    merged["Content-Type"] = String(contentType);
  }

  if (auth !== null && auth !== undefined && auth.headers !== null && auth.headers !== undefined) {
    for (key in auth.headers) {
      if (Object.prototype.hasOwnProperty.call(auth.headers, key)) {
        merged[String(key)] = String(auth.headers[key]);
      }
    }
  }

  if (headers !== null && headers !== undefined) {
    for (key in headers) {
      if (Object.prototype.hasOwnProperty.call(headers, key)) {
        merged[String(key)] = String(headers[key]);
      }
    }
  }

  return merged;
}

function spop_extractTextNodes(element) {
  if (element === null || element === undefined) {
    return "";
  }
  var Node = Packages.org.w3c.dom.Node;
  var nodes = element.getChildNodes();
  var parts = [];
  var index;

  for (index = 0; index < nodes.getLength(); index++) {
    var node = nodes.item(index);
    var nodeType = node.getNodeType();
    if (nodeType === Node.TEXT_NODE || nodeType === Node.CDATA_SECTION_NODE) {
      var value = spop_safeString(node.getNodeValue());
      if (!spop_isBlank(value)) {
        parts.push(value);
      }
    }
  }

  return parts.join("");
}

function spop_findChildElement(parent, name) {
  if (parent === null || parent === undefined) {
    return null;
  }
  var Node = Packages.org.w3c.dom.Node;
  var children = parent.getChildNodes();
  var index;

  for (index = 0; index < children.getLength(); index++) {
    var child = children.item(index);
    if (child.getNodeType() === Node.ELEMENT_NODE && String(child.getNodeName()) === String(name)) {
      return child;
    }
  }

  return null;
}

function spop_findResponseHeaderValue(httpInfoElement, headerName) {
  var responseHeadersElement = spop_findChildElement(httpInfoElement, "responseHeaders");
  if (responseHeadersElement === null) {
    return "";
  }

  var Node = Packages.org.w3c.dom.Node;
  var headers = responseHeadersElement.getChildNodes();
  var index;
  var expected = String(headerName).toLowerCase();

  for (index = 0; index < headers.getLength(); index++) {
    var header = headers.item(index);
    if (header.getNodeType() === Node.ELEMENT_NODE && String(header.getNodeName()) === "header") {
      var name = spop_safeString(header.getAttribute("name"));
      if (String(name).toLowerCase() === expected) {
        return spop_safeString(header.getAttribute("value"));
      }
    }
  }

  return "";
}

function spop_bytesToUtf8(payloadBytes) {
  if (payloadBytes === null || payloadBytes === undefined) {
    return "";
  }
  return String(new java.lang.String(payloadBytes, "UTF-8"));
}

function spop_isTextContentType(contentType) {
  var normalizedType = spop_defaultString(contentType, "").trim().toLowerCase();
  if (spop_isBlank(normalizedType)) {
    return false;
  }
  return normalizedType.indexOf("application/json") === 0
    || normalizedType.indexOf("application/xml") === 0
    || normalizedType.indexOf("text/") === 0
    || normalizedType.indexOf("application/x-www-form-urlencoded") === 0;
}

function spop_pickConnectorTransaction(method, payloadBytes, contentType, binaryResponse) {
  var normalizedMethod = spop_defaultString(method, "GET").trim().toUpperCase();
  if (normalizedMethod === "GET") {
    return binaryResponse === true ? "onPremGetBinaryRequest" : "onPremGetRequest";
  }
  if (normalizedMethod === "POST") {
    if (payloadBytes !== null && payloadBytes !== undefined && !spop_isTextContentType(contentType)) {
      return "onPremBinaryRequest";
    }
    return "onPremRequest";
  }
  throw new java.lang.IllegalArgumentException("Unsupported HTTP method for on-prem connector: " + normalizedMethod);
}

function spop_decodeConnectorBase64Body(rawBody) {
  var compact = String(spop_defaultString(rawBody, "")).replace(/\s/g, "");
  if (compact.length === 0) {
    return java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 0);
  }
  try {
    return java.util.Base64.getDecoder().decode(compact);
  } catch (decodeError) {
    throw new java.lang.RuntimeException("Unable to decode connector binary payload: " + spop_safeString(decodeError.message));
  }
}

function spop_writeTempBinaryFile(payloadBytes) {
  var tempFile = java.io.File.createTempFile("spop_upload_", ".bin");
  var output = new java.io.FileOutputStream(tempFile);
  try {
    output.write(payloadBytes);
  } finally {
    try {
      output.close();
    } catch (ignoreOutputClose) {
    }
  }
  return tempFile;
}

function spop_resolveNtlmCredentials(auth) {
  var connector = spop_getOnPremConnector();
  var username = "";
  var password = "";
  var domain = "";

  if (connector !== null && connector !== undefined) {
    if (connector.getAuthUser) {
      username = spop_safeString(connector.getAuthUser());
    }
    if (connector.getAuthPassword) {
      password = spop_safeString(connector.getAuthPassword());
    }
    if (connector.getNTLMAuthenticationDomain) {
      domain = spop_safeString(connector.getNTLMAuthenticationDomain());
    }
  }

  if (auth !== null && auth !== undefined) {
    if (spop_isConfiguredValue(auth.connectorGivenUser)) {
      username = String(auth.connectorGivenUser);
    }
    if (spop_isConfiguredValue(auth.connectorGivenPassword)) {
      password = String(auth.connectorGivenPassword);
    }
  }

  if (!spop_isConfiguredValue(domain)) {
    domain = "";
  }

  var splitIdentity = spop_splitDomainQualifiedUser(username);
  if (splitIdentity.hasDomain) {
    if (!spop_isConfiguredValue(domain)) {
      domain = splitIdentity.domain;
    }
    username = splitIdentity.username;
  }

  if (!spop_isConfiguredValue(username) || !spop_isConfiguredValue(password)) {
    throw new java.lang.IllegalArgumentException("Missing NTLM credentials for on-prem connector mode");
  }

  return {
    username: String(username),
    password: String(password),
    domain: String(domain)
  };
}

function spop_httpRequestViaHttpClient5Ntlm(method, endpoint, auth, payloadBytes, contentType, accept, headers, binaryResponse) {
  var credentials = spop_resolveNtlmCredentials(auth);
  var normalizedMethod = spop_defaultString(method, "GET").trim().toUpperCase();
  var mergedHeaders = spop_collectHeaders(auth, headers, accept, contentType);

  var RequestConfig = Packages.org.apache.hc.client5.http.config.RequestConfig;
  var StandardAuthScheme = Packages.org.apache.hc.client5.http.auth.StandardAuthScheme;
  var BasicCredentialsProvider = Packages.org.apache.hc.client5.http.impl.auth.BasicCredentialsProvider;
  var AuthScope = Packages.org.apache.hc.client5.http.auth.AuthScope;
  var NTCredentials = Packages.org.apache.hc.client5.http.auth.NTCredentials;
  var HttpClients = Packages.org.apache.hc.client5.http.impl.classic.HttpClients;
  var HttpGet = Packages.org.apache.hc.client5.http.classic.methods.HttpGet;
  var HttpPost = Packages.org.apache.hc.client5.http.classic.methods.HttpPost;
  var StringEntity = Packages.org.apache.hc.core5.http.io.entity.StringEntity;
  var ByteArrayEntity = Packages.org.apache.hc.core5.http.io.entity.ByteArrayEntity;
  var ContentType = Packages.org.apache.hc.core5.http.ContentType;
  var EntityUtils = Packages.org.apache.hc.core5.http.io.entity.EntityUtils;

  var requestConfig = RequestConfig.custom()
    .setTargetPreferredAuthSchemes(java.util.Arrays.asList(StandardAuthScheme.NTLM, StandardAuthScheme.SPNEGO))
    .build();

  var credentialsProvider = new BasicCredentialsProvider();
  var passwordChars = new java.lang.String(credentials.password).toCharArray();
  credentialsProvider.setCredentials(
    new AuthScope(null, -1),
    new NTCredentials(credentials.username, passwordChars, null, credentials.domain)
  );

  var client = HttpClients.custom()
    .setDefaultCredentialsProvider(credentialsProvider)
    .setDefaultRequestConfig(requestConfig)
    .build();

  var request = null;
  if (normalizedMethod === "GET") {
    request = new HttpGet(String(endpoint));
  } else if (normalizedMethod === "POST") {
    request = new HttpPost(String(endpoint));
  } else {
    try {
      client.close();
    } catch (ignoreClientCloseMethod) {
    }
    throw new java.lang.IllegalArgumentException("Unsupported HTTP method for NTLM HttpClient5 path: " + normalizedMethod);
  }

  var headerName;
  for (headerName in mergedHeaders) {
    if (Object.prototype.hasOwnProperty.call(mergedHeaders, headerName)) {
      request.setHeader(String(headerName), String(mergedHeaders[headerName]));
    }
  }

  if (normalizedMethod === "POST" && payloadBytes !== null && payloadBytes !== undefined) {
    var entity = null;
    var safeContentType = spop_defaultString(contentType, "");
    if (spop_isTextContentType(contentType)) {
      var payloadText = new java.lang.String(payloadBytes, "UTF-8");
      entity = new StringEntity(String(payloadText), ContentType.parse(spop_isBlank(safeContentType) ? "application/json; charset=UTF-8" : safeContentType));
    } else {
      entity = new ByteArrayEntity(payloadBytes, ContentType.parse(spop_isBlank(safeContentType) ? "application/octet-stream" : safeContentType));
    }
    request.setEntity(entity);
  }

  var response = null;
  try {
    response = client.execute(request);
    var statusCode = response.getCode();
    var responseEntity = response.getEntity();
    var responseBytes = responseEntity === null ? java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 0) : EntityUtils.toByteArray(responseEntity);
    var locationHeader = response.getFirstHeader("Location");
    var etagHeader = response.getFirstHeader("ETag");
    var contentTypeHeader = response.getFirstHeader("Content-Type");
    var contentDispositionHeader = response.getFirstHeader("Content-Disposition");
    var contentLengthHeader = response.getFirstHeader("Content-Length");

    return {
      statusCode: statusCode,
      body: binaryResponse === true ? "" : String(new java.lang.String(responseBytes, "UTF-8")),
      bytes: binaryResponse === true ? responseBytes : null,
      location: locationHeader === null ? "" : spop_safeString(locationHeader.getValue()),
      etag: etagHeader === null ? "" : spop_safeString(etagHeader.getValue()),
      contentType: contentTypeHeader === null ? "" : spop_safeString(contentTypeHeader.getValue()),
      contentDisposition: contentDispositionHeader === null ? "" : spop_safeString(contentDispositionHeader.getValue()),
      contentLengthHeader: contentLengthHeader === null ? "" : spop_safeString(contentLengthHeader.getValue())
    };
  } finally {
    try {
      if (response !== null && response !== undefined) {
        response.close();
      }
    } catch (ignoreResponseClose) {
    }
    try {
      client.close();
    } catch (ignoreClientClose) {
    }
  }
}

function spop_httpRequestViaConnector(method, endpoint, auth, payloadBytes, contentType, accept, headers, binaryResponse) {
  var targetProject = spop_currentProjectName();
  if (spop_isBlank(targetProject)) {
    throw new java.lang.RuntimeException("Unable to resolve target project for on-prem HTTP connector call");
  }

  var normalizedMethod = spop_defaultString(method, "GET").trim().toUpperCase();
  var transactionName = spop_pickConnectorTransaction(normalizedMethod, payloadBytes, contentType, binaryResponse);

  var HashMap = java.util.HashMap;
  var InternalRequester = Packages.com.twinsoft.convertigo.engine.requesters.InternalRequester;
  var Engine = Packages.com.twinsoft.convertigo.engine.Engine;

  var parameters = new HashMap();
  var projectArray = java.lang.reflect.Array.newInstance(java.lang.String, 1);
  projectArray[0] = targetProject;
  parameters.put("__project", projectArray);
  parameters.put("__connector", "sharepointOnPremHttp");
  parameters.put("__transaction", transactionName);
  parameters.put("__context", "syncContext_" + java.lang.System.currentTimeMillis());
  parameters.put("__uri", String(endpoint));

  var mergedHeaders = spop_collectHeaders(auth, headers, accept, contentType);
  var headerName;
  for (headerName in mergedHeaders) {
    if (Object.prototype.hasOwnProperty.call(mergedHeaders, headerName)) {
      parameters.put("__header_" + String(headerName), String(mergedHeaders[headerName]));
    }
  }

  var requester = null;
  var tempBinaryFile = null;
  var connector = null;
  var previousGivenAuthUser = null;
  var previousGivenAuthPassword = null;
  var previousNtlmDomain = null;
  var hasNtlmDomainOverride = false;
  try {
    connector = spop_getOnPremConnector();
    if (connector !== null && connector !== undefined && connector.getGivenAuthUser && connector.getGivenAuthPassword && connector.setGivenAuthUser && connector.setGivenAuthPassword) {
      var connectorAuthenticationType = spop_getConnectorAuthenticationType();
      previousGivenAuthUser = connector.getGivenAuthUser();
      previousGivenAuthPassword = connector.getGivenAuthPassword();

      var effectiveGivenUser = auth !== null && auth !== undefined ? auth.connectorGivenUser : null;
      var effectiveGivenPassword = auth !== null && auth !== undefined ? auth.connectorGivenPassword : null;

      if (connectorAuthenticationType === "ntlm") {
        var effectiveDomain = "";
        if (connector.getNTLMAuthenticationDomain && connector.setNTLMAuthenticationDomain) {
          previousNtlmDomain = connector.getNTLMAuthenticationDomain();
          hasNtlmDomainOverride = true;

          effectiveDomain = spop_safeString(previousNtlmDomain);
          if (!spop_isConfiguredValue(effectiveDomain)) {
            effectiveDomain = "";
          }

          var sourceUser = spop_isConfiguredValue(effectiveGivenUser) ? String(effectiveGivenUser) : (connector.getAuthUser ? spop_safeString(connector.getAuthUser()) : "");
          var sourcePassword = spop_isConfiguredValue(effectiveGivenPassword) ? String(effectiveGivenPassword) : (connector.getAuthPassword ? spop_safeString(connector.getAuthPassword()) : "");
          var splitIdentity = spop_splitDomainQualifiedUser(sourceUser);

          if (splitIdentity.hasDomain) {
            if (!spop_isConfiguredValue(effectiveDomain)) {
              effectiveDomain = splitIdentity.domain;
            }
          }

          // Keep DOMAIN\user untouched for connector NTLM. Some farms reject bare usernames.
          if (!spop_isConfiguredValue(effectiveGivenUser) && spop_isConfiguredValue(sourceUser) && spop_isConfiguredValue(sourcePassword)) {
            effectiveGivenUser = sourceUser;
            effectiveGivenPassword = sourcePassword;
          }

          connector.setNTLMAuthenticationDomain(effectiveDomain);
        }
      }

      if (spop_isConfiguredValue(effectiveGivenUser) && spop_isConfiguredValue(effectiveGivenPassword)) {
        connector.setGivenAuthUser(String(effectiveGivenUser));
        connector.setGivenAuthPassword(String(effectiveGivenPassword));
      } else {
        connector.setGivenAuthUser(null);
        connector.setGivenAuthPassword(null);
      }
    }

    if (normalizedMethod === "POST" && payloadBytes !== null && payloadBytes !== undefined) {
      if (transactionName === "onPremBinaryRequest") {
        tempBinaryFile = spop_writeTempBinaryFile(payloadBytes);
        parameters.put("__body", tempBinaryFile.getAbsolutePath());
        parameters.put("__contentType", spop_defaultString(contentType, "application/octet-stream"));
      } else {
        parameters.put("payload", spop_bytesToUtf8(payloadBytes));
      }
    }

    requester = new InternalRequester(parameters, context.httpServletRequest);
    var response = requester.processRequest();
    var root = response.getDocumentElement();
    var httpInfoElement = spop_findChildElement(root, "HttpInfo");

    var statusCode = 0;
    if (httpInfoElement !== null) {
      var statusElement = spop_findChildElement(httpInfoElement, "status");
      if (statusElement !== null) {
        var rawStatusCode = spop_safeString(statusElement.getAttribute("code"));
        statusCode = spop_parseInteger("connector_status_code", rawStatusCode, 0);
      }
    }

    var rawBody = spop_extractTextNodes(root);
    var result = {
      statusCode: statusCode,
      body: rawBody,
      bytes: null,
      location: httpInfoElement === null ? "" : spop_findResponseHeaderValue(httpInfoElement, "Location"),
      etag: httpInfoElement === null ? "" : spop_findResponseHeaderValue(httpInfoElement, "ETag"),
      contentType: httpInfoElement === null ? "" : spop_findResponseHeaderValue(httpInfoElement, "Content-Type"),
      contentDisposition: httpInfoElement === null ? "" : spop_findResponseHeaderValue(httpInfoElement, "Content-Disposition"),
      contentLengthHeader: httpInfoElement === null ? "" : spop_findResponseHeaderValue(httpInfoElement, "Content-Length")
    };

    if (binaryResponse === true) {
      result.bytes = spop_decodeConnectorBase64Body(rawBody);
      result.body = "";
    }

    return result;
  } finally {
    try {
      if (tempBinaryFile !== null && tempBinaryFile !== undefined) {
        tempBinaryFile["delete"]();
      }
    } catch (ignoreTempDelete) {
    }
    try {
      org.apache.log4j.MDC.put("ContextualParameters", context.logParameters);
    } catch (ignoreMdc) {
    }
    try {
      if (requester !== null && requester !== undefined) {
        var innerContext = requester.getContext();
        if (innerContext !== null && innerContext !== undefined) {
          Engine.theApp.contextManager.remove(innerContext);
        }
      }
    } catch (ignoreContextCleanup) {
    }
    try {
      if (connector !== null && connector !== undefined && connector.setGivenAuthUser && connector.setGivenAuthPassword) {
        connector.setGivenAuthUser(previousGivenAuthUser);
        connector.setGivenAuthPassword(previousGivenAuthPassword);
      }
    } catch (ignoreGivenAuthRestore) {
    }
    try {
      if (hasNtlmDomainOverride && connector !== null && connector !== undefined && connector.setNTLMAuthenticationDomain) {
        connector.setNTLMAuthenticationDomain(previousNtlmDomain === null || previousNtlmDomain === undefined ? "" : String(previousNtlmDomain));
      }
    } catch (ignoreNtlmDomainRestore) {
    }
  }
}

function spop_httpRequest(method, endpoint, auth, payloadBytes, contentType, accept, headers, binaryResponse) {
  var effectiveHeaders = {};
  var headerName;
  if (headers !== null && headers !== undefined) {
    for (headerName in headers) {
      if (Object.prototype.hasOwnProperty.call(headers, headerName)) {
        effectiveHeaders[headerName] = headers[headerName];
      }
    }
  }
  if (auth !== null && auth !== undefined && spop_safeString(auth.mode).indexOf("onprem_") === 0) {
    if (effectiveHeaders["X-FORMS_BASED_AUTH_ACCEPTED"] === undefined && effectiveHeaders["x-forms_based_auth_accepted"] === undefined) {
      effectiveHeaders["X-FORMS_BASED_AUTH_ACCEPTED"] = "f";
    }
  }

  if (auth !== null && auth !== undefined && auth.mode === "onprem_connector_ntlm") {
    var transportMode = spop_getNtlmTransportMode();
    if (spop_shouldPreferHttpClientNtlm(endpoint)) {
      if (transportMode === "connector_then_httpclient") {
        transportMode = "httpclient_then_connector";
      } else if (transportMode === "connector_only") {
        transportMode = "httpclient_then_connector";
      }
    }
    var connectorResult = null;
    var clientResult = null;

    if (transportMode === "connector_only") {
      return spop_httpRequestViaConnector(method, endpoint, auth, payloadBytes, contentType, accept, effectiveHeaders, binaryResponse);
    }

    if (transportMode === "httpclient_only") {
      return spop_httpRequestViaHttpClient5Ntlm(method, endpoint, auth, payloadBytes, contentType, accept, effectiveHeaders, binaryResponse);
    }

    if (transportMode === "httpclient_then_connector") {
      try {
        clientResult = spop_httpRequestViaHttpClient5Ntlm(method, endpoint, auth, payloadBytes, contentType, accept, effectiveHeaders, binaryResponse);
        if (clientResult.statusCode >= 200 && clientResult.statusCode < 300) {
          return clientResult;
        }
      } catch (ignoreHttpClientFirst) {
      }
      return spop_httpRequestViaConnector(method, endpoint, auth, payloadBytes, contentType, accept, effectiveHeaders, binaryResponse);
    }

    try {
      connectorResult = spop_httpRequestViaConnector(method, endpoint, auth, payloadBytes, contentType, accept, effectiveHeaders, binaryResponse);
      if (connectorResult.statusCode >= 200 && connectorResult.statusCode < 300) {
        return connectorResult;
      }
      if (connectorResult.statusCode !== 401 && connectorResult.statusCode !== 403) {
        return connectorResult;
      }
    } catch (connectorError) {
      connectorResult = null;
    }

    try {
      return spop_httpRequestViaHttpClient5Ntlm(method, endpoint, auth, payloadBytes, contentType, accept, effectiveHeaders, binaryResponse);
    } catch (httpClientFallbackError) {
      if (connectorResult !== null && connectorResult !== undefined) {
        return connectorResult;
      }
      throw httpClientFallbackError;
    }
  }

  return spop_httpRequestViaConnector(method, endpoint, auth, payloadBytes, contentType, accept, effectiveHeaders, binaryResponse);
}

function spop_httpRequestJson(method, endpoint, auth, payloadJson, headers) {
  var payloadBytes = null;
  if (payloadJson !== null && payloadJson !== undefined) {
    payloadBytes = new java.lang.String(JSON.stringify(payloadJson)).getBytes("UTF-8");
  }

  var result = spop_httpRequest(
    method,
    endpoint,
    auth,
    payloadBytes,
    "application/json;odata=verbose;charset=utf-8",
    "application/json;odata=verbose",
    headers,
    false
  );

  if (result.statusCode < 200 || result.statusCode >= 300) {
    throw new java.lang.RuntimeException("SharePoint request failed (" + result.statusCode + "): " + result.body);
  }

  if (spop_isBlank(result.body)) {
    return {};
  }
  return JSON.parse(result.body);
}

function spop_expectSuccess(result, label) {
  if (result.statusCode >= 200 && result.statusCode < 300) {
    return;
  }
  var text = spop_safeString(result.body);
  if (spop_isBlank(text) && result.bytes !== null) {
    try {
      text = String(new java.lang.String(result.bytes, "UTF-8"));
    } catch (ignoreBytesText) {
      text = "";
    }
  }
  throw new java.lang.RuntimeException(spop_defaultString(label, "SharePoint request") + " failed (" + result.statusCode + "): " + text);
}

function spop_parseVerbose(payload) {
  if (payload === null || payload === undefined) {
    return {};
  }
  if (payload.d !== undefined && payload.d !== null) {
    return payload.d;
  }
  return payload;
}

function spop_extractArray(payload) {
  if (payload === null || payload === undefined) {
    return [];
  }
  if (payload.results !== undefined && payload.results !== null && payload.results instanceof Array) {
    return payload.results;
  }
  if (payload instanceof Array) {
    return payload;
  }
  return [];
}

function spop_decodeBase64(name, rawValue) {
  var source = spop_require(name, rawValue);
  var compact = String(source).replace(/\s/g, "");
  try {
    return java.util.Base64.getDecoder().decode(compact);
  } catch (decodeError) {
    throw new java.lang.IllegalArgumentException("Invalid base64 for " + name + ": " + spop_safeString(decodeError.message));
  }
}

function spop_encodeBase64(bytes) {
  if (bytes === null || bytes === undefined) {
    return "";
  }
  return java.util.Base64.getEncoder().encodeToString(bytes);
}

function spop_appendQueryParam(url, name, value) {
  if (spop_isBlank(value)) {
    return url;
  }
  var separator = url.indexOf("?") >= 0 ? "&" : "?";
  return url + separator + String(name) + "=" + spop_encodeQueryValue(value);
}

function spop_buildListEndpoint(siteBaseUrl, listId, listName) {
  if (!spop_isBlank(listId)) {
    return siteBaseUrl + "/_api/web/lists(guid'" + spop_encodeODataStringLiteral(String(listId)) + "')";
  }
  return siteBaseUrl + "/_api/web/lists/GetByTitle('" + spop_escapeListTitle(listName) + "')";
}

function spop_getContextDigest(siteBaseUrl, auth) {
  var emptyBytes = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 0);
  var result = spop_httpRequest(
    "POST",
    siteBaseUrl + "/_api/contextinfo",
    auth,
    emptyBytes,
    "application/json;odata=verbose;charset=utf-8",
    "application/json;odata=verbose",
    null,
    false
  );

  spop_expectSuccess(result, "contextinfo");

  var payload = spop_parseVerbose(JSON.parse(result.body));
  var info = payload.GetContextWebInformation;
  if (info === null || info === undefined || spop_isBlank(info.FormDigestValue)) {
    throw new java.lang.RuntimeException("contextinfo response does not contain FormDigestValue");
  }
  return {
    formDigestValue: String(info.FormDigestValue),
    formDigestTimeoutSeconds: info.FormDigestTimeoutSeconds === undefined ? 0 : info.FormDigestTimeoutSeconds,
    libraryVersion: spop_safeString(info.LibraryVersion),
    siteFullUrl: spop_safeString(info.SiteFullUrl),
    webFullUrl: spop_safeString(info.WebFullUrl)
  };
}

function spop_loadListMetadata(siteBaseUrl, auth, listId, listName) {
  var endpoint = spop_buildListEndpoint(siteBaseUrl, listId, listName);
  endpoint = spop_appendQueryParam(endpoint, "$select", "Id,Title,BaseTemplate,Hidden,ItemCount,DefaultViewUrl,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl");
  endpoint = spop_appendQueryParam(endpoint, "$expand", "RootFolder");

  var payload = spop_httpRequestJson("GET", endpoint, auth, null, null);
  var listObj = spop_parseVerbose(payload);
  return {
    endpoint: spop_buildListEndpoint(siteBaseUrl, listObj.Id, listObj.Title),
    list: listObj
  };
}

function spop_normalizeServerRelativeUrl(value) {
  return spop_encodePathKeepSlash(spop_ensureLeadingSlash(value));
}

function spop_buildGetFolderEndpoint(siteBaseUrl, folderServerRelativeUrl) {
  return siteBaseUrl + "/_api/web/GetFolderByServerRelativeUrl('" + spop_normalizeServerRelativeUrl(folderServerRelativeUrl) + "')";
}

function spop_buildGetFileEndpoint(siteBaseUrl, fileServerRelativeUrl) {
  return siteBaseUrl + "/_api/web/GetFileByServerRelativeUrl('" + spop_normalizeServerRelativeUrl(fileServerRelativeUrl) + "')";
}

function spop_payloadErrorMessage(payload) {
  try {
    var verbose = spop_parseVerbose(payload);
    if (verbose !== null && verbose !== undefined && verbose.error !== undefined && verbose.error !== null) {
      if (verbose.error.message !== undefined && verbose.error.message !== null) {
        var messageObject = verbose.error.message;
        if (messageObject.value !== undefined && messageObject.value !== null) {
          return String(messageObject.value);
        }
        return String(messageObject);
      }
    }
  } catch (ignorePayloadError) {
  }
  return "";
}
