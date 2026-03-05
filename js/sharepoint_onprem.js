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
  if (!spop_isBlank(siteBaseUrl)) {
    return spop_trimTrailingSlash(siteBaseUrl);
  }

  var host = spop_require("siteHostname", siteHostname).trim();
  var scheme = spop_defaultString(protocol, "https").trim().toLowerCase();
  var path = spop_normalizeSitePath(spop_defaultString(sitePath, "/"));
  if (path === "/") {
    return scheme + "://" + host;
  }
  return scheme + "://" + host + path;
}

function spop_buildAuth(accessToken, username, password, cookieHeader) {
  var auth = {
    mode: "onprem_anonymous",
    headers: {}
  };

  if (!spop_isBlank(accessToken)) {
    auth.mode = "onprem_bearer";
    auth.headers.Authorization = "Bearer " + String(accessToken);
  } else if (!spop_isBlank(username) && !spop_isBlank(password)) {
    auth.mode = "onprem_basic";
    var raw = String(username) + ":" + String(password);
    var encoded = java.util.Base64.getEncoder().encodeToString(new java.lang.String(raw).getBytes("UTF-8"));
    auth.headers.Authorization = "Basic " + String(encoded);
  }

  if (!spop_isBlank(cookieHeader)) {
    auth.headers.Cookie = String(cookieHeader);
    if (auth.mode === "onprem_anonymous") {
      auth.mode = "onprem_cookie";
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
  try {
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
      etag: httpInfoElement === null ? "" : spop_findResponseHeaderValue(httpInfoElement, "ETag")
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
  }
}

function spop_httpRequest(method, endpoint, auth, payloadBytes, contentType, accept, headers, binaryResponse) {
  return spop_httpRequestViaConnector(method, endpoint, auth, payloadBytes, contentType, accept, headers, binaryResponse);
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
