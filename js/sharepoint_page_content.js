// The content of a modern page of a SharePoint site (its home page by default), read by Microsoft Graph and turned into
// a neutral structure any application can rebuild: sections, columns, blocks (text, image, links, list, news, activity,
// divider), each column also given as one clean HTML fragment, and the logo of the site. Needs
// js/sharepoint_graph_common.js (included first).
//
// GET /sites/{site}/pages/microsoft.graph.sitePage (the pages, the home one chosen), then
// GET /sites/{site}/pages/{page}/microsoft.graph.sitePage?$expand=canvasLayout (its sections and web parts).
// The images of the page are read by Graph too (the library that holds them, by path). The logo of the site only by the
// REST API of its host (GET <site>/_api/siteiconmanager/getsitelogo), with a token of that host when one is given.

// The web parts this library turns into blocks (the ids of their types, as SharePoint gives them).
var GS_PAGE_WEBPARTS = {
  "d1d91016-032f-456d-98a4-721247c305e8": "image",
  "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd": "quicklinks",
  "f92bf067-bc19-489e-a556-7fe95f508720": "list",
  "8c88f208-6c77-4bdb-86a0-0c47b4316588": "news",
  "eb95c819-ab8f-4689-bd03-0c2d65d47b1f": "activity",
  "2161a1c6-db61-4731-b97c-3cdb303f7cbb": "divider",
  "8654b779-4886-46d4-8ffb-b5ed960ee986": "spacer"
};
var GS_PAGE_MAX_IMAGES = 20;
var GS_PAGE_MAX_IMAGE_BYTES = 1000000;
var GS_PAGE_MAX_HTML = 200000;
var GS_PAGE_MAX_NEWS = 6;
var GS_PAGE_MAX_ACTIVITY = 30;

// ---- HTML kept: a list of tags and attributes, everything else dropped ------------------------------------------------

var GS_HTML_TAGS = ["p", "br", "h1", "h2", "h3", "h4", "h5", "h6", "strong", "b", "em", "i", "u", "s", "strike", "sub", "sup",
  "span", "div", "blockquote", "pre", "code", "ul", "ol", "li", "a", "table", "thead", "tbody", "tfoot", "tr", "th", "td",
  "hr", "img", "figure", "figcaption"];
var GS_HTML_VOID = ["br", "hr", "img"];
// dropped with what they hold
var GS_HTML_DROPPED = ["script", "style", "template", "noscript", "iframe", "object", "embed", "svg", "math", "head",
  "title", "textarea", "select", "button", "form"];
var GS_HTML_ATTRIBUTES = { a: ["href", "title"], img: ["src", "alt", "width", "height"], td: ["colspan", "rowspan"], th: ["colspan", "rowspan"] };

function gs_htmlEscape(value) {
  return gs_safeString(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// The text of an attribute: its entities decoded (enough to check an URL).
function gs_htmlDecode(value) {
  return gs_safeString(value).replace(/&#x([0-9a-fA-F]+);?/g, function (all, hex) {
    return String.fromCharCode(parseInt(hex, 16));
  }).replace(/&#([0-9]+);?/g, function (all, dec) {
    return String.fromCharCode(parseInt(dec, 10));
  }).replace(/&colon;/gi, ":").replace(/&tab;/gi, "\t").replace(/&newline;/gi, "\n").replace(/&quot;/gi, "\"")
    .replace(/&apos;/gi, "'").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">").replace(/&amp;/gi, "&");
}

// An URL a link may have: https, http or mailto, or a path of the host of the site (made absolute). "" for any other.
function gs_pageSafeUrl(value, hostUrl) {
  var url = gs_htmlDecode(value).replace(/[\u0000- \u007f-\u009f]/g, "");
  if (url === "") {
    return "";
  }
  if (/^\/[^\/\\]/.test(url) && !gs_isBlank(hostUrl)) {
    return hostUrl + url;
  }
  if (/^https?:\/\/[^\/\\]/i.test(url) || /^mailto:[^\s<>"]+$/i.test(url)) {
    return url;
  }
  return "";
}

// The source an image may have: an image inlined (data:image/png|jpeg|gif|webp;base64) or https.
function gs_pageSafeImageSource(value) {
  var src = gs_htmlDecode(value).replace(/[\u0000- \u007f-\u009f]/g, "");
  if (/^data:image\/(png|jpeg|gif|webp);base64,[A-Za-z0-9+\/=]+$/i.test(src) || /^https:\/\/[^\/\\]/i.test(src)) {
    return src;
  }
  return "";
}

// Clean HTML: the tags and attributes of the lists only (no style, class, event nor script), every other "<" escaped.
function gs_cleanHtml(htmlValue, hostUrl) {
  var html = gs_safeString(htmlValue);
  if (html.length > GS_PAGE_MAX_HTML) {
    html = html.substring(0, GS_PAGE_MAX_HTML);
  }
  // comments, then what is dropped with its content
  html = html.replace(/<!--[\s\S]*?(-->|$)/g, "").replace(/<!\[CDATA\[[\s\S]*?(\]\]>|$)/g, "");
  var i;
  for (i = 0; i < GS_HTML_DROPPED.length; i++) {
    var tag = GS_HTML_DROPPED[i];
    html = html.replace(new RegExp("<" + tag + "\\b[\\s\\S]*?(<\\/" + tag + "\\s*>|$)", "gi"), "");
  }
  var out = [];
  var open = [];
  var tokens = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)\b((?:[^>"']|"[^"]*"|'[^']*')*)>/g;
  var last = 0;
  var match;
  var text = function (value) {
    out.push(value.replace(/&(?!(#[0-9]{1,7}|#x[0-9a-fA-F]{1,6}|[a-zA-Z][a-zA-Z0-9]{1,31});)/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"));
  };
  while ((match = tokens.exec(html)) !== null) {
    text(html.substring(last, match.index));
    last = tokens.lastIndex;
    var closing = match[1] === "/";
    var name = match[2].toLowerCase();
    if (GS_HTML_TAGS.indexOf(name) === -1) {
      continue;
    }
    if (closing) {
      var at = open.lastIndexOf(name);
      if (at !== -1) {
        while (open.length > at) {
          out.push("</" + open.pop() + ">");
        }
      }
      continue;
    }
    var allowed = GS_HTML_ATTRIBUTES[name] || [];
    var attributes = "";
    var attribute = /([a-zA-Z_:][-a-zA-Z0-9_:.]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
    var found;
    var seen = {};
    while ((found = attribute.exec(match[3])) !== null) {
      var key = found[1].toLowerCase();
      if (allowed.indexOf(key) === -1 || seen[key]) {
        continue;
      }
      var raw = found[2] !== undefined ? found[2] : (found[3] !== undefined ? found[3] : (found[4] !== undefined ? found[4] : ""));
      var value = "";
      if (key === "href") {
        value = gs_pageSafeUrl(raw, hostUrl);
      } else if (key === "src") {
        value = gs_pageSafeImageSource(raw);
      } else if (key === "width" || key === "height" || key === "colspan" || key === "rowspan") {
        value = /^[0-9]{1,4}$/.test(gs_htmlDecode(raw).trim()) ? gs_htmlDecode(raw).trim() : "";
      } else {
        value = gs_htmlDecode(raw).substring(0, 500);
      }
      if (value === "") {
        continue;
      }
      seen[key] = true;
      attributes += " " + key + "=\"" + gs_htmlEscape(value) + "\"";
    }
    if (name === "img" && !seen.src) {
      continue;
    }
    if (name === "a" && seen.href) {
      attributes += " target=\"_blank\" rel=\"noopener noreferrer\"";
    }
    out.push("<" + name + attributes + ">");
    if (GS_HTML_VOID.indexOf(name) === -1) {
      open.push(name);
    }
  }
  text(html.substring(last));
  while (open.length > 0) {
    out.push("</" + open.pop() + ">");
  }
  return out.join("");
}

// ---- The values of a web part: Graph gives its dictionaries as [{key, value}] (metaDataKeyStringPair) ------------------

function gs_pageDictionary(value) {
  var dictionary = {};
  if (Array.isArray(value)) {
    var i;
    for (i = 0; i < value.length; i++) {
      if (value[i] && typeof value[i] === "object" && !gs_isBlank(value[i].key)) {
        dictionary[String(value[i].key)] = gs_safeString(value[i].value);
      }
    }
  } else if (value && typeof value === "object") {
    var key;
    for (key in value) {
      if (Object.prototype.hasOwnProperty.call(value, key) && typeof value[key] !== "object") {
        dictionary[key] = gs_safeString(value[key]);
      }
    }
  }
  return dictionary;
}

function gs_pageWebPartType(webPart) {
  var odata = gs_safeString(webPart && webPart["@odata.type"]).toLowerCase();
  if (odata.indexOf("textwebpart") !== -1 || (webPart && typeof webPart.innerHtml === "string" && gs_isBlank(webPart.webPartType))) {
    return "text";
  }
  var id = gs_safeString(webPart && webPart.webPartType).toLowerCase();
  return GS_PAGE_WEBPARTS[id] || "";
}

// ---- The site: its host, its lists and libraries (to know what a link or a list web part shows), its drives ----------

function gs_pageGraphBase(token) {
  return String(gs_getGraphClientContext(token).adapter.getBaseUrl());
}

function gs_pageSiteUrl(token, siteId) {
  return gs_pageGraphBase(token) + "/sites/" + gs_encodePathSegment(siteId);
}

// An URL compared to another: no scheme case, no trailing slash, decoded, lower case.
function gs_pageUrlKey(value) {
  var url = gs_safeString(value);
  try {
    url = decodeURIComponent(url);
  } catch (ignoreDecode) {
  }
  return url.replace(/[?#].*$/, "").replace(/\/+$/, "").toLowerCase();
}

function gs_pageSiteContext(token, siteId) {
  var site = gs_graphGetJson(gs_pageSiteUrl(token, siteId) + "?$select=id,webUrl,displayName", token);
  var webUrl = gs_safeString(site.webUrl);
  var hostMatch = /^(https:\/\/[^\/]+)/i.exec(webUrl);
  // system in $select: Graph lists the system lists too (Site Assets, which holds the images of the pages)
  var lists = gs_graphCollectPagedValues(gs_pageSiteUrl(token, siteId) + "/lists?$select=id,name,displayName,webUrl,list,system&$top=200", token, 5).values;
  var context = { siteId: gs_safeString(site.id) || siteId, webUrl: webUrl, hostUrl: hostMatch ? hostMatch[1] : "", lists: [], drives: null, defaultListId: null };
  var i;
  for (i = 0; i < lists.length; i++) {
    context.lists.push({ id: gs_safeString(lists[i].id), name: gs_safeString(lists[i].displayName) || gs_safeString(lists[i].name),
      webUrl: gs_safeString(lists[i].webUrl), key: gs_pageUrlKey(lists[i].webUrl),
      library: !!(lists[i].list && lists[i].list.template === "documentLibrary") });
  }
  return context;
}

// The list or library of the site an URL is in (the longest web address that starts it), or null.
function gs_pageListOfUrl(context, url) {
  var key = gs_pageUrlKey(url);
  var best = null;
  var i;
  for (i = 0; i < context.lists.length; i++) {
    var list = context.lists[i];
    if (list.key !== "" && (key === list.key || key.indexOf(list.key + "/") === 0) && (best === null || list.key.length > best.key.length)) {
      best = list;
    }
  }
  return best;
}

function gs_pageListById(context, id) {
  var wanted = gs_safeString(id).replace(/[{}]/g, "").toLowerCase();
  var i;
  for (i = 0; i < context.lists.length; i++) {
    if (context.lists[i].id.toLowerCase() === wanted) {
      return context.lists[i];
    }
  }
  return null;
}

// The default library of the site (the one a Documents web part shows by default).
function gs_pageDefaultLibrary(token, context) {
  if (context.defaultListId === null) {
    context.defaultListId = "";
    try {
      var drive = gs_graphGetJson(gs_pageSiteUrl(token, context.siteId) + "/drive?$select=id,webUrl", token);
      var list = gs_pageListOfUrl(context, drive.webUrl);
      context.defaultListId = list === null ? "" : list.id;
    } catch (ignoreDrive) {
      context.defaultListId = "";
    }
  }
  return gs_pageListById(context, context.defaultListId);
}

// The library of the site an URL is in, as a drive: {id, key}, or null. By the library among the lists of the site (Graph
// leaves some libraries out of /drives, Site Assets among them), its drive read once; else among /drives.
function gs_pageDriveOfUrl(token, context, url) {
  var key = gs_pageUrlKey(url);
  var list = gs_pageListOfUrl(context, url);
  if (list !== null && list.library) {
    if (context.listDrives === undefined) {
      context.listDrives = {};
    }
    if (context.listDrives[list.id] === undefined) {
      context.listDrives[list.id] = null;
      try {
        var drive = gs_graphGetJson(gs_pageSiteUrl(token, context.siteId) + "/lists/" + gs_encodePathSegment(list.id) + "/drive?$select=id,webUrl", token);
        if (!gs_isBlank(drive.id)) {
          context.listDrives[list.id] = { id: gs_safeString(drive.id), key: gs_pageUrlKey(drive.webUrl) || list.key };
        }
      } catch (ignoreListDrive) {
        context.listDrives[list.id] = null;
      }
    }
    var listDrive = context.listDrives[list.id];
    if (listDrive !== null && listDrive.key !== "" && key.indexOf(listDrive.key + "/") === 0) {
      return listDrive;
    }
  }
  var drives = gs_pageDrives(token, context);
  var found = null;
  var i;
  for (i = 0; i < drives.length; i++) {
    if (drives[i].key !== "" && key.indexOf(drives[i].key + "/") === 0 && (found === null || drives[i].key.length > found.key.length)) {
      found = drives[i];
    }
  }
  return found;
}

// The drives (libraries) of the site, read once: {id, key}.
function gs_pageDrives(token, context) {
  if (context.drives === null) {
    context.drives = [];
    var drives = gs_graphCollectPagedValues(gs_pageSiteUrl(token, context.siteId) + "/drives?$select=id,webUrl", token, 5).values;
    var i;
    for (i = 0; i < drives.length; i++) {
      context.drives.push({ id: gs_safeString(drives[i].id), key: gs_pageUrlKey(drives[i].webUrl) });
    }
  }
  return context.drives;
}

// ---- The images: a file of a library of the site, read by Graph by its path, inlined (data:) under a size -----------

function gs_pageDownloadBytes(url, token, maxBytes) {
  var response = gs_getGraphHttpClient().newCall(gs_buildNativeRequest("GET", url, token, null, null, "", "*/*", null)).execute();
  try {
    var statusCode = response.code();
    if (statusCode < 200 || statusCode >= 300) {
      throw gs_httpFailure("Image download failed", { statusCode: statusCode, body: "", retryAfter: gs_safeString(response.header("Retry-After")) });
    }
    if (response.body() === null || response.body().contentLength() > maxBytes) {
      throw gs_httpFailure("Image too large", { statusCode: 413, body: "", retryAfter: "" });
    }
    var input = response.body().byteStream();
    var output = new java.io.ByteArrayOutputStream();
    var buffer = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 65536);
    var size = 0;
    var read = input.read(buffer);
    while (read !== -1) {
      size += read;
      if (size > maxBytes) {
        throw gs_httpFailure("Image too large", { statusCode: 413, body: "", retryAfter: "" });
      }
      output.write(buffer, 0, read);
      read = input.read(buffer);
    }
    return { bytes: output.toByteArray(), contentType: gs_safeString(response.header("Content-Type")).split(";")[0].trim().toLowerCase() };
  } finally {
    response.close();
  }
}

function gs_pageBase64(bytes) {
  return String(java.util.Base64.getEncoder().encodeToString(bytes));
}

// The type of an image by its first bytes (the Content-Type of a download is not trusted): png, jpeg, gif, webp or "".
function gs_pageImageType(bytes) {
  var b = function (i) {
    return i < bytes.length ? (bytes[i] & 255) : -1;
  };
  if (b(0) === 0x89 && b(1) === 0x50 && b(2) === 0x4e && b(3) === 0x47) {
    return "png";
  }
  if (b(0) === 0xff && b(1) === 0xd8 && b(2) === 0xff) {
    return "jpeg";
  }
  if (b(0) === 0x47 && b(1) === 0x49 && b(2) === 0x46 && b(3) === 0x38) {
    return "gif";
  }
  if (b(0) === 0x52 && b(1) === 0x49 && b(2) === 0x46 && b(3) === 0x46 && b(8) === 0x57 && b(9) === 0x45 && b(10) === 0x42 && b(11) === 0x50) {
    return "webp";
  }
  return "";
}

// An image of the page: {status: embedded (src a data: URL), external (src an https URL of another host), too_large,
// unavailable, limit (too many images)}.
function gs_pageImage(token, context, source, options, state) {
  var url = gs_safeString(source);
  var absolute = /^\//.test(url) ? context.hostUrl + url : url;
  if (!/^https:\/\/[^\/\\]/i.test(absolute)) {
    return { status: "unavailable", src: "" };
  }
  var sameHost = context.hostUrl !== "" && absolute.toLowerCase().indexOf(context.hostUrl.toLowerCase() + "/") === 0;
  if (!sameHost) {
    return { status: "external", src: gs_pageSafeImageSource(absolute) };
  }
  if (!options.embedImages) {
    return { status: "unavailable", src: "" };
  }
  if (state.images >= options.maxImages) {
    return { status: "limit", src: "" };
  }
  var key = gs_pageUrlKey(absolute);
  var drive = gs_pageDriveOfUrl(token, context, absolute);
  var i;
  if (drive === null) {
    return { status: "unavailable", src: "" };
  }
  var path = key.substring(drive.key.length + 1);
  // the path as SharePoint wrote it (its case), from the decoded URL
  var decoded = absolute;
  try {
    decoded = decodeURIComponent(absolute);
  } catch (ignoreDecode) {
  }
  decoded = decoded.replace(/[?#].*$/, "");
  path = decoded.substring(decoded.length - path.length);
  var segments = path.split("/");
  var encoded = [];
  for (i = 0; i < segments.length; i++) {
    if (segments[i] === "" || segments[i] === "." || segments[i] === "..") {
      return { status: "unavailable", src: "" };
    }
    encoded.push(gs_encodePathSegment(segments[i]));
  }
  state.images++;
  try {
    var download = gs_pageDownloadBytes(gs_pageGraphBase(token) + "/drives/" + gs_encodePathSegment(drive.id) + "/root:/" + encoded.join("/") + ":/content",
      token, options.maxImageBytes);
    var type = gs_pageImageType(download.bytes);
    if (type === "") {
      return { status: "unavailable", src: "" };
    }
    return { status: "embedded", src: "data:image/" + type + ";base64," + gs_pageBase64(download.bytes) };
  } catch (e) {
    var http = gs_httpStatusOf(e);
    return { status: http.status === 413 ? "too_large" : "unavailable", src: "" };
  }
}

// The news of the site (its pages promoted as news), the newest first, count at most: {title, description, url, date,
// image: {status, src}} (the thumbnail of the post, read as the images of the page).
function gs_pageNews(token, context, options, state, count) {
  var pages = gs_graphCollectPagedValues(gs_pageSiteUrl(token, context.siteId) +
    "/pages/microsoft.graph.sitePage?$select=id,title,description,webUrl,thumbnailWebUrl,promotionKind,createdDateTime&$top=100", token, 5).values;
  var posts = [];
  var i;
  for (i = 0; i < pages.length; i++) {
    if (gs_safeString(pages[i].promotionKind).toLowerCase() === "newspost") {
      posts.push(pages[i]);
    }
  }
  posts.sort(function(a, b) {
    var left = gs_safeString(a.createdDateTime);
    var right = gs_safeString(b.createdDateTime);
    return left < right ? 1 : (left > right ? -1 : 0);
  });
  var news = [];
  for (i = 0; i < posts.length && news.length < count; i++) {
    var url = gs_pageSafeUrl(posts[i].webUrl, context.hostUrl);
    var thumbnail = gs_safeString(posts[i].thumbnailWebUrl);
    var image = thumbnail === "" ? { status: "unavailable", src: "" } : gs_pageImage(token, context, thumbnail, options, state);
    news.push({ title: gs_safeString(posts[i].title).substring(0, 500), description: gs_safeString(posts[i].description).substring(0, 1000),
      url: url, date: gs_safeString(posts[i].createdDateTime).substring(0, 40), image: { status: image.status, src: image.src } });
  }
  return news;
}

// ---- The news of the site, whole: for an application that replaces the site ------------------------------------------

// A whole number that stands for a news post (its page id, a GUID): the first 12 hexadecimal digits of the GUID, a number
// below 2^48 (an exact number of JavaScript and of a no-code database).
function gs_newsKey(pageId) {
  var hex = gs_safeString(pageId).replace(/[^0-9a-fA-F]/g, "");
  if (hex.length < 12) {
    return null;
  }
  var key = parseInt(hex.substring(0, 12), 16);
  return key > 0 ? key : null;
}

// The file of the site an image URL is (a file of one of its libraries, read by Graph): {driveId, itemId, name, size}, or
// null (another host, the default thumbnail of SharePoint under _layouts, a file not found).
function gs_newsImageFile(token, context, source) {
  var url = gs_safeString(source);
  // a preview of SharePoint (getpreview.ashx?path=...): the file it shows
  var preview = /\/_layouts\/15\/getpreview\.ashx\?(?:[^#]*&)?path=([^&#]+)/i.exec(url);
  if (preview) {
    try {
      url = decodeURIComponent(preview[1]);
    } catch (ignoreDecode) {
      return null;
    }
  }
  var absolute = /^\//.test(url) ? context.hostUrl + url : url;
  if (context.hostUrl === "" || absolute.toLowerCase().indexOf(context.hostUrl.toLowerCase() + "/") !== 0 || /\/_layouts\//i.test(absolute)) {
    return null;
  }
  var drive = gs_pageDriveOfUrl(token, context, absolute);
  if (drive === null) {
    return null;
  }
  var key = gs_pageUrlKey(absolute);
  var decoded = absolute;
  try {
    decoded = decodeURIComponent(absolute);
  } catch (ignoreDecode2) {
  }
  decoded = decoded.replace(/[?#].*$/, "");
  var path = decoded.substring(decoded.length - (key.length - drive.key.length - 1));
  var segments = path.split("/");
  var encoded = [];
  var i;
  for (i = 0; i < segments.length; i++) {
    if (segments[i] === "" || segments[i] === "." || segments[i] === "..") {
      return null;
    }
    encoded.push(gs_encodePathSegment(segments[i]));
  }
  try {
    var item = gs_graphGetJson(gs_pageGraphBase(token) + "/drives/" + gs_encodePathSegment(drive.id) + "/root:/" + encoded.join("/") + ":?$select=id,name,size,file", token);
    if (gs_isBlank(item.id) || !item.file) {
      return null;
    }
    return { driveId: drive.id, itemId: gs_safeString(item.id), name: gs_safeString(item.name).substring(0, 200), size: item.size === undefined ? null : Number(item.size) };
  } catch (ignoreItem) {
    return null;
  }
}

// The text of a news post: the HTML of its text web parts, cleaned (no image, no link kept: the post is read in another
// application).
function gs_newsContent(token, context, pageId) {
  var page = gs_graphGetJson(gs_pageSiteUrl(token, context.siteId) + "/pages/" + gs_encodePathSegment(pageId) + "/microsoft.graph.sitePage?$expand=canvasLayout", token);
  var layout = page.canvasLayout || {};
  var all = Array.isArray(layout.horizontalSections) ? layout.horizontalSections.slice() : [];
  if (layout.verticalSection && Array.isArray(layout.verticalSection.webparts)) {
    all.push({ columns: [{ webparts: layout.verticalSection.webparts }] });
  }
  var html = [];
  var s;
  for (s = 0; s < all.length; s++) {
    var columns = all[s] && Array.isArray(all[s].columns) ? all[s].columns : [];
    var c;
    for (c = 0; c < columns.length; c++) {
      var parts = Array.isArray(columns[c].webparts) ? columns[c].webparts : [];
      var w;
      for (w = 0; w < parts.length; w++) {
        if (gs_pageWebPartType(parts[w]) === "text") {
          html.push(gs_cleanHtml(parts[w].innerHtml, context.hostUrl).replace(/<a\b[^>]*>|<\/a>/gi, "").replace(/<img\b[^>]*>/gi, ""));
        }
      }
    }
  }
  return html.join("");
}

// The news of a site, the newest first, maxItems at most: {items: [{key, pageId, title, description, date, content,
// image: {driveId, itemId, name, size} or null}], count}. options: {withContent: the text and the image of each post (a
// read of Graph for each), maxItems}.
function gs_siteNews(token, siteIdValue, options) {
  var siteId = gs_require("siteId", siteIdValue);
  var context = gs_pageSiteContext(token, siteId);
  var pages = gs_graphCollectPagedValues(gs_pageSiteUrl(token, context.siteId) +
    "/pages/microsoft.graph.sitePage?$select=id,title,description,webUrl,thumbnailWebUrl,promotionKind,createdDateTime&$top=100", token, 5).values;
  var posts = [];
  var keys = {};
  var i;
  for (i = 0; i < pages.length; i++) {
    var newsKey = gs_newsKey(pages[i].id);
    // a key met twice (two GUIDs of the same first 48 bits): the first post only
    if (gs_safeString(pages[i].promotionKind).toLowerCase() === "newspost" && newsKey !== null && keys[newsKey] !== true) {
      keys[newsKey] = true;
      posts.push(pages[i]);
    }
  }
  posts.sort(function(a, b) {
    var left = gs_safeString(a.createdDateTime);
    var right = gs_safeString(b.createdDateTime);
    return left < right ? 1 : (left > right ? -1 : 0);
  });
  var items = [];
  for (i = 0; i < posts.length && items.length < options.maxItems; i++) {
    var post = posts[i];
    var item = { key: gs_newsKey(post.id), pageId: gs_safeString(post.id), title: gs_safeString(post.title).substring(0, 500),
      description: gs_safeString(post.description).substring(0, 2000), date: gs_safeString(post.createdDateTime).substring(0, 40) };
    if (options.withContent) {
      try {
        item.content = gs_newsContent(token, context, post.id).substring(0, GS_PAGE_MAX_HTML);
      } catch (ignoreContent) {
        item.content = "";
      }
      item.image = gs_isBlank(post.thumbnailWebUrl) ? null : gs_newsImageFile(token, context, post.thumbnailWebUrl);
    }
    items.push(item);
  }
  return { items: items, count: posts.length };
}

// The color of the top left corner of an image (#RRGGBB), "" when it is transparent or cannot be read: for the logo of a
// site without one of its own, its initials on the color of its theme, that color.
function gs_pageCornerColor(bytes) {
  try {
    var image = javax.imageio.ImageIO.read(new java.io.ByteArrayInputStream(bytes));
    if (image === null || image.getWidth() < 3 || image.getHeight() < 3) {
      return "";
    }
    var argb = image.getRGB(2, 2);
    if (((argb >>> 24) & 255) < 200) {
      return "";
    }
    return "#" + ((argb & 0xffffff) | 0x1000000).toString(16).substring(1).toUpperCase();
  } catch (ignoreImage) {
    return "";
  }
}

// The logo of the site, as SharePoint shows it in the header of the site (its own image, else its initials on the color
// of its theme), read by the REST API of its host with sharePointToken: {status: embedded, too_large, unavailable, src,
// color: the color of its top left corner}.
function gs_pageLogo(context, sharePointToken, maxBytes) {
  if (gs_isBlank(sharePointToken) || gs_isBlank(context.webUrl)) {
    return { status: "unavailable", src: "", color: "" };
  }
  try {
    var download = gs_pageDownloadBytes(gs_restSiteUrl(context.webUrl) + "/_api/siteiconmanager/getsitelogo?type=%271%27", sharePointToken, maxBytes);
    var type = gs_pageImageType(download.bytes);
    if (type === "") {
      return { status: "unavailable", src: "", color: "" };
    }
    return { status: "embedded", src: "data:image/" + type + ";base64," + gs_pageBase64(download.bytes), color: gs_pageCornerColor(download.bytes) };
  } catch (e) {
    var http = gs_httpStatusOf(e);
    return { status: http.status === 413 ? "too_large" : "unavailable", src: "", color: "" };
  }
}

// ---- The blocks of the web parts ---------------------------------------------------------------------------------------

function gs_pageLinkItems(context, texts, links) {
  var items = [];
  var n;
  for (n = 0; n < 100; n++) {
    var url = links["items[" + n + "].sourceItem.url"];
    if (url === undefined) {
      url = links["items[" + n + "].url"];
    }
    if (url === undefined) {
      if (texts["items[" + n + "].title"] === undefined) {
        break;
      }
      continue;
    }
    var safe = gs_pageSafeUrl(url, context.hostUrl);
    if (safe === "") {
      continue;
    }
    var list = gs_pageListOfUrl(context, safe);
    var item = { title: gs_safeString(texts["items[" + n + "].title"]).substring(0, 500) || safe, url: safe };
    if (list !== null) {
      item.listId = list.id;
    }
    items.push(item);
  }
  return items;
}

function gs_pageBlock(token, context, webPart, options, state, skipped) {
  var type = gs_pageWebPartType(webPart);
  var data = webPart && webPart.data ? webPart.data : {};
  var properties = data.properties && typeof data.properties === "object" ? data.properties : {};
  var processed = data.serverProcessedContent || {};
  var texts = gs_pageDictionary(processed.searchablePlainTexts);
  var links = gs_pageDictionary(processed.links);
  var images = gs_pageDictionary(processed.imageSources);
  if (type === "text") {
    var html = gs_cleanHtml(webPart.innerHtml, context.hostUrl);
    return html.replace(/<[^>]*>|\s|&nbsp;/g, "") === "" ? null : { type: "text", html: html };
  }
  if (type === "image") {
    var image = gs_pageImage(token, context, images.imageSource || properties.imageSource || "", options, state);
    var block = { type: "image", status: image.status, src: image.src, alt: gs_safeString(properties.altText).substring(0, 500),
      caption: gs_safeString(texts.captionText || properties.captionText).substring(0, 500),
      link: gs_pageSafeUrl(links.linkUrl || properties.linkUrl, context.hostUrl) };
    if (image.src === "") {
      skipped.push({ type: "image", title: block.alt || block.caption || gs_safeString(data.title), reason: image.status });
    }
    return block;
  }
  if (type === "quicklinks") {
    var title = gs_safeString(texts.title);
    if (title === "" && !gs_isBlank(properties.titleHTML)) {
      title = gs_cleanHtml(properties.titleHTML, context.hostUrl).replace(/<[^>]*>/g, "");
    }
    var items = gs_pageLinkItems(context, texts, links);
    // layout: the one of the web part (List, CompactCard, Grid, Button, Filmstrip...)
    return items.length === 0 ? null : { type: "links", title: gs_htmlDecode(title).substring(0, 500), layout: gs_safeString(properties.layoutId).substring(0, 50), items: items };
  }
  if (type === "list") {
    var list = null;
    if (!gs_isBlank(properties.selectedListId)) {
      list = gs_pageListById(context, properties.selectedListId);
    }
    if (list === null && !gs_isBlank(properties.selectedListUrl)) {
      list = gs_pageListOfUrl(context, gs_pageSafeUrl(properties.selectedListUrl, context.hostUrl));
    }
    if (list === null && (properties.showDefaultDocumentLibrary === true || (gs_isBlank(properties.selectedListId) && properties.isDocumentLibrary === true))) {
      list = gs_pageDefaultLibrary(token, context);
    }
    if (list === null) {
      skipped.push({ type: "list", title: gs_safeString(data.title), reason: "unavailable" });
      return null;
    }
    return { type: "list", title: list.name, listId: list.id, library: list.library, url: list.webUrl };
  }
  if (type === "news") {
    // the news of this site (the web part may show the ones of other sites too: not read)
    var newsTitle = gs_safeString(data.title).substring(0, 500);
    var newsCount = parseInt(properties.prefetchCount, 10);
    newsCount = isNaN(newsCount) || newsCount < 1 ? 4 : (newsCount > GS_PAGE_MAX_NEWS ? GS_PAGE_MAX_NEWS : newsCount);
    var news = null;
    try {
      news = gs_pageNews(token, context, options, state, newsCount);
    } catch (ignoreNews) {
      skipped.push({ type: "news", title: newsTitle, reason: "unavailable" });
      return null;
    }
    if (news.length === 0) {
      // as SharePoint shows it to the readers of the site: nothing
      skipped.push({ type: "news", title: newsTitle, reason: "empty" });
      return null;
    }
    return { type: "news", title: newsTitle, layout: gs_safeString(properties.layoutId).substring(0, 50), items: news };
  }
  if (type === "activity") {
    // the activity of the site is a feed of SharePoint that Graph does not give: its default library (the documents
    // changed most), maxItems as the web part shows them
    var activityLibrary = gs_pageDefaultLibrary(token, context);
    if (activityLibrary === null) {
      skipped.push({ type: "activity", title: gs_safeString(data.title).substring(0, 200), reason: "unavailable" });
      return null;
    }
    var maxItems = parseInt(properties.maxItems, 10);
    maxItems = isNaN(maxItems) || maxItems < 1 ? 9 : (maxItems > GS_PAGE_MAX_ACTIVITY ? GS_PAGE_MAX_ACTIVITY : maxItems);
    return { type: "activity", title: gs_safeString(data.title).substring(0, 500), maxItems: maxItems, listId: activityLibrary.id,
      library: true, url: activityLibrary.webUrl };
  }
  if (type === "divider") {
    return { type: "divider" };
  }
  if (type === "spacer") {
    return null;
  }
  skipped.push({ type: gs_safeString(webPart && webPart.webPartType) || "unknown", title: gs_safeString(data.title).substring(0, 200), reason: "not_supported" });
  return null;
}

// One HTML fragment of the blocks of a column.
function gs_pageColumnHtml(blocks) {
  var html = [];
  var i;
  for (i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    if (block.type === "text") {
      html.push(block.html);
    } else if (block.type === "image" && block.src !== "") {
      var img = "<img src=\"" + gs_htmlEscape(block.src) + "\" alt=\"" + gs_htmlEscape(block.alt) + "\">";
      if (block.link !== "") {
        img = "<a href=\"" + gs_htmlEscape(block.link) + "\" target=\"_blank\" rel=\"noopener noreferrer\">" + img + "</a>";
      }
      html.push("<p>" + img + "</p>");
      if (block.caption !== "") {
        html.push("<p><em>" + gs_htmlEscape(block.caption) + "</em></p>");
      }
    } else if (block.type === "links") {
      var items = [];
      var n;
      for (n = 0; n < block.items.length; n++) {
        items.push("<li><a href=\"" + gs_htmlEscape(block.items[n].url) + "\" target=\"_blank\" rel=\"noopener noreferrer\">" + gs_htmlEscape(block.items[n].title) + "</a></li>");
      }
      html.push((block.title !== "" ? "<h3>" + gs_htmlEscape(block.title) + "</h3>" : "") + "<ul>" + items.join("") + "</ul>");
    } else if (block.type === "news") {
      var posts = [];
      var k;
      for (k = 0; k < block.items.length; k++) {
        var post = block.items[k];
        var label = gs_htmlEscape(post.title) + (post.description !== "" ? " - " + gs_htmlEscape(post.description) : "");
        posts.push("<li>" + (post.url !== "" ? "<a href=\"" + gs_htmlEscape(post.url) + "\" target=\"_blank\" rel=\"noopener noreferrer\">" + label + "</a>" : label) + "</li>");
      }
      html.push((block.title !== "" ? "<h3>" + gs_htmlEscape(block.title) + "</h3>" : "") + "<ul>" + posts.join("") + "</ul>");
    } else if (block.type === "list" || block.type === "activity") {
      html.push("<h3>" + gs_htmlEscape(block.title) + "</h3>");
    } else if (block.type === "divider") {
      html.push("<hr>");
    }
  }
  return html.join("");
}

// The pages of the site: the one of pageId, else of pageName, else the home page (layout home, else Home.aspx).
function gs_pageChoose(token, siteId, pageIdValue, pageNameValue) {
  var pages = gs_graphCollectPagedValues(gs_pageSiteUrl(token, siteId) + "/pages/microsoft.graph.sitePage?$select=id,name,title,pageLayout,webUrl", token, 5).values;
  var wantedId = gs_safeString(pageIdValue).toLowerCase();
  var wantedName = gs_safeString(pageNameValue).toLowerCase();
  var home = null;
  var named = null;
  var i;
  for (i = 0; i < pages.length; i++) {
    var page = pages[i];
    if (wantedId !== "" && gs_safeString(page.id).toLowerCase() === wantedId) {
      return page;
    }
    if (wantedName !== "" && gs_safeString(page.name).toLowerCase() === wantedName) {
      return page;
    }
    if (home === null && gs_safeString(page.pageLayout).toLowerCase() === "home") {
      home = page;
    }
    if (named === null && gs_safeString(page.name).toLowerCase() === "home.aspx") {
      named = page;
    }
  }
  if (wantedId !== "" || wantedName !== "") {
    return null;
  }
  return home !== null ? home : named;
}

// The content of a page: {page, sections: [{columns: [{width, blocks, html}]}], skipped, imagesEmbedded, logo}.
// options: {pageId, pageName, embedImages, maxImages, maxImageBytes, sharePointToken (optional: a token of the host of
// the site, for its logo)}.
function gs_pageContent(token, siteIdValue, options) {
  var siteId = gs_require("siteId", siteIdValue);
  var context = gs_pageSiteContext(token, siteId);
  var chosen = gs_pageChoose(token, context.siteId, options.pageId, options.pageName);
  if (chosen === null) {
    throw gs_httpFailure("Page not found", { statusCode: 404, body: "", retryAfter: "" });
  }
  var page = gs_graphGetJson(gs_pageSiteUrl(token, context.siteId) + "/pages/" + gs_encodePathSegment(chosen.id) + "/microsoft.graph.sitePage?$expand=canvasLayout", token);
  var layout = page.canvasLayout || {};
  var state = { images: 0 };
  var skipped = [];
  var sections = [];
  var horizontal = Array.isArray(layout.horizontalSections) ? layout.horizontalSections : [];
  var all = horizontal.slice();
  if (layout.verticalSection && Array.isArray(layout.verticalSection.webparts)) {
    all.push({ columns: [{ width: 12, webparts: layout.verticalSection.webparts }] });
  }
  var s;
  for (s = 0; s < all.length; s++) {
    var columns = [];
    var source = all[s] && Array.isArray(all[s].columns) ? all[s].columns : [];
    var c;
    for (c = 0; c < source.length; c++) {
      var width = parseInt(source[c].width, 10);
      var blocks = [];
      var parts = Array.isArray(source[c].webparts) ? source[c].webparts : [];
      var w;
      for (w = 0; w < parts.length; w++) {
        var block = gs_pageBlock(token, context, parts[w], options, state, skipped);
        if (block !== null) {
          blocks.push(block);
        }
      }
      columns.push({ width: width > 0 && width <= 12 ? width : 12, blocks: blocks, html: gs_pageColumnHtml(blocks) });
    }
    sections.push({ layout: gs_safeString(all[s] && all[s].layout), columns: columns });
  }
  var embedded = 0;
  for (s = 0; s < sections.length; s++) {
    var cols = sections[s].columns;
    var k;
    for (k = 0; k < cols.length; k++) {
      var b;
      for (b = 0; b < cols[k].blocks.length; b++) {
        if (cols[k].blocks[b].type === "image" && cols[k].blocks[b].status === "embedded") {
          embedded++;
        }
      }
    }
  }
  return {
    siteId: context.siteId,
    page: { id: gs_safeString(page.id), name: gs_safeString(page.name), title: gs_safeString(page.title), webUrl: gs_safeString(page.webUrl),
      pageLayout: gs_safeString(page.pageLayout) },
    sections: sections,
    skipped: skipped,
    imagesEmbedded: embedded,
    logo: gs_pageLogo(context, options.sharePointToken, options.maxImageBytes)
  };
}
