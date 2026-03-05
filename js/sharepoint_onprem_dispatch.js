// On-prem operation handlers executed directly from routed public sequences.
// No internal OnPrem* sequence call is performed.

include("js/sharepoint_onprem.js");

function sppr_onprem_get(parametersJS, name, defaultValue) {
  if (parametersJS !== null && parametersJS !== undefined && parametersJS[name] !== undefined && parametersJS[name] !== null) {
    return parametersJS[name];
  }
  return defaultValue;
}

function sppr_onPremResolveSite(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var includeRawSite = sppr_onprem_get(parametersJS, "includeRawSite", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_resolve_site";
  var requiredPermission = "SharePoint on-prem web read";
  var tokenMode = "onprem_unknown";

  try {
    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var endpoint = targetSiteBaseUrl + "/_api/web";
    endpoint = spop_appendQueryParam(endpoint, "$select", "Id,Title,Url,ServerRelativeUrl,Created,Language");

    var payload = spop_httpRequestJson("GET", endpoint, auth, null, null);
    var web = spop_parseVerbose(payload);

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      siteId: spop_safeString(web.Id),
      title: spop_safeString(web.Title),
      url: spop_safeString(web.Url),
      serverRelativeUrl: spop_safeString(web.ServerRelativeUrl),
      created: spop_safeString(web.Created),
      language: web.Language === undefined || web.Language === null ? "" : String(web.Language),
      rawSite: spop_toBoolean(includeRawSite, false) ? web : null
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

function sppr_onPremResolveList(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var includeRawList = sppr_onprem_get(parametersJS, "includeRawList", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_resolve_list";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  try {
    if (spop_isBlank(listId) && spop_isBlank(listName)) {
      throw new java.lang.IllegalArgumentException("Provide listId or listName");
    }

    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var metadata = spop_loadListMetadata(targetSiteBaseUrl, auth, listId, listName);
    var listObject = metadata.list;
    var rootFolder = listObject.RootFolder === undefined || listObject.RootFolder === null ? {} : listObject.RootFolder;

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      listId: spop_safeString(listObject.Id),
      listTitle: spop_safeString(listObject.Title),
      baseTemplate: listObject.BaseTemplate === undefined || listObject.BaseTemplate === null ? "" : String(listObject.BaseTemplate),
      hidden: listObject.Hidden === true,
      itemCount: listObject.ItemCount === undefined || listObject.ItemCount === null ? 0 : listObject.ItemCount,
      defaultViewUrl: spop_safeString(listObject.DefaultViewUrl),
      listItemEntityTypeFullName: spop_safeString(listObject.ListItemEntityTypeFullName),
      rootFolderServerRelativeUrl: spop_safeString(rootFolder.ServerRelativeUrl),
      rawList: spop_toBoolean(includeRawList, false) ? listObject : null
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

function sppr_onPremListSiteLists(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var top = sppr_onprem_get(parametersJS, "top", "");
  var maxPages = sppr_onprem_get(parametersJS, "maxPages", "");
  var orderBy = sppr_onprem_get(parametersJS, "orderBy", "");
  var filter = sppr_onprem_get(parametersJS, "filter", "");
  var search = sppr_onprem_get(parametersJS, "search", "");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var includeRawList = sppr_onprem_get(parametersJS, "includeRawList", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_list_site_lists";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  var spop_splitCsv = function(csvValue) {
    if (spop_isBlank(csvValue)) {
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
  };

  var spop_toLowerSafe = function(value) {
    return spop_safeString(value).toLowerCase();
  };

  var spop_extractHost = function(urlValue, fallbackHost) {
    try {
      if (!spop_isBlank(urlValue)) {
        return String(new java.net.URL(String(urlValue)).getHost()).toLowerCase();
      }
    } catch (ignoreExtractHost) {
    }
    return spop_toLowerSafe(fallbackHost);
  };

  try {
    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var pageTop = spop_parseInteger("top", top, 200);
    if (pageTop < 1) {
      pageTop = 1;
    }
    if (pageTop > 5000) {
      pageTop = 5000;
    }
    var maxPageCount = spop_parseInteger("maxPages", maxPages, 20);
    if (maxPageCount < 1) {
      maxPageCount = 1;
    }
    if (maxPageCount > 100) {
      maxPageCount = 100;
    }

    var requestedOrderBy = spop_safeString(orderBy);
    var requestedFilter = spop_safeString(filter);
    var requestedSearch = spop_toLowerSafe(search).trim();
    var includeRaw = spop_toBoolean(includeRawList, false);

    var fields = spop_splitCsv(selectFields);
    if (fields.length === 0) {
      fields = [
        "Id",
        "Title",
        "Description",
        "EntityTypeName",
        "BaseTemplate",
        "Hidden",
        "ItemCount",
        "DefaultViewUrl",
        "ListItemEntityTypeFullName",
        "Created",
        "LastItemModifiedDate",
        "RootFolder/ServerRelativeUrl"
      ];
    }

    var endpoint = targetSiteBaseUrl + "/_api/web/lists";
    endpoint = spop_appendQueryParam(endpoint, "$top", String(pageTop));
    endpoint = spop_appendQueryParam(endpoint, "$select", fields.join(","));
    endpoint = spop_appendQueryParam(endpoint, "$expand", "RootFolder");
    endpoint = spop_appendQueryParam(endpoint, "$filter", requestedFilter);
    endpoint = spop_appendQueryParam(endpoint, "$orderby", requestedOrderBy);

    var siteMetaEndpoint = targetSiteBaseUrl + "/_api/web";
    siteMetaEndpoint = spop_appendQueryParam(siteMetaEndpoint, "$select", "Id,Url,ServerRelativeUrl");
    var siteMetaPayload = spop_httpRequestJson("GET", siteMetaEndpoint, auth, null, null);
    var siteMeta = spop_parseVerbose(siteMetaPayload);

    var items = [];
    var pageCount = 0;
    var nextLink = endpoint;
    while (!spop_isBlank(nextLink) && pageCount < maxPageCount) {
      pageCount++;
      var payload = spop_httpRequestJson("GET", nextLink, auth, null, null);
      var root = spop_parseVerbose(payload);
      var rawLists = spop_extractArray(root);
      var i;
      for (i = 0; i < rawLists.length; i++) {
        var rawList = rawLists[i];
        var listTitle = spop_safeString(rawList.Title);
        var listDescription = spop_safeString(rawList.Description);
        if (!spop_isBlank(requestedSearch)) {
          var haystack = (listTitle + " " + listDescription + " " + spop_safeString(rawList.EntityTypeName)).toLowerCase();
          if (haystack.indexOf(requestedSearch) < 0) {
            continue;
          }
        }
        var rootFolder = rawList.RootFolder === undefined || rawList.RootFolder === null ? {} : rawList.RootFolder;
        var listNameValue = spop_safeString(rawList.EntityTypeName);
        if (spop_isBlank(listNameValue)) {
          listNameValue = listTitle;
        }
        var mapped = {
          listId: spop_safeString(rawList.Id),
          listName: listNameValue,
          listDisplayName: listTitle,
          description: listDescription,
          webUrl: spop_safeString(rawList.DefaultViewUrl),
          listInfo: {
            baseTemplate: rawList.BaseTemplate === undefined || rawList.BaseTemplate === null ? "" : String(rawList.BaseTemplate),
            hidden: rawList.Hidden === true,
            itemCount: rawList.ItemCount === undefined || rawList.ItemCount === null ? 0 : rawList.ItemCount,
            listItemEntityTypeFullName: spop_safeString(rawList.ListItemEntityTypeFullName),
            rootFolderServerRelativeUrl: spop_safeString(rootFolder.ServerRelativeUrl)
          },
          sharepointIds: {
            listId: spop_safeString(rawList.Id),
            siteId: spop_safeString(siteMeta.Id)
          },
          createdDateTime: spop_safeString(rawList.Created),
          lastModifiedDateTime: spop_safeString(rawList.LastItemModifiedDate)
        };
        if (includeRaw) {
          mapped.raw = rawList;
        }
        items.push(mapped);
      }
      nextLink = spop_safeString(root.__next);
    }

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      siteId: spop_safeString(siteMeta.Id),
      siteHostname: spop_extractHost(siteMeta.Url, siteHostname),
      sitePath: spop_safeString(siteMeta.ServerRelativeUrl),
      query: {
        top: pageTop,
        maxPages: maxPageCount,
        filter: requestedFilter,
        orderBy: requestedOrderBy,
        search: spop_safeString(search),
        selectFields: fields,
        includeRawList: includeRaw
      },
      pageCount: pageCount,
      itemCount: items.length,
      nextLink: nextLink,
      items: items
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

function sppr_onPremListSiteDrives(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var top = sppr_onprem_get(parametersJS, "top", "");
  var maxPages = sppr_onprem_get(parametersJS, "maxPages", "");
  var orderBy = sppr_onprem_get(parametersJS, "orderBy", "");
  var filter = sppr_onprem_get(parametersJS, "filter", "");
  var search = sppr_onprem_get(parametersJS, "search", "");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var includeRawDrive = sppr_onprem_get(parametersJS, "includeRawDrive", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_list_site_drives";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  var spop_splitCsv = function(csvValue) {
    if (spop_isBlank(csvValue)) {
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
  };

  var spop_toLowerSafe = function(value) {
    return spop_safeString(value).toLowerCase();
  };

  var spop_extractHost = function(urlValue, fallbackHost) {
    try {
      if (!spop_isBlank(urlValue)) {
        return String(new java.net.URL(String(urlValue)).getHost()).toLowerCase();
      }
    } catch (ignoreExtractHost) {
    }
    return spop_toLowerSafe(fallbackHost);
  };

  try {
    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var pageTop = spop_parseInteger("top", top, 200);
    if (pageTop < 1) {
      pageTop = 1;
    }
    if (pageTop > 5000) {
      pageTop = 5000;
    }
    var maxPageCount = spop_parseInteger("maxPages", maxPages, 20);
    if (maxPageCount < 1) {
      maxPageCount = 1;
    }
    if (maxPageCount > 100) {
      maxPageCount = 100;
    }

    var requestedOrderBy = spop_safeString(orderBy);
    var requestedFilter = spop_safeString(filter);
    var requestedSearch = spop_toLowerSafe(search).trim();
    var includeRaw = spop_toBoolean(includeRawDrive, false);

    var fields = spop_splitCsv(selectFields);
    if (fields.length === 0) {
      fields = [
        "Id",
        "Title",
        "BaseTemplate",
        "Hidden",
        "ItemCount",
        "DefaultViewUrl",
        "Created",
        "LastItemModifiedDate",
        "RootFolder/ServerRelativeUrl"
      ];
    }

    var composedFilter = "BaseTemplate eq 101 and Hidden eq false";
    if (!spop_isBlank(requestedFilter)) {
      composedFilter = "(" + composedFilter + ") and (" + requestedFilter + ")";
    }

    var endpoint = targetSiteBaseUrl + "/_api/web/lists";
    endpoint = spop_appendQueryParam(endpoint, "$top", String(pageTop));
    endpoint = spop_appendQueryParam(endpoint, "$select", fields.join(","));
    endpoint = spop_appendQueryParam(endpoint, "$expand", "RootFolder");
    endpoint = spop_appendQueryParam(endpoint, "$filter", composedFilter);
    endpoint = spop_appendQueryParam(endpoint, "$orderby", requestedOrderBy);

    var siteMetaEndpoint = targetSiteBaseUrl + "/_api/web";
    siteMetaEndpoint = spop_appendQueryParam(siteMetaEndpoint, "$select", "Id,Url,ServerRelativeUrl");
    var siteMetaPayload = spop_httpRequestJson("GET", siteMetaEndpoint, auth, null, null);
    var siteMeta = spop_parseVerbose(siteMetaPayload);

    var items = [];
    var pageCount = 0;
    var nextLink = endpoint;
    while (!spop_isBlank(nextLink) && pageCount < maxPageCount) {
      pageCount++;
      var payload = spop_httpRequestJson("GET", nextLink, auth, null, null);
      var root = spop_parseVerbose(payload);
      var rawLists = spop_extractArray(root);
      var i;
      for (i = 0; i < rawLists.length; i++) {
        var rawList = rawLists[i];
        var driveName = spop_safeString(rawList.Title);
        if (!spop_isBlank(requestedSearch) && driveName.toLowerCase().indexOf(requestedSearch) < 0) {
          continue;
        }
        var rootFolder = rawList.RootFolder === undefined || rawList.RootFolder === null ? {} : rawList.RootFolder;
        var mapped = {
          driveId: spop_safeString(rawList.Id),
          driveName: driveName,
          driveWebUrl: spop_safeString(rawList.DefaultViewUrl),
          driveType: "documentLibrary",
          createdDateTime: spop_safeString(rawList.Created),
          lastModifiedDateTime: spop_safeString(rawList.LastItemModifiedDate),
          sharepointIds: {
            listId: spop_safeString(rawList.Id),
            siteId: spop_safeString(siteMeta.Id)
          },
          quota: {},
          owner: {},
          library: {
            baseTemplate: rawList.BaseTemplate === undefined || rawList.BaseTemplate === null ? "" : String(rawList.BaseTemplate),
            hidden: rawList.Hidden === true,
            itemCount: rawList.ItemCount === undefined || rawList.ItemCount === null ? 0 : rawList.ItemCount,
            rootFolderServerRelativeUrl: spop_safeString(rootFolder.ServerRelativeUrl)
          }
        };
        if (includeRaw) {
          mapped.raw = rawList;
        }
        items.push(mapped);
      }
      nextLink = spop_safeString(root.__next);
    }

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      siteId: spop_safeString(siteMeta.Id),
      siteHostname: spop_extractHost(siteMeta.Url, siteHostname),
      sitePath: spop_safeString(siteMeta.ServerRelativeUrl),
      query: {
        top: pageTop,
        maxPages: maxPageCount,
        filter: composedFilter,
        orderBy: requestedOrderBy,
        search: spop_safeString(search),
        selectFields: fields,
        includeRawDrive: includeRaw
      },
      pageCount: pageCount,
      itemCount: items.length,
      nextLink: nextLink,
      items: items
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

function sppr_onPremListGetItems(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var top = sppr_onprem_get(parametersJS, "top", "");
  var maxPages = sppr_onprem_get(parametersJS, "maxPages", "");
  var orderByExpression = sppr_onprem_get(parametersJS, "orderByExpression", "");
  var filterExpression = sppr_onprem_get(parametersJS, "filterExpression", "");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var expandFieldValuesAsText = sppr_onprem_get(parametersJS, "expandFieldValuesAsText", "false");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_list_get_items";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  try {
    if (spop_isBlank(listId) && spop_isBlank(listName)) {
      throw new java.lang.IllegalArgumentException("Provide listId or listName");
    }

    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var metadata = spop_loadListMetadata(targetSiteBaseUrl, auth, listId, listName);
    var listObject = metadata.list;

    var endpoint = metadata.endpoint + "/items";
    endpoint = spop_appendQueryParam(endpoint, "$select", selectFields);
    if (spop_toBoolean(expandFieldValuesAsText, false)) {
      endpoint = spop_appendQueryParam(endpoint, "$expand", "FieldValuesAsText");
    }
    endpoint = spop_appendQueryParam(endpoint, "$filter", filterExpression);
    endpoint = spop_appendQueryParam(endpoint, "$orderby", orderByExpression);

    var pageSize = spop_parseInteger("top", top, 100);
    if (pageSize > 0) {
      endpoint = spop_appendQueryParam(endpoint, "$top", String(pageSize));
    }

    var maxPageCount = spop_parseInteger("maxPages", maxPages, 1);
    if (maxPageCount < 1) {
      maxPageCount = 1;
    }

    var includeRaw = spop_toBoolean(includeRawItem, false);
    var pageCount = 0;
    var totalCount = 0;
    var nextLink = "";
    var currentUrl = endpoint;
    var items = [];

    while (!spop_isBlank(currentUrl) && pageCount < maxPageCount) {
      var payload = spop_httpRequestJson("GET", currentUrl, auth, null, null);
      var root = spop_parseVerbose(payload);
      var pageItems = spop_extractArray(root);
      var i;
      for (i = 0; i < pageItems.length; i++) {
        var rawItem = pageItems[i];
        var mapped = {
          itemId: spop_safeString(rawItem.Id),
          id: spop_safeString(rawItem.Id),
          title: spop_safeString(rawItem.Title),
          modified: spop_safeString(rawItem.Modified),
          created: spop_safeString(rawItem.Created),
          authorId: spop_safeString(rawItem.AuthorId),
          editorId: spop_safeString(rawItem.EditorId),
          fields: rawItem
        };
        if (rawItem.FieldValuesAsText !== undefined && rawItem.FieldValuesAsText !== null) {
          mapped.fieldValuesAsText = rawItem.FieldValuesAsText;
        }
        if (includeRaw) {
          mapped.raw = rawItem;
        }
        items.push(mapped);
      }
      totalCount += pageItems.length;
      pageCount += 1;
      nextLink = spop_safeString(root.__next);
      currentUrl = nextLink;
    }

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      listId: spop_safeString(listObject.Id),
      listTitle: spop_safeString(listObject.Title),
      items: items,
      count: totalCount,
      pageCount: pageCount,
      nextLink: nextLink
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

function sppr_onPremGetListItem(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var itemId = sppr_onprem_get(parametersJS, "itemId", "");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var expandFieldValuesAsText = sppr_onprem_get(parametersJS, "expandFieldValuesAsText", "false");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_get_list_item";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  try {
    if (spop_isBlank(listId) && spop_isBlank(listName)) {
      throw new java.lang.IllegalArgumentException("Provide listId or listName");
    }
    var targetItemId = spop_require("itemId", itemId);

    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var metadata = spop_loadListMetadata(targetSiteBaseUrl, auth, listId, listName);
    var listObject = metadata.list;

    var endpoint = metadata.endpoint + "/items(" + targetItemId + ")";
    endpoint = spop_appendQueryParam(endpoint, "$select", selectFields);
    if (spop_toBoolean(expandFieldValuesAsText, false)) {
      endpoint = spop_appendQueryParam(endpoint, "$expand", "FieldValuesAsText");
    }

    var payload = spop_httpRequestJson("GET", endpoint, auth, null, null);
    var itemObject = spop_parseVerbose(payload);

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      listId: spop_safeString(listObject.Id),
      listTitle: spop_safeString(listObject.Title),
      item: {
        itemId: spop_safeString(itemObject.Id),
        id: spop_safeString(itemObject.Id),
        title: spop_safeString(itemObject.Title),
        created: spop_safeString(itemObject.Created),
        modified: spop_safeString(itemObject.Modified),
        fields: itemObject,
        fieldValuesAsText: itemObject.FieldValuesAsText === undefined || itemObject.FieldValuesAsText === null ? {} : itemObject.FieldValuesAsText,
        raw: spop_toBoolean(includeRawItem, false) ? itemObject : null
      }
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

function sppr_onPremCreateListItem(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var fieldsJson = sppr_onprem_get(parametersJS, "fieldsJson", "");
  var returnCreatedItem = sppr_onprem_get(parametersJS, "returnCreatedItem", "true");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_create_list_item";
  var requiredPermission = "SharePoint on-prem list write";
  var tokenMode = "onprem_unknown";

  try {
    if (spop_isBlank(listId) && spop_isBlank(listName)) {
      throw new java.lang.IllegalArgumentException("Provide listId or listName");
    }

    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var metadata = spop_loadListMetadata(targetSiteBaseUrl, auth, listId, listName);
    var listObject = metadata.list;
    var entityType = spop_safeString(listObject.ListItemEntityTypeFullName);
    if (spop_isBlank(entityType)) {
      throw new java.lang.RuntimeException("ListItemEntityTypeFullName is empty for list " + spop_safeString(listObject.Title));
    }

    var fieldsObject = {};
    if (!spop_isBlank(fieldsJson)) {
      fieldsObject = JSON.parse(String(fieldsJson));
      if (fieldsObject === null || fieldsObject === undefined || fieldsObject instanceof Array) {
        throw new java.lang.IllegalArgumentException("fieldsJson must be a JSON object");
      }
    }

    fieldsObject.__metadata = {
      type: entityType
    };

    var contextInfo = spop_getContextDigest(targetSiteBaseUrl, auth);
    var headers = {
      "X-RequestDigest": contextInfo.formDigestValue
    };
    var endpoint = metadata.endpoint + "/items";
    var createdPayload = spop_httpRequestJson("POST", endpoint, auth, fieldsObject, headers);
    var created = spop_parseVerbose(createdPayload);

    var returnItem = spop_toBoolean(returnCreatedItem, true);

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      listId: spop_safeString(listObject.Id),
      listTitle: spop_safeString(listObject.Title),
      item: returnItem ? {
        itemId: spop_safeString(created.Id),
        id: spop_safeString(created.Id),
        title: spop_safeString(created.Title),
        created: spop_safeString(created.Created),
        modified: spop_safeString(created.Modified),
        fields: created,
        raw: spop_toBoolean(includeRawItem, false) ? created : null
      } : null,
      createdItemId: spop_safeString(created.Id),
      contextInfo: {
        formDigestTimeoutSeconds: contextInfo.formDigestTimeoutSeconds,
        libraryVersion: contextInfo.libraryVersion
      }
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

function sppr_onPremUpdateListItem(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var itemId = sppr_onprem_get(parametersJS, "itemId", "");
  var fieldsJson = sppr_onprem_get(parametersJS, "fieldsJson", "");
  var ifMatch = sppr_onprem_get(parametersJS, "ifMatch", "*");
  var returnUpdatedItem = sppr_onprem_get(parametersJS, "returnUpdatedItem", "true");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var expandFieldValuesAsText = sppr_onprem_get(parametersJS, "expandFieldValuesAsText", "false");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_update_list_item";
  var requiredPermission = "SharePoint on-prem list write";
  var tokenMode = "onprem_unknown";

  try {
    if (spop_isBlank(listId) && spop_isBlank(listName)) {
      throw new java.lang.IllegalArgumentException("Provide listId or listName");
    }
    var targetItemId = spop_require("itemId", itemId);

    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var metadata = spop_loadListMetadata(targetSiteBaseUrl, auth, listId, listName);
    var listObject = metadata.list;
    var entityType = spop_safeString(listObject.ListItemEntityTypeFullName);
    if (spop_isBlank(entityType)) {
      throw new java.lang.RuntimeException("ListItemEntityTypeFullName is empty for list " + spop_safeString(listObject.Title));
    }

    var fieldsObject = {};
    if (!spop_isBlank(fieldsJson)) {
      fieldsObject = JSON.parse(String(fieldsJson));
      if (fieldsObject === null || fieldsObject === undefined || fieldsObject instanceof Array) {
        throw new java.lang.IllegalArgumentException("fieldsJson must be a JSON object");
      }
    }
    fieldsObject.__metadata = {
      type: entityType
    };

    var contextInfo = spop_getContextDigest(targetSiteBaseUrl, auth);
    var updateEndpoint = metadata.endpoint + "/items(" + targetItemId + ")";
    var updateHeaders = {
      "X-RequestDigest": contextInfo.formDigestValue,
      "IF-MATCH": spop_defaultString(ifMatch, "*"),
      "X-HTTP-Method": "MERGE"
    };

    var payloadBytes = new java.lang.String(JSON.stringify(fieldsObject)).getBytes("UTF-8");
    var updateResult = spop_httpRequest(
      "POST",
      updateEndpoint,
      auth,
      payloadBytes,
      "application/json;odata=verbose;charset=utf-8",
      "application/json;odata=verbose",
      updateHeaders,
      false
    );
    spop_expectSuccess(updateResult, "Update list item");

    var returnItem = spop_toBoolean(returnUpdatedItem, true);
    var updatedItem = null;
    if (returnItem) {
      var readEndpoint = updateEndpoint;
      readEndpoint = spop_appendQueryParam(readEndpoint, "$select", selectFields);
      if (spop_toBoolean(expandFieldValuesAsText, false)) {
        readEndpoint = spop_appendQueryParam(readEndpoint, "$expand", "FieldValuesAsText");
      }
      var readPayload = spop_httpRequestJson("GET", readEndpoint, auth, null, null);
      updatedItem = spop_parseVerbose(readPayload);
    }

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      listId: spop_safeString(listObject.Id),
      listTitle: spop_safeString(listObject.Title),
      itemId: targetItemId,
      updated: true,
      item: updatedItem === null ? null : {
        itemId: spop_safeString(updatedItem.Id),
        id: spop_safeString(updatedItem.Id),
        title: spop_safeString(updatedItem.Title),
        created: spop_safeString(updatedItem.Created),
        modified: spop_safeString(updatedItem.Modified),
        fields: updatedItem,
        fieldValuesAsText: updatedItem.FieldValuesAsText === undefined || updatedItem.FieldValuesAsText === null ? {} : updatedItem.FieldValuesAsText,
        raw: spop_toBoolean(includeRawItem, false) ? updatedItem : null
      },
      contextInfo: {
        formDigestTimeoutSeconds: contextInfo.formDigestTimeoutSeconds,
        libraryVersion: contextInfo.libraryVersion
      }
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

function sppr_onPremDeleteListItem(parametersJS) {
  var siteBaseUrl = sppr_onprem_get(parametersJS, "siteBaseUrl", "");
  var siteHostname = sppr_onprem_get(parametersJS, "siteHostname", "");
  var sitePath = sppr_onprem_get(parametersJS, "sitePath", "");
  var onPremProtocol = sppr_onprem_get(parametersJS, "onPremProtocol", "https");
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var itemId = sppr_onprem_get(parametersJS, "itemId", "");
  var ifMatch = sppr_onprem_get(parametersJS, "ifMatch", "*");
  var includeDeletedItemSnapshot = sppr_onprem_get(parametersJS, "includeDeletedItemSnapshot", "false");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");
  var accessToken = sppr_onprem_get(parametersJS, "accessToken", "");
  var onPremUsername = sppr_onprem_get(parametersJS, "onPremUsername", "");
  var onPremPassword = sppr_onprem_get(parametersJS, "onPremPassword", "");
  var cookieHeader = sppr_onprem_get(parametersJS, "cookieHeader", "");

  var graphOperation = "onprem_delete_list_item";
  var requiredPermission = "SharePoint on-prem list write";
  var tokenMode = "onprem_unknown";

  try {
    if (spop_isBlank(listId) && spop_isBlank(listName)) {
      throw new java.lang.IllegalArgumentException("Provide listId or listName");
    }
    var targetItemId = spop_require("itemId", itemId);

    var targetSiteBaseUrl = spop_buildSiteBaseUrl(siteBaseUrl, siteHostname, sitePath, onPremProtocol);
    var auth = spop_buildAuth(accessToken, onPremUsername, onPremPassword, cookieHeader);
    tokenMode = auth.mode;

    var metadata = spop_loadListMetadata(targetSiteBaseUrl, auth, listId, listName);
    var listObject = metadata.list;
    var itemEndpoint = metadata.endpoint + "/items(" + targetItemId + ")";

    var deletedSnapshot = null;
    if (spop_toBoolean(includeDeletedItemSnapshot, false)) {
      var snapshotPayload = spop_httpRequestJson("GET", itemEndpoint, auth, null, null);
      deletedSnapshot = spop_parseVerbose(snapshotPayload);
    }

    var contextInfo = spop_getContextDigest(targetSiteBaseUrl, auth);
    var deleteHeaders = {
      "X-RequestDigest": contextInfo.formDigestValue,
      "IF-MATCH": spop_defaultString(ifMatch, "*"),
      "X-HTTP-Method": "DELETE"
    };

    var deleteResult = spop_httpRequest(
      "POST",
      itemEndpoint,
      auth,
      java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 0),
      "application/json;odata=verbose;charset=utf-8",
      "application/json;odata=verbose",
      deleteHeaders,
      false
    );
    spop_expectSuccess(deleteResult, "Delete list item");

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: targetSiteBaseUrl,
      listId: spop_safeString(listObject.Id),
      listTitle: spop_safeString(listObject.Title),
      itemId: targetItemId,
      deleted: true,
      deletedItemSnapshot: deletedSnapshot === null ? null : {
        itemId: spop_safeString(deletedSnapshot.Id),
        id: spop_safeString(deletedSnapshot.Id),
        title: spop_safeString(deletedSnapshot.Title),
        created: spop_safeString(deletedSnapshot.Created),
        modified: spop_safeString(deletedSnapshot.Modified),
        fields: deletedSnapshot,
        raw: spop_toBoolean(includeRawItem, false) ? deletedSnapshot : null
      },
      contextInfo: {
        formDigestTimeoutSeconds: contextInfo.formDigestTimeoutSeconds,
        libraryVersion: contextInfo.libraryVersion
      }
    });
  } catch (e) {
    return spop_responseError(graphOperation, tokenMode, requiredPermission, e);
  }
}

