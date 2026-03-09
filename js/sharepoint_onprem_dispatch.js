// On-prem operation handlers executed directly from routed public sequences.
// No internal OnPrem* sequence call is performed.

include("js/sharepoint_onprem.js");

function sppr_onprem_get(parametersJS, name, defaultValue) {
  if (parametersJS !== null && parametersJS !== undefined && parametersJS[name] !== undefined && parametersJS[name] !== null) {
    return parametersJS[name];
  }
  return defaultValue;
}

function sppr_onprem_buildRuntimeContext(parametersJS) {
  var runtimeContext = {
    siteBaseUrl: sppr_onprem_get(parametersJS, "siteBaseUrl", ""),
    siteHostname: sppr_onprem_get(parametersJS, "siteHostname", ""),
    sitePath: sppr_onprem_get(parametersJS, "sitePath", ""),
    onPremProtocol: sppr_onprem_get(parametersJS, "onPremProtocol", ""),
    accessToken: sppr_onprem_get(parametersJS, "accessToken", ""),
    onPremUsername: sppr_onprem_get(parametersJS, "onPremUsername", ""),
    onPremPassword: sppr_onprem_get(parametersJS, "onPremPassword", ""),
    cookieHeader: sppr_onprem_get(parametersJS, "cookieHeader", "")
  };
  runtimeContext.targetSiteBaseUrl = spop_buildSiteBaseUrl(
    runtimeContext.siteBaseUrl,
    runtimeContext.siteHostname,
    runtimeContext.sitePath,
    runtimeContext.onPremProtocol
  );
  runtimeContext.auth = spop_buildAuth(
    runtimeContext.accessToken,
    runtimeContext.onPremUsername,
    runtimeContext.onPremPassword,
    runtimeContext.cookieHeader
  );
  runtimeContext.tokenMode = runtimeContext.auth.mode;
  return runtimeContext;
}

function sppr_onprem_requireListReference(listId, listName) {
  if (spop_isBlank(listId) && spop_isBlank(listName)) {
    throw new java.lang.IllegalArgumentException("Provide listId or listName");
  }
}

function sppr_onprem_resolveLibraryReference(parametersJS) {
  var driveId = sppr_onprem_get(parametersJS, "driveId", "");
  var driveName = sppr_onprem_get(parametersJS, "driveName", "");
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var effectiveListId = spop_isBlank(driveId) ? spop_safeString(listId) : spop_safeString(driveId);
  var effectiveListName = spop_isBlank(driveName) ? spop_safeString(listName) : spop_safeString(driveName);
  if (spop_isBlank(effectiveListId) && spop_isBlank(effectiveListName)) {
    throw new java.lang.IllegalArgumentException("Provide driveId, driveName, listId or listName");
  }
  return {
    listId: effectiveListId,
    listName: effectiveListName
  };
}

function sppr_onprem_parseBoundedInteger(name, value, defaultValue, minValue, maxValue) {
  var parsedValue = spop_parseInteger(name, value, defaultValue);
  if (parsedValue < minValue) {
    parsedValue = minValue;
  }
  if (maxValue !== null && maxValue !== undefined && parsedValue > maxValue) {
    parsedValue = maxValue;
  }
  return parsedValue;
}

function sppr_onprem_loadSiteMetadata(targetSiteBaseUrl, auth) {
  var siteMetaEndpoint = targetSiteBaseUrl + "/_api/web";
  siteMetaEndpoint = spop_appendQueryParam(siteMetaEndpoint, "$select", "Id,Url,ServerRelativeUrl");
  return spop_parseVerbose(spop_httpRequestJson("GET", siteMetaEndpoint, auth, null, null));
}

function sppr_onprem_parseJsonObject(name, rawJson) {
  if (spop_isBlank(rawJson)) {
    return {};
  }
  var parsedObject = JSON.parse(String(rawJson));
  if (parsedObject === null || parsedObject === undefined || parsedObject instanceof Array) {
    throw new java.lang.IllegalArgumentException(name + " must be a JSON object");
  }
  return parsedObject;
}

function sppr_onPremResolveSite(parametersJS) {
  var includeRawSite = sppr_onprem_get(parametersJS, "includeRawSite", "false");

  var graphOperation = "onprem_resolve_site";
  var requiredPermission = "SharePoint on-prem web read";
  var tokenMode = "onprem_unknown";

  try {
    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var endpoint = runtimeContext.targetSiteBaseUrl + "/_api/web";
    endpoint = spop_appendQueryParam(endpoint, "$select", "Id,Title,Url,ServerRelativeUrl,Created,Language");

    var payload = spop_httpRequestJson("GET", endpoint, runtimeContext.auth, null, null);
    var web = spop_parseVerbose(payload);

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
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
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var includeRawList = sppr_onprem_get(parametersJS, "includeRawList", "false");

  var graphOperation = "onprem_resolve_list";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  try {
    sppr_onprem_requireListReference(listId, listName);
    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var metadata = spop_loadListMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth, listId, listName);
    var listObject = metadata.list;
    var rootFolder = listObject.RootFolder === undefined || listObject.RootFolder === null ? {} : listObject.RootFolder;

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
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
  var top = sppr_onprem_get(parametersJS, "top", "");
  var maxPages = sppr_onprem_get(parametersJS, "maxPages", "");
  var orderBy = sppr_onprem_get(parametersJS, "orderBy", "");
  var filter = sppr_onprem_get(parametersJS, "filter", "");
  var search = sppr_onprem_get(parametersJS, "search", "");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var includeRawList = sppr_onprem_get(parametersJS, "includeRawList", "false");

  var graphOperation = "onprem_list_site_lists";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  try {
    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var pageTop = sppr_onprem_parseBoundedInteger("top", top, 200, 1, 5000);
    var maxPageCount = sppr_onprem_parseBoundedInteger("maxPages", maxPages, 20, 1, 100);

    var requestedOrderBy = spop_safeString(orderBy);
    var requestedFilter = spop_safeString(filter);
    var requestedSearch = spop_safeString(search).toLowerCase().trim();
    var includeRaw = spop_toBoolean(includeRawList, false);

    var fields = sppr_onprem_splitCsv(selectFields);
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

    var endpoint = runtimeContext.targetSiteBaseUrl + "/_api/web/lists";
    endpoint = spop_appendQueryParam(endpoint, "$top", String(pageTop));
    endpoint = spop_appendQueryParam(endpoint, "$select", fields.join(","));
    endpoint = spop_appendQueryParam(endpoint, "$expand", "RootFolder");
    endpoint = spop_appendQueryParam(endpoint, "$filter", requestedFilter);
    endpoint = spop_appendQueryParam(endpoint, "$orderby", requestedOrderBy);

    var siteMeta = sppr_onprem_loadSiteMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth);

    var items = [];
    var pageCount = 0;
    var nextLink = endpoint;
    while (!spop_isBlank(nextLink) && pageCount < maxPageCount) {
      pageCount++;
      var payload = spop_httpRequestJson("GET", nextLink, runtimeContext.auth, null, null);
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
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
      siteId: spop_safeString(siteMeta.Id),
      siteHostname: sppr_onprem_extractHost(siteMeta.Url, runtimeContext.siteHostname),
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
  var top = sppr_onprem_get(parametersJS, "top", "");
  var maxPages = sppr_onprem_get(parametersJS, "maxPages", "");
  var orderBy = sppr_onprem_get(parametersJS, "orderBy", "");
  var filter = sppr_onprem_get(parametersJS, "filter", "");
  var search = sppr_onprem_get(parametersJS, "search", "");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var includeRawDrive = sppr_onprem_get(parametersJS, "includeRawDrive", "false");

  var graphOperation = "onprem_list_site_drives";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  try {
    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var pageTop = sppr_onprem_parseBoundedInteger("top", top, 200, 1, 5000);
    var maxPageCount = sppr_onprem_parseBoundedInteger("maxPages", maxPages, 20, 1, 100);

    var requestedOrderBy = spop_safeString(orderBy);
    var requestedFilter = spop_safeString(filter);
    var requestedSearch = spop_safeString(search).toLowerCase().trim();
    var includeRaw = spop_toBoolean(includeRawDrive, false);

    var fields = sppr_onprem_splitCsv(selectFields);
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

    var endpoint = runtimeContext.targetSiteBaseUrl + "/_api/web/lists";
    endpoint = spop_appendQueryParam(endpoint, "$top", String(pageTop));
    endpoint = spop_appendQueryParam(endpoint, "$select", fields.join(","));
    endpoint = spop_appendQueryParam(endpoint, "$expand", "RootFolder");
    endpoint = spop_appendQueryParam(endpoint, "$filter", composedFilter);
    endpoint = spop_appendQueryParam(endpoint, "$orderby", requestedOrderBy);

    var siteMeta = sppr_onprem_loadSiteMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth);

    var items = [];
    var pageCount = 0;
    var nextLink = endpoint;
    while (!spop_isBlank(nextLink) && pageCount < maxPageCount) {
      pageCount++;
      var payload = spop_httpRequestJson("GET", nextLink, runtimeContext.auth, null, null);
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
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
      siteId: spop_safeString(siteMeta.Id),
      siteHostname: sppr_onprem_extractHost(siteMeta.Url, runtimeContext.siteHostname),
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
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var top = sppr_onprem_get(parametersJS, "top", "");
  var maxPages = sppr_onprem_get(parametersJS, "maxPages", "");
  var orderByExpression = sppr_onprem_get(parametersJS, "orderByExpression", "");
  var filterExpression = sppr_onprem_get(parametersJS, "filterExpression", "");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var expandFieldValuesAsText = sppr_onprem_get(parametersJS, "expandFieldValuesAsText", "false");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");

  var graphOperation = "onprem_list_get_items";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  try {
    sppr_onprem_requireListReference(listId, listName);
    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var metadata = spop_loadListMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth, listId, listName);
    var listObject = metadata.list;

    var effectiveSelectFields = sppr_onprem_mergeSelectFields(selectFields);
    var normalizedFilterExpression = sppr_onprem_normalizeListExpression(filterExpression);
    var normalizedOrderByExpression = sppr_onprem_normalizeListExpression(orderByExpression);

    var endpoint = metadata.endpoint + "/items";
    endpoint = spop_appendQueryParam(endpoint, "$select", effectiveSelectFields);
    if (spop_toBoolean(expandFieldValuesAsText, false)) {
      endpoint = spop_appendQueryParam(endpoint, "$expand", "FieldValuesAsText");
    }
    endpoint = spop_appendQueryParam(endpoint, "$filter", normalizedFilterExpression);
    endpoint = spop_appendQueryParam(endpoint, "$orderby", normalizedOrderByExpression);

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
      var payload = spop_httpRequestJson("GET", currentUrl, runtimeContext.auth, null, null);
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
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
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
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var itemId = sppr_onprem_get(parametersJS, "itemId", "");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var expandFieldValuesAsText = sppr_onprem_get(parametersJS, "expandFieldValuesAsText", "false");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");

  var graphOperation = "onprem_get_list_item";
  var requiredPermission = "SharePoint on-prem list read";
  var tokenMode = "onprem_unknown";

  try {
    sppr_onprem_requireListReference(listId, listName);
    var targetItemId = spop_require("itemId", itemId);

    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var metadata = spop_loadListMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth, listId, listName);
    var listObject = metadata.list;

    var effectiveSelectFields = sppr_onprem_mergeSelectFields(selectFields);
    var endpoint = metadata.endpoint + "/items(" + targetItemId + ")";
    endpoint = spop_appendQueryParam(endpoint, "$select", effectiveSelectFields);
    if (spop_toBoolean(expandFieldValuesAsText, false)) {
      endpoint = spop_appendQueryParam(endpoint, "$expand", "FieldValuesAsText");
    }

    var payload = spop_httpRequestJson("GET", endpoint, runtimeContext.auth, null, null);
    var itemObject = spop_parseVerbose(payload);

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
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
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var fieldsJson = sppr_onprem_get(parametersJS, "fieldsJson", "");
  var returnCreatedItem = sppr_onprem_get(parametersJS, "returnCreatedItem", "true");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");

  var graphOperation = "onprem_create_list_item";
  var requiredPermission = "SharePoint on-prem list write";
  var tokenMode = "onprem_unknown";

  try {
    sppr_onprem_requireListReference(listId, listName);
    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var metadata = spop_loadListMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth, listId, listName);
    var listObject = metadata.list;
    var entityType = spop_safeString(listObject.ListItemEntityTypeFullName);
    if (spop_isBlank(entityType)) {
      throw new java.lang.RuntimeException("ListItemEntityTypeFullName is empty for list " + spop_safeString(listObject.Title));
    }

    var fieldsObject = sppr_onprem_parseJsonObject("fieldsJson", fieldsJson);

    fieldsObject.__metadata = {
      type: entityType
    };

    var contextInfo = spop_getContextDigest(runtimeContext.targetSiteBaseUrl, runtimeContext.auth);
    var headers = {
      "X-RequestDigest": contextInfo.formDigestValue
    };
    var endpoint = metadata.endpoint + "/items";
    var createdPayload = spop_httpRequestJson("POST", endpoint, runtimeContext.auth, fieldsObject, headers);
    var created = spop_parseVerbose(createdPayload);

    var returnItem = spop_toBoolean(returnCreatedItem, true);

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
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
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var itemId = sppr_onprem_get(parametersJS, "itemId", "");
  var fieldsJson = sppr_onprem_get(parametersJS, "fieldsJson", "");
  var ifMatch = sppr_onprem_get(parametersJS, "ifMatch", "*");
  var returnUpdatedItem = sppr_onprem_get(parametersJS, "returnUpdatedItem", "true");
  var selectFields = sppr_onprem_get(parametersJS, "selectFields", "");
  var expandFieldValuesAsText = sppr_onprem_get(parametersJS, "expandFieldValuesAsText", "false");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");

  var graphOperation = "onprem_update_list_item";
  var requiredPermission = "SharePoint on-prem list write";
  var tokenMode = "onprem_unknown";

  try {
    sppr_onprem_requireListReference(listId, listName);
    var targetItemId = spop_require("itemId", itemId);

    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var metadata = spop_loadListMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth, listId, listName);
    var listObject = metadata.list;
    var entityType = spop_safeString(listObject.ListItemEntityTypeFullName);
    if (spop_isBlank(entityType)) {
      throw new java.lang.RuntimeException("ListItemEntityTypeFullName is empty for list " + spop_safeString(listObject.Title));
    }

    var fieldsObject = sppr_onprem_parseJsonObject("fieldsJson", fieldsJson);
    fieldsObject.__metadata = {
      type: entityType
    };

    var contextInfo = spop_getContextDigest(runtimeContext.targetSiteBaseUrl, runtimeContext.auth);
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
      runtimeContext.auth,
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
      var readPayload = spop_httpRequestJson("GET", readEndpoint, runtimeContext.auth, null, null);
      updatedItem = spop_parseVerbose(readPayload);
    }

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
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
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  var itemId = sppr_onprem_get(parametersJS, "itemId", "");
  var ifMatch = sppr_onprem_get(parametersJS, "ifMatch", "*");
  var includeDeletedItemSnapshot = sppr_onprem_get(parametersJS, "includeDeletedItemSnapshot", "false");
  var includeRawItem = sppr_onprem_get(parametersJS, "includeRawItem", "false");

  var graphOperation = "onprem_delete_list_item";
  var requiredPermission = "SharePoint on-prem list write";
  var tokenMode = "onprem_unknown";

  try {
    sppr_onprem_requireListReference(listId, listName);
    var targetItemId = spop_require("itemId", itemId);

    var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
    tokenMode = runtimeContext.tokenMode;

    var metadata = spop_loadListMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth, listId, listName);
    var listObject = metadata.list;
    var itemEndpoint = metadata.endpoint + "/items(" + targetItemId + ")";

    var deletedSnapshot = null;
    if (spop_toBoolean(includeDeletedItemSnapshot, false)) {
      var snapshotPayload = spop_httpRequestJson("GET", itemEndpoint, runtimeContext.auth, null, null);
      deletedSnapshot = spop_parseVerbose(snapshotPayload);
    }

    var contextInfo = spop_getContextDigest(runtimeContext.targetSiteBaseUrl, runtimeContext.auth);
    var deleteHeaders = {
      "X-RequestDigest": contextInfo.formDigestValue,
      "IF-MATCH": spop_defaultString(ifMatch, "*"),
      "X-HTTP-Method": "DELETE"
    };

    var deleteResult = spop_httpRequest(
      "POST",
      itemEndpoint,
      runtimeContext.auth,
      sppr_onprem_emptyBytes(),
      "application/json;odata=verbose;charset=utf-8",
      "application/json;odata=verbose",
      deleteHeaders,
      false
    );
    spop_expectSuccess(deleteResult, "Delete list item");

    return spop_responseOk(graphOperation, tokenMode, requiredPermission, {
      providerUsed: "onprem",
      siteBaseUrl: runtimeContext.targetSiteBaseUrl,
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

function sppr_onprem_splitCsv(csvValue) {
  if (spop_isBlank(csvValue)) {
    return [];
  }
  var rawParts = String(csvValue).split(",");
  var normalized = [];
  var index;
  for (index = 0; index < rawParts.length; index++) {
    var trimmed = String(rawParts[index]).trim();
    if (trimmed.length > 0) {
      normalized.push(trimmed);
    }
  }
  return normalized;
}

function sppr_onprem_extractHost(urlValue, fallbackHost) {
  try {
    if (!spop_isBlank(urlValue)) {
      return String(new java.net.URL(String(urlValue)).getHost()).toLowerCase();
    }
  } catch (ignoreExtractHost) {
  }
  return spop_safeString(fallbackHost).toLowerCase();
}

function sppr_onprem_composeAbsoluteUrl(siteBaseUrl, serverRelativeUrl) {
  var normalizedRelativeUrl = spop_ensureLeadingSlash(spop_safeString(serverRelativeUrl));
  if (spop_isBlank(normalizedRelativeUrl)) {
    return "";
  }
  try {
    var parsedUrl = new java.net.URL(String(siteBaseUrl));
    var authority = parsedUrl.getProtocol() + "://" + parsedUrl.getHost();
    var port = parsedUrl.getPort();
    if (port > 0 && port !== parsedUrl.getDefaultPort()) {
      authority += ":" + port;
    }
    return authority + normalizedRelativeUrl;
  } catch (ignoreComposeAbsoluteUrl) {
    return normalizedRelativeUrl;
  }
}

function sppr_onprem_joinServerRelativePath(parentPath, leafName) {
  var normalizedParent = spop_ensureLeadingSlash(spop_safeString(parentPath));
  if (normalizedParent.length > 1 && normalizedParent.charAt(normalizedParent.length - 1) === "/") {
    normalizedParent = normalizedParent.substring(0, normalizedParent.length - 1);
  }
  return normalizedParent + "/" + String(leafName);
}

function sppr_onprem_splitLeafName(fileName) {
  var normalized = String(fileName);
  var lastDot = normalized.lastIndexOf(".");
  if (lastDot <= 0 || lastDot === normalized.length - 1) {
    return {
      base: normalized,
      extension: ""
    };
  }
  return {
    base: normalized.substring(0, lastDot),
    extension: normalized.substring(lastDot)
  };
}

function sppr_onprem_conflictRename(fileName, attemptIndex) {
  var parts = sppr_onprem_splitLeafName(fileName);
  return parts.base + "-" + String(attemptIndex) + parts.extension;
}

function sppr_onprem_emptyBytes() {
  return java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 0);
}

function sppr_onprem_expectOk(response, label) {
  if (response === null || response === undefined) {
    throw new java.lang.RuntimeException(label + " returned no response");
  }
  if (response.ok !== true) {
    throw new java.lang.RuntimeException(label + " failed: " + spop_safeString(response.errorMessage));
  }
  return response.data;
}

function sppr_onprem_toBoolean(value, defaultValue) {
  return spop_toBoolean(value, defaultValue);
}

function sppr_onprem_baseDriveData(resolvedLibrary) {
  return {
    siteId: spop_safeString(resolvedLibrary.siteId),
    siteHostname: spop_safeString(resolvedLibrary.siteHostname),
    sitePath: spop_safeString(resolvedLibrary.sitePath),
    driveId: spop_safeString(resolvedLibrary.driveId),
    driveName: spop_safeString(resolvedLibrary.driveName),
    driveWebUrl: spop_safeString(resolvedLibrary.driveWebUrl),
    driveType: spop_safeString(resolvedLibrary.driveType),
    sharepointIds: resolvedLibrary.sharepointIds === undefined || resolvedLibrary.sharepointIds === null ? {} : resolvedLibrary.sharepointIds,
    quota: {},
    owner: {},
    resolvedFromList: true,
    listId: spop_safeString(resolvedLibrary.listId),
    listName: spop_safeString(resolvedLibrary.listName),
    listDisplayName: spop_safeString(resolvedLibrary.listDisplayName),
    listWebUrl: spop_safeString(resolvedLibrary.listWebUrl),
    listInfo: resolvedLibrary.listInfo === undefined || resolvedLibrary.listInfo === null ? {} : resolvedLibrary.listInfo,
    listSharepointIds: resolvedLibrary.listSharepointIds === undefined || resolvedLibrary.listSharepointIds === null ? {} : resolvedLibrary.listSharepointIds
  };
}

function sppr_onprem_getSelectFields(parametersJS) {
  return sppr_onprem_splitCsv(sppr_onprem_get(parametersJS, "selectFields", ""));
}

function sppr_onprem_buildDriveResponse(resolvedLibrary, queryData) {
  var responseData = sppr_onprem_baseDriveData(resolvedLibrary);
  responseData.query = queryData === undefined || queryData === null ? {} : queryData;
  return responseData;
}

function sppr_onprem_applyPagedItems(responseData, pageCount, itemCount, nextLink, items, deltaLink) {
  responseData.pageCount = pageCount;
  responseData.itemCount = itemCount;
  responseData.nextLink = spop_safeString(nextLink);
  if (deltaLink !== undefined) {
    responseData.deltaLink = spop_safeString(deltaLink);
  }
  responseData.items = items === undefined || items === null ? [] : items;
  return responseData;
}

function sppr_onprem_resolveLibraryContext(parametersJS) {
  var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
  var libraryReference = sppr_onprem_resolveLibraryReference(parametersJS);
  var siteMeta = sppr_onprem_loadSiteMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth);
  var metadata = spop_loadListMetadata(
    runtimeContext.targetSiteBaseUrl,
    runtimeContext.auth,
    libraryReference.listId,
    libraryReference.listName
  );
  var listObject = metadata.list;
  var rootFolder = listObject.RootFolder === undefined || listObject.RootFolder === null ? {} : listObject.RootFolder;

  return {
    siteBaseUrl: runtimeContext.targetSiteBaseUrl,
    onPremProtocol: runtimeContext.onPremProtocol,
    accessToken: runtimeContext.accessToken,
    onPremUsername: runtimeContext.onPremUsername,
    onPremPassword: runtimeContext.onPremPassword,
    cookieHeader: runtimeContext.cookieHeader,
    auth: runtimeContext.auth,
    tokenMode: runtimeContext.tokenMode,
    siteId: spop_safeString(siteMeta.Id),
    siteHostname: sppr_onprem_extractHost(siteMeta.Url, runtimeContext.siteHostname),
    sitePath: spop_safeString(siteMeta.ServerRelativeUrl),
    listEndpoint: metadata.endpoint,
    listObject: listObject,
    rootFolderServerRelativeUrl: spop_safeString(rootFolder.ServerRelativeUrl),
    driveId: spop_safeString(listObject.Id),
    driveName: spop_safeString(listObject.Title),
    driveWebUrl: spop_safeString(listObject.DefaultViewUrl),
    driveType: "documentLibrary",
    sharepointIds: {
      listId: spop_safeString(listObject.Id),
      siteId: spop_safeString(siteMeta.Id)
    },
    listId: spop_safeString(listObject.Id),
    listName: spop_safeString(listObject.Title),
    listDisplayName: spop_safeString(listObject.Title),
    listWebUrl: spop_safeString(listObject.DefaultViewUrl),
    listInfo: {
      baseTemplate: listObject.BaseTemplate === undefined || listObject.BaseTemplate === null ? "" : String(listObject.BaseTemplate),
      hidden: listObject.Hidden === true,
      itemCount: listObject.ItemCount === undefined || listObject.ItemCount === null ? 0 : listObject.ItemCount,
      rootFolderServerRelativeUrl: spop_safeString(rootFolder.ServerRelativeUrl)
    },
    listSharepointIds: {
      listId: spop_safeString(listObject.Id),
      siteId: spop_safeString(siteMeta.Id)
    },
    createdDateTime: spop_safeString(listObject.Created),
    lastModifiedDateTime: spop_safeString(listObject.LastItemModifiedDate)
  };
}

function sppr_onprem_resolveListContext(parametersJS) {
  var listId = sppr_onprem_get(parametersJS, "listId", "");
  var listName = sppr_onprem_get(parametersJS, "listName", "");
  sppr_onprem_requireListReference(listId, listName);

  var runtimeContext = sppr_onprem_buildRuntimeContext(parametersJS);
  var siteMeta = sppr_onprem_loadSiteMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth);
  var metadata = spop_loadListMetadata(runtimeContext.targetSiteBaseUrl, runtimeContext.auth, listId, listName);
  var listObject = metadata.list;
  var rootFolder = listObject.RootFolder === undefined || listObject.RootFolder === null ? {} : listObject.RootFolder;

  return {
    siteBaseUrl: runtimeContext.targetSiteBaseUrl,
    onPremProtocol: runtimeContext.onPremProtocol,
    accessToken: runtimeContext.accessToken,
    onPremUsername: runtimeContext.onPremUsername,
    onPremPassword: runtimeContext.onPremPassword,
    cookieHeader: runtimeContext.cookieHeader,
    auth: runtimeContext.auth,
    tokenMode: runtimeContext.tokenMode,
    siteId: spop_safeString(siteMeta.Id),
    siteHostname: sppr_onprem_extractHost(siteMeta.Url, runtimeContext.siteHostname),
    sitePath: spop_safeString(siteMeta.ServerRelativeUrl),
    listEndpoint: metadata.endpoint,
    listObject: listObject,
    rootFolderServerRelativeUrl: spop_safeString(rootFolder.ServerRelativeUrl),
    listId: spop_safeString(listObject.Id),
    listName: spop_safeString(listObject.Title),
    listDisplayName: spop_safeString(listObject.Title),
    listWebUrl: spop_safeString(listObject.DefaultViewUrl),
    listInfo: {
      baseTemplate: listObject.BaseTemplate === undefined || listObject.BaseTemplate === null ? "" : String(listObject.BaseTemplate),
      hidden: listObject.Hidden === true,
      itemCount: listObject.ItemCount === undefined || listObject.ItemCount === null ? 0 : listObject.ItemCount,
      rootFolderServerRelativeUrl: spop_safeString(rootFolder.ServerRelativeUrl)
    },
    listSharepointIds: {
      listId: spop_safeString(listObject.Id),
      siteId: spop_safeString(siteMeta.Id)
    },
    createdDateTime: spop_safeString(listObject.Created),
    lastModifiedDateTime: spop_safeString(listObject.LastItemModifiedDate)
  };
}

function sppr_onprem_withResolvedLibraryRequest(baseRequest, resolvedLibrary) {
  baseRequest.siteBaseUrl = resolvedLibrary.siteBaseUrl;
  baseRequest.siteHostname = resolvedLibrary.siteHostname;
  baseRequest.sitePath = resolvedLibrary.sitePath;
  baseRequest.onPremProtocol = resolvedLibrary.onPremProtocol;
  baseRequest.listId = resolvedLibrary.listId;
  baseRequest.listName = resolvedLibrary.listName;
  baseRequest.accessToken = resolvedLibrary.accessToken;
  baseRequest.onPremUsername = resolvedLibrary.onPremUsername;
  baseRequest.onPremPassword = resolvedLibrary.onPremPassword;
  baseRequest.cookieHeader = resolvedLibrary.cookieHeader;
  return baseRequest;
}

function sppr_onprem_requiredListItemFields() {
  return [
    "Id",
    "Title",
    "Created",
    "Modified",
    "AuthorId",
    "EditorId",
    "FileRef",
    "FileDirRef",
    "FileLeafRef",
    "FSObjType",
    "FileSystemObjectType",
    "UniqueId"
  ];
}

function sppr_onprem_mergeSelectFields(selectFields) {
  var merged = [];
  var seen = {};
  var requestedFields = sppr_onprem_splitCsv(selectFields);
  var i;
  var fieldName = "";
  var normalizedKey = "";

  for (i = 0; i < requestedFields.length; i++) {
    fieldName = spop_safeString(requestedFields[i]).trim();
    if (fieldName.length > 0) {
      normalizedKey = fieldName.toLowerCase();
      if (seen[normalizedKey] !== true) {
        merged.push(fieldName);
        seen[normalizedKey] = true;
      }
    }
  }

  var requiredFields = sppr_onprem_requiredListItemFields();
  for (i = 0; i < requiredFields.length; i++) {
    fieldName = requiredFields[i];
    normalizedKey = fieldName.toLowerCase();
    if (seen[normalizedKey] !== true) {
      merged.push(fieldName);
      seen[normalizedKey] = true;
    }
  }

  return merged.join(",");
}

function sppr_onprem_normalizeListExpression(expressionValue) {
  var normalized = spop_safeString(expressionValue);
  if (spop_isBlank(normalized)) {
    return "";
  }
  return normalized.replace(/\bid\b/g, "Id");
}

function sppr_onprem_debugResultSummary(result) {
  if (result === null || result === undefined) {
    return "result=null";
  }
  var body = spop_safeString(result.body);
  if (body.length > 300) {
    body = body.substring(0, 300);
  }
  return "status=" + spop_safeString(result.statusCode) +
    ", location=" + spop_safeString(result.location) +
    ", etag=" + spop_safeString(result.etag) +
    ", body=" + body;
}

function sppr_onprem_extractLookupId(rawValue) {
  var text = spop_safeString(rawValue);
  var separatorIndex = text.indexOf(";#");
  if (separatorIndex >= 0) {
    return text.substring(0, separatorIndex);
  }
  return text;
}

function sppr_onprem_extractLookupValue(rawValue) {
  var text = spop_safeString(rawValue);
  var separatorIndex = text.indexOf(";#");
  if (separatorIndex >= 0) {
    return text.substring(separatorIndex + 2);
  }
  return text;
}

function sppr_onprem_normalizeDateTime(value) {
  var text = spop_safeString(value);
  if (text.length === 19 && text.charAt(10) === " ") {
    return text.substring(0, 10) + "T" + text.substring(11) + "Z";
  }
  if (text.length === 17 && text.charAt(8) === " ") {
    return text.substring(0, 4) + "-" + text.substring(4, 6) + "-" + text.substring(6, 8) + "T" + text.substring(9) + "Z";
  }
  return text;
}

function sppr_onprem_parseIntegerValue(rawValue, defaultValue) {
  var text = spop_safeString(rawValue);
  if (spop_isBlank(text)) {
    return defaultValue;
  }
  var normalized = text;
  var semicolonIndex = normalized.indexOf(";#");
  if (semicolonIndex >= 0) {
    normalized = normalized.substring(0, semicolonIndex);
  }
  return spop_parseInteger("integerValue", normalized, defaultValue);
}

function sppr_onprem_escapeXml(value) {
  return spop_safeString(value)
    .split("&").join("&amp;")
    .split("<").join("&lt;")
    .split(">").join("&gt;")
    .split("\"").join("&quot;")
    .split("'").join("&apos;");
}

function sppr_onprem_parseXmlDocument(xmlText) {
  var DocumentBuilderFactory = Packages.javax.xml.parsers.DocumentBuilderFactory;
  var ByteArrayInputStream = Packages.java.io.ByteArrayInputStream;
  var factory = DocumentBuilderFactory.newInstance();
  factory.setNamespaceAware(true);
  var builder = factory.newDocumentBuilder();
  var xmlBytes = new java.lang.String(spop_defaultString(xmlText, "")).getBytes("UTF-8");
  return builder.parse(new ByteArrayInputStream(xmlBytes));
}

function sppr_onprem_nodeLocalName(node) {
  if (node === null || node === undefined) {
    return "";
  }
  var localName = "";
  try {
    localName = spop_safeString(node.getLocalName());
  } catch (ignoreLocalName) {
  }
  if (!spop_isBlank(localName)) {
    return localName;
  }
  var nodeName = "";
  try {
    nodeName = spop_safeString(node.getNodeName());
  } catch (ignoreNodeName) {
  }
  var separatorIndex = nodeName.indexOf(":");
  if (separatorIndex >= 0) {
    return nodeName.substring(separatorIndex + 1);
  }
  return nodeName;
}

function sppr_onprem_findFirstElementByLocalName(parentNode, localName) {
  if (parentNode === null || parentNode === undefined) {
    return null;
  }
  var elements = parentNode.getElementsByTagName("*");
  var index;
  for (index = 0; index < elements.getLength(); index++) {
    var current = elements.item(index);
    if (sppr_onprem_nodeLocalName(current) === localName) {
      return current;
    }
  }
  return null;
}

function sppr_onprem_collectChildElementsByLocalName(parentNode, localName) {
  var matches = [];
  if (parentNode === null || parentNode === undefined) {
    return matches;
  }
  var children = parentNode.getChildNodes();
  var Node = Packages.org.w3c.dom.Node;
  var index;
  for (index = 0; index < children.getLength(); index++) {
    var current = children.item(index);
    if (current.getNodeType() === Node.ELEMENT_NODE && sppr_onprem_nodeLocalName(current) === localName) {
      matches.push(current);
    }
  }
  return matches;
}

function sppr_onprem_collectAttributes(elementNode) {
  var attributes = {};
  if (elementNode === null || elementNode === undefined || !elementNode.getAttributes) {
    return attributes;
  }
  var namedNodeMap = elementNode.getAttributes();
  var index;
  for (index = 0; index < namedNodeMap.getLength(); index++) {
    var attributeNode = namedNodeMap.item(index);
    attributes[spop_safeString(attributeNode.getNodeName())] = spop_safeString(attributeNode.getNodeValue());
  }
  return attributes;
}

function sppr_onprem_buildChangesQueryOptions(folderServerRelativeUrl, pagingToken) {
  var parts = [
    "<QueryOptions>",
    "<DateInUtc>True</DateInUtc>",
    "<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>",
    "<ViewAttributes Scope=\"RecursiveAll\" />"
  ];
  if (!spop_isBlank(folderServerRelativeUrl)) {
    parts.push("<Folder>" + sppr_onprem_escapeXml(folderServerRelativeUrl) + "</Folder>");
  }
  if (!spop_isBlank(pagingToken)) {
    parts.push("<Paging ListItemCollectionPositionNext=\"" + sppr_onprem_escapeXml(pagingToken) + "\" />");
  }
  parts.push("</QueryOptions>");
  return parts.join("");
}

function sppr_onprem_mapChangeRowToListItem(attributes, includeRaw) {
  var fields = {};
  var attributeName;
  for (attributeName in attributes) {
    if (Object.prototype.hasOwnProperty.call(attributes, attributeName) && attributeName.indexOf("ows_") === 0) {
      var fieldName = attributeName.substring(4);
      var rawValue = spop_safeString(attributes[attributeName]);
      var normalizedValue = rawValue;
      if (fieldName === "FileRef" || fieldName === "FileDirRef" || fieldName === "FileLeafRef" || fieldName === "UniqueId") {
        normalizedValue = sppr_onprem_extractLookupValue(rawValue);
      } else if (fieldName === "FSObjType" || fieldName === "FileSystemObjectType") {
        normalizedValue = sppr_onprem_extractLookupId(rawValue);
      } else if (fieldName === "Author" || fieldName === "Editor") {
        normalizedValue = sppr_onprem_extractLookupId(rawValue);
      }
      fields[fieldName] = normalizedValue;
    }
  }

  if (spop_isBlank(fields.FileSystemObjectType)) {
    fields.FileSystemObjectType = spop_safeString(fields.FSObjType);
  }

  var itemId = spop_safeString(fields.ID);
  if (spop_isBlank(itemId)) {
    itemId = spop_safeString(fields.Id);
  }
  var item = {
    itemId: itemId,
    id: itemId,
    title: spop_safeString(fields.Title),
    modified: sppr_onprem_normalizeDateTime(fields.Modified),
    created: sppr_onprem_normalizeDateTime(fields.Created),
    authorId: sppr_onprem_extractLookupId(attributes.ows_Author),
    editorId: sppr_onprem_extractLookupId(attributes.ows_Editor),
    fields: fields
  };
  if (includeRaw) {
    item.raw = attributes;
  }
  return item;
}

function sppr_onprem_buildRemovedListItem(itemId) {
  return {
    itemId: spop_safeString(itemId),
    id: spop_safeString(itemId),
    title: "",
    modified: "",
    created: "",
    authorId: "",
    editorId: "",
    fields: {},
    removed: {
      reason: "deleted"
    }
  };
}

function sppr_onprem_buildRemovedDriveItem(resolvedLibrary, itemId) {
  var normalizedItemId = spop_safeString(itemId);
  return {
    id: normalizedItemId,
    itemId: normalizedItemId,
    name: "",
    webUrl: "",
    size: 0,
    eTag: "",
    cTag: "",
    createdDateTime: "",
    lastModifiedDateTime: "",
    file: {},
    folder: {},
    image: {},
    shared: {},
    parentReference: sppr_onprem_buildDriveParentReference(resolvedLibrary, ""),
    createdBy: {},
    lastModifiedBy: {},
    fileSystemInfo: sppr_onprem_buildDriveFileSystemInfo("", ""),
    sharepointIds: sppr_onprem_buildDriveSharepointIds(resolvedLibrary, normalizedItemId, ""),
    removed: {
      reason: "deleted"
    }
  };
}

function sppr_onprem_requestListChanges(siteBaseUrl, auth, listId, listName, folderServerRelativeUrl, changeToken, pagingToken, rowLimit) {
  var contextInfo = spop_getContextDigest(siteBaseUrl, auth);
  var headers = {
    "X-RequestDigest": contextInfo.formDigestValue
  };
  var payloadObject = {
    query: {
      "__metadata": {
        type: "SP.ChangeLogItemQuery"
      },
      QueryOptions: sppr_onprem_buildChangesQueryOptions(folderServerRelativeUrl, pagingToken),
      RowLimit: String(rowLimit),
      ChangeToken: spop_isBlank(changeToken) ? null : String(changeToken)
    }
  };
  var result = spop_httpRequest(
    "POST",
    spop_buildListEndpoint(siteBaseUrl, listId, listName) + "/GetListItemChangesSinceToken",
    auth,
    new java.lang.String(JSON.stringify(payloadObject)).getBytes("UTF-8"),
    "application/json;odata=verbose;charset=utf-8",
    "application/json;odata=verbose",
    headers,
    false
  );
  spop_expectSuccess(result, "GetListItemChangesSinceToken");
  return result.body;
}

function sppr_onprem_parseListChangesResponse(xmlText, includeRaw) {
  var document = sppr_onprem_parseXmlDocument(xmlText);
  var changesElement = sppr_onprem_findFirstElementByLocalName(document, "Changes");
  var dataElement = sppr_onprem_findFirstElementByLocalName(document, "data");
  var rowElements = document.getElementsByTagName("*");
  var rows = [];
  var removedIds = [];
  var index;

  for (index = 0; index < rowElements.getLength(); index++) {
    var currentElement = rowElements.item(index);
    if (sppr_onprem_nodeLocalName(currentElement) === "row") {
      rows.push(sppr_onprem_mapChangeRowToListItem(sppr_onprem_collectAttributes(currentElement), includeRaw));
    }
  }

  var changeEntries = sppr_onprem_collectChildElementsByLocalName(changesElement, "Id");
  for (index = 0; index < changeEntries.length; index++) {
    var currentChange = changeEntries[index];
    var changeType = spop_safeString(currentChange.getAttribute("ChangeType")).toLowerCase();
    if (changeType.indexOf("delete") >= 0) {
      removedIds.push(spop_safeString(currentChange.getTextContent()));
    }
  }

  return {
    lastChangeToken: changesElement === null ? "" : spop_safeString(changesElement.getAttribute("LastChangeToken")),
    pagingToken: dataElement === null ? "" : spop_safeString(dataElement.getAttribute("ListItemCollectionPositionNext")),
    rows: rows,
    removedIds: removedIds,
    rawXml: includeRaw ? spop_safeString(xmlText) : ""
  };
}

function sppr_onprem_resolveChangeToken(deltaLink, deltaToken) {
  if (!spop_isBlank(deltaLink)) {
    return spop_safeString(deltaLink);
  }
  return spop_safeString(deltaToken);
}

function sppr_onprem_queryListChanges(resolvedLibrary, folderServerRelativeUrl, deltaLink, deltaToken, top, maxPages, includeRaw) {
  var pageSize = sppr_onprem_parseBoundedInteger("top", top, 200, 1, null);
  var maxPageCount = sppr_onprem_parseBoundedInteger("maxPages", maxPages, 20, 1, null);

  var changeToken = sppr_onprem_resolveChangeToken(deltaLink, deltaToken);
  var pagingToken = "";
  var pageCount = 0;
  var latestDeltaLink = "";
  var items = [];
  var removedIds = [];
  var removedLookup = {};

  while (pageCount < maxPageCount) {
    pageCount++;
    var xmlText = sppr_onprem_requestListChanges(
      resolvedLibrary.siteBaseUrl,
      resolvedLibrary.auth,
      resolvedLibrary.listId,
      resolvedLibrary.listName,
      folderServerRelativeUrl,
      changeToken,
      pagingToken,
      pageSize
    );
    var parsed = sppr_onprem_parseListChangesResponse(xmlText, includeRaw);
    var i;
    for (i = 0; i < parsed.rows.length; i++) {
      items.push(parsed.rows[i]);
    }
    for (i = 0; i < parsed.removedIds.length; i++) {
      var removedId = spop_safeString(parsed.removedIds[i]);
      if (!spop_isBlank(removedId) && removedLookup[removedId] !== true) {
        removedLookup[removedId] = true;
        removedIds.push(removedId);
      }
    }
    latestDeltaLink = spop_safeString(parsed.lastChangeToken);
    pagingToken = spop_safeString(parsed.pagingToken);
    if (spop_isBlank(pagingToken)) {
      break;
    }
  }

  return {
    pageCount: pageCount,
    nextLink: pagingToken,
    deltaLink: latestDeltaLink,
    items: items,
    removedIds: removedIds
  };
}

function sppr_onprem_buildCopyMonitorUrl(copiedItemId, normalizedName) {
  return "onprem-copy://completed?copiedItemId=" + spop_encodeQueryValue(copiedItemId) + "&normalizedName=" + spop_encodeQueryValue(normalizedName);
}

function sppr_onprem_parseCopyMonitorUrl(monitorUrl) {
  var parsed = {
    copiedItemId: "",
    normalizedName: ""
  };
  var text = spop_safeString(monitorUrl);
  var queryIndex = text.indexOf("?");
  if (queryIndex < 0) {
    return parsed;
  }
  var query = text.substring(queryIndex + 1).split("&");
  var index;
  for (index = 0; index < query.length; index++) {
    var chunk = query[index];
    var separatorIndex = chunk.indexOf("=");
    var key = separatorIndex >= 0 ? chunk.substring(0, separatorIndex) : chunk;
    var value = separatorIndex >= 0 ? chunk.substring(separatorIndex + 1) : "";
    if (key === "copiedItemId") {
      parsed.copiedItemId = java.net.URLDecoder.decode(value, "UTF-8");
    } else if (key === "normalizedName") {
      parsed.normalizedName = java.net.URLDecoder.decode(value, "UTF-8");
    }
  }
  return parsed;
}

function sppr_onprem_buildDriveParentReference(resolvedLibrary, parentPath) {
  return {
    driveId: spop_safeString(resolvedLibrary.driveId),
    siteId: spop_safeString(resolvedLibrary.siteId),
    path: spop_safeString(parentPath)
  };
}

function sppr_onprem_buildDriveFileSystemInfo(createdDateTime, lastModifiedDateTime) {
  return {
    createdDateTime: spop_safeString(createdDateTime),
    lastModifiedDateTime: spop_safeString(lastModifiedDateTime)
  };
}

function sppr_onprem_buildDriveSharepointIds(resolvedLibrary, listItemId, uniqueId) {
  var sharepointIds = {
    listId: spop_safeString(resolvedLibrary.listId),
    siteId: spop_safeString(resolvedLibrary.siteId)
  };
  if (!spop_isBlank(listItemId)) {
    sharepointIds.listItemId = spop_safeString(listItemId);
  }
  if (!spop_isBlank(uniqueId)) {
    sharepointIds.uniqueId = spop_safeString(uniqueId);
  }
  return sharepointIds;
}

function sppr_onprem_extractFileSystemType(fields) {
  var fileSystemType = spop_safeString(fields.FileSystemObjectType);
  if (spop_isBlank(fileSystemType)) {
    fileSystemType = spop_safeString(fields.FSObjType);
  }
  return fileSystemType;
}

function sppr_onprem_requireChildItemId(rawItemId, rootErrorMessage) {
  var targetItemId = spop_require("itemId", rawItemId);
  if (String(targetItemId).toLowerCase() === "root") {
    throw new java.lang.IllegalArgumentException(rootErrorMessage);
  }
  return targetItemId;
}

function sppr_onprem_loadDriveItemContext(resolvedLibrary, itemId, includeRaw) {
  var itemData = sppr_onprem_expectOk(
    sppr_onPremGetListItem(sppr_onprem_withResolvedLibraryRequest({
      itemId: itemId,
      selectFields: "",
      includeRawItem: includeRaw === true ? "true" : "false"
    }, resolvedLibrary)),
    "onprem_get_list_item"
  );
  var item = itemData.item === undefined || itemData.item === null ? {} : itemData.item;
  var fields = item.fields === undefined || item.fields === null ? {} : item.fields;
  return {
    itemData: itemData,
    item: item,
    fields: fields,
    itemType: sppr_onprem_extractFileSystemType(fields),
    itemPath: spop_safeString(fields.FileRef),
    parentPath: spop_safeString(fields.FileDirRef),
    itemName: spop_safeString(fields.FileLeafRef)
  };
}

function sppr_onprem_resolveDriveTargetPath(resolvedLibrary, currentParentPath, currentName, requestedDestinationParent, requestedNewName, replaceTimestampTokens) {
  var targetParentPath = currentParentPath;
  var resolvedDestinationParentId = "";
  if (!spop_isBlank(requestedDestinationParent)) {
    targetParentPath = sppr_onprem_resolveParentFolderServerRelativeUrl(resolvedLibrary, requestedDestinationParent);
    resolvedDestinationParentId = requestedDestinationParent;
  }
  var targetName = spop_isBlank(requestedNewName) ? currentName : String(requestedNewName);
  if (replaceTimestampTokens) {
    targetName = targetName.split("[[timestamp]]").join(String(java.lang.System.currentTimeMillis()));
  }
  targetName = targetName.trim();
  return {
    targetParentPath: targetParentPath,
    resolvedDestinationParentItemId: resolvedDestinationParentId,
    targetName: targetName,
    targetPath: sppr_onprem_joinServerRelativePath(targetParentPath, targetName)
  };
}

function sppr_onprem_mapListItemToDriveItem(resolvedLibrary, listItemObject, includeRaw) {
  var listItem = listItemObject === undefined || listItemObject === null ? {} : listItemObject;
  var fields = listItem.fields === undefined || listItem.fields === null ? {} : listItem.fields;
  var fileSystemType = sppr_onprem_extractFileSystemType(fields);
  var isFolder = fileSystemType === "1";
  var itemServerRelativeUrl = spop_ensureLeadingSlash(spop_safeString(fields.FileRef));
  var parentServerRelativeUrl = spop_ensureLeadingSlash(spop_safeString(fields.FileDirRef));
  var itemName = spop_safeString(fields.FileLeafRef);
  if (spop_isBlank(itemName)) {
    itemName = spop_safeString(listItem.title);
  }

  var sizeValue = 0;
  if (!isFolder && !spop_isBlank(fields.File_x0020_Size)) {
    sizeValue = spop_parseInteger("size", fields.File_x0020_Size, 0);
  }

  var normalizedItem = {
    id: spop_safeString(listItem.itemId),
    itemId: spop_safeString(listItem.itemId),
    name: itemName,
    webUrl: sppr_onprem_composeAbsoluteUrl(resolvedLibrary.siteBaseUrl, itemServerRelativeUrl),
    size: sizeValue,
    eTag: "",
    cTag: "",
    createdDateTime: spop_safeString(listItem.created),
    lastModifiedDateTime: spop_safeString(listItem.modified),
    file: isFolder ? {} : {},
    folder: isFolder ? { childCount: 0 } : {},
    image: {},
    shared: {},
    parentReference: sppr_onprem_buildDriveParentReference(resolvedLibrary, parentServerRelativeUrl),
    createdBy: {
      userId: spop_safeString(fields.AuthorId)
    },
    lastModifiedBy: {
      userId: spop_safeString(fields.EditorId)
    },
    fileSystemInfo: sppr_onprem_buildDriveFileSystemInfo(listItem.created, listItem.modified),
    sharepointIds: sppr_onprem_buildDriveSharepointIds(resolvedLibrary, listItem.itemId, fields.UniqueId),
    removed: {}
  };

  if (includeRaw) {
    normalizedItem.raw = listItem.raw === undefined || listItem.raw === null ? fields : listItem.raw;
  }
  return normalizedItem;
}

function sppr_onprem_buildRootDriveItem(resolvedLibrary, includeRaw) {
  var rootFolderServerRelativeUrl = spop_safeString(resolvedLibrary.rootFolderServerRelativeUrl);
  var item = {
    id: "root",
    itemId: "root",
    name: spop_safeString(resolvedLibrary.driveName),
    webUrl: sppr_onprem_composeAbsoluteUrl(resolvedLibrary.siteBaseUrl, rootFolderServerRelativeUrl),
    size: 0,
    eTag: "",
    cTag: "",
    createdDateTime: spop_safeString(resolvedLibrary.createdDateTime),
    lastModifiedDateTime: spop_safeString(resolvedLibrary.lastModifiedDateTime),
    file: {},
    folder: {},
    image: {},
    shared: {},
    parentReference: sppr_onprem_buildDriveParentReference(resolvedLibrary, rootFolderServerRelativeUrl),
    createdBy: {},
    lastModifiedBy: {},
    fileSystemInfo: sppr_onprem_buildDriveFileSystemInfo(resolvedLibrary.createdDateTime, resolvedLibrary.lastModifiedDateTime),
    sharepointIds: sppr_onprem_buildDriveSharepointIds(resolvedLibrary, "", ""),
    removed: {}
  };
  if (includeRaw) {
    item.raw = {
      rootFolderServerRelativeUrl: rootFolderServerRelativeUrl
    };
  }
  return item;
}

function sppr_onprem_getListItemResponse(resolvedLibrary, itemId, includeRaw) {
  return sppr_onPremGetListItem(sppr_onprem_withResolvedLibraryRequest({
    itemId: itemId,
    selectFields: "",
    includeRawItem: includeRaw === true ? "true" : "false"
  }, resolvedLibrary));
}

function sppr_onprem_queryListItems(resolvedLibrary, filterExpression, orderByExpression, top, maxPages, includeRaw) {
  return sppr_onPremListGetItems(sppr_onprem_withResolvedLibraryRequest({
    top: String(top),
    maxPages: String(maxPages),
    orderByExpression: orderByExpression,
    filterExpression: filterExpression,
    selectFields: "",
    expandFieldValuesAsText: "false",
    includeRawItem: includeRaw === true ? "true" : "false"
  }, resolvedLibrary));
}

function sppr_onprem_findDriveItemByFileRef(resolvedLibrary, fileRef, includeRaw) {
  var filterExpression = "FileRef eq '" + spop_encodeODataStringLiteral(fileRef) + "'";
  var responseData = sppr_onprem_expectOk(
    sppr_onPremListGetItems(sppr_onprem_withResolvedLibraryRequest({
      top: "1",
      maxPages: "1",
      orderByExpression: "Id desc",
      filterExpression: filterExpression,
      selectFields: "",
      expandFieldValuesAsText: "false",
      includeRawItem: includeRaw === true ? "true" : "false"
    }, resolvedLibrary)),
    "onprem_list_get_items"
  );
  var items = responseData.items === undefined || responseData.items === null ? [] : responseData.items;
  if (!(items instanceof Array) || items.length === 0) {
    return null;
  }
  return sppr_onprem_mapListItemToDriveItem(resolvedLibrary, items[0], includeRaw);
}

function sppr_onprem_findDriveItemByParentAndName(resolvedLibrary, parentServerRelativeUrl, itemName, includeRaw) {
  var filterExpression =
    "FileDirRef eq '" + spop_encodeODataStringLiteral(parentServerRelativeUrl) + "'" +
    " and FileLeafRef eq '" + spop_encodeODataStringLiteral(itemName) + "'";
  var responseData = sppr_onprem_expectOk(
    sppr_onPremListGetItems(sppr_onprem_withResolvedLibraryRequest({
      top: "1",
      maxPages: "1",
      orderByExpression: "Id desc",
      filterExpression: filterExpression,
      selectFields: "",
      expandFieldValuesAsText: "false",
      includeRawItem: includeRaw === true ? "true" : "false"
    }, resolvedLibrary)),
    "onprem_list_get_items"
  );
  var items = responseData.items === undefined || responseData.items === null ? [] : responseData.items;
  if (!(items instanceof Array) || items.length === 0) {
    return null;
  }
  return sppr_onprem_mapListItemToDriveItem(resolvedLibrary, items[0], includeRaw);
}

function sppr_onprem_makeAvailableChildName(resolvedLibrary, parentServerRelativeUrl, requestedName) {
  var candidateName = requestedName;
  var attemptIndex = 0;
  while (sppr_onprem_findDriveItemByParentAndName(resolvedLibrary, parentServerRelativeUrl, candidateName, false) !== null) {
    attemptIndex++;
    if (attemptIndex > 100) {
      throw new java.lang.RuntimeException("Unable to generate a unique child name under " + parentServerRelativeUrl);
    }
    candidateName = sppr_onprem_conflictRename(requestedName, attemptIndex);
  }
  return candidateName;
}

function sppr_onprem_translateDriveExpression(expressionValue) {
  var translated = spop_safeString(expressionValue);
  if (spop_isBlank(translated)) {
    return "";
  }
  translated = translated.replace(/\bname\b/g, "FileLeafRef");
  translated = translated.replace(/\blastModifiedDateTime\b/g, "Modified");
  translated = translated.replace(/\bcreatedDateTime\b/g, "Created");
  translated = translated.replace(/\bsize\b/g, "File_x0020_Size");
  return translated;
}

function sppr_onprem_resolveParentFolderServerRelativeUrl(resolvedLibrary, parentItemId) {
  var normalizedParentId = spop_safeString(parentItemId);
  if (spop_isBlank(normalizedParentId) || normalizedParentId.toLowerCase() === "root") {
    return spop_safeString(resolvedLibrary.rootFolderServerRelativeUrl);
  }
  var parentData = sppr_onprem_expectOk(
    sppr_onPremGetListItem(sppr_onprem_withResolvedLibraryRequest({
      itemId: normalizedParentId,
      selectFields: "",
      includeRawItem: "false"
    }, resolvedLibrary)),
    "onprem_get_list_item"
  );
  var parentItem = parentData.item === undefined || parentData.item === null ? {} : parentData.item;
  var parentFields = parentItem.fields === undefined || parentItem.fields === null ? {} : parentItem.fields;
  var parentType = spop_safeString(parentFields.FileSystemObjectType);
  if (spop_isBlank(parentType)) {
    parentType = spop_safeString(parentFields.FSObjType);
  }
  if (parentType !== "1") {
    throw new java.lang.IllegalArgumentException("parentItemId must reference a folder in on-prem mode");
  }
  return spop_safeString(parentFields.FileRef);
}

function sppr_onPremGetDriveItem(parametersJS) {
  var graphOperation = "onprem_get_drive_item";
  var requiredPermission = "SharePoint on-prem list read";

  try {
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var requestedItemId = spop_safeString(sppr_onprem_get(parametersJS, "itemId", ""));
    var normalizedItemId = spop_isBlank(requestedItemId) ? "root" : requestedItemId;
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var item = null;

    if (normalizedItemId === "root") {
      item = sppr_onprem_buildRootDriveItem(resolvedLibrary, includeRaw);
    } else {
      var itemData = sppr_onprem_expectOk(
        sppr_onPremGetListItem(sppr_onprem_withResolvedLibraryRequest({
          itemId: normalizedItemId,
          selectFields: "",
          includeRawItem: includeRaw === true ? "true" : "false"
        }, resolvedLibrary)),
        "onprem_get_list_item"
      );
      item = sppr_onprem_mapListItemToDriveItem(resolvedLibrary, itemData.item, includeRaw);
    }

    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      itemId: normalizedItemId,
      selectFields: sppr_onprem_getSelectFields(parametersJS),
      includeRawItem: includeRaw
    });
    responseData.item = item;
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremListDriveItems(parametersJS) {
  var graphOperation = "onprem_list_drive_items";
  var requiredPermission = "SharePoint on-prem list read";

  try {
    var pageTop = sppr_onprem_parseBoundedInteger("top", sppr_onprem_get(parametersJS, "top", ""), 200, 1, 5000);
    var maxPageCount = sppr_onprem_parseBoundedInteger("maxPages", sppr_onprem_get(parametersJS, "maxPages", ""), 20, 1, 100);

    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var parentId = spop_safeString(sppr_onprem_get(parametersJS, "parentItemId", ""));
    var normalizedParentId = spop_isBlank(parentId) ? "root" : parentId;
    var translatedOrderBy = sppr_onprem_translateDriveExpression(sppr_onprem_get(parametersJS, "orderBy", ""));
    var translatedFilter = sppr_onprem_translateDriveExpression(sppr_onprem_get(parametersJS, "filter", ""));
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var parentFolderServerRelativeUrl = sppr_onprem_resolveParentFolderServerRelativeUrl(resolvedLibrary, normalizedParentId);

    var parentFilter = "FileDirRef eq '" + spop_encodeODataStringLiteral(parentFolderServerRelativeUrl) + "'";
    var combinedFilter = parentFilter;
    if (!spop_isBlank(translatedFilter)) {
      combinedFilter = "(" + parentFilter + ") and (" + translatedFilter + ")";
    }

    var itemsData = sppr_onprem_expectOk(
      sppr_onprem_queryListItems(resolvedLibrary, combinedFilter, translatedOrderBy, pageTop, maxPageCount, includeRaw),
      "onprem_list_get_items"
    );

    var rawItems = itemsData.items === undefined || itemsData.items === null ? [] : itemsData.items;
    var normalizedItems = [];
    var index;
    for (index = 0; index < rawItems.length; index++) {
      normalizedItems.push(sppr_onprem_mapListItemToDriveItem(resolvedLibrary, rawItems[index], includeRaw));
    }

    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      top: pageTop,
      maxPages: maxPageCount,
      orderBy: spop_safeString(sppr_onprem_get(parametersJS, "orderBy", "")),
      filter: spop_safeString(sppr_onprem_get(parametersJS, "filter", "")),
      selectFields: sppr_onprem_getSelectFields(parametersJS),
      includeRawItem: includeRaw
    });
    responseData.parentItemId = normalizedParentId;
    responseData.count = normalizedItems.length;
    responseData.pagesFetched = spop_parseInteger("pagesFetched", itemsData.pageCount, 0);
    responseData.hasMore = !spop_isBlank(itemsData.nextLink);
    responseData.nextLink = spop_safeString(itemsData.nextLink);
    responseData.items = normalizedItems;
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremCreateDriveFolder(parametersJS) {
  var graphOperation = "onprem_create_drive_folder";
  var requiredPermission = "SharePoint on-prem list write";

  try {
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var requestedParentId = spop_safeString(sppr_onprem_get(parametersJS, "parentItemId", ""));
    var normalizedParentId = spop_isBlank(requestedParentId) ? "root" : requestedParentId;
    var targetFolderName = spop_require("folderName", sppr_onprem_get(parametersJS, "folderName", "")).trim();
    var conflictMode = spop_defaultString(sppr_onprem_get(parametersJS, "conflictBehavior", "rename"), "rename").toLowerCase();
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var parentFolderServerRelativeUrl = sppr_onprem_resolveParentFolderServerRelativeUrl(resolvedLibrary, normalizedParentId);
    var existingItem = sppr_onprem_findDriveItemByParentAndName(resolvedLibrary, parentFolderServerRelativeUrl, targetFolderName, false);

    if (existingItem !== null) {
      if (conflictMode === "fail") {
        throw new java.lang.RuntimeException("A child named " + targetFolderName + " already exists under " + parentFolderServerRelativeUrl);
      }
      if (conflictMode === "replace") {
        sppr_onprem_expectOk(
          sppr_onPremDeleteListItem(sppr_onprem_withResolvedLibraryRequest({
            itemId: existingItem.itemId,
            includeDeletedItemSnapshot: "false",
            includeRawItem: "false"
          }, resolvedLibrary)),
          "onprem_delete_list_item"
        );
      } else {
        targetFolderName = sppr_onprem_makeAvailableChildName(resolvedLibrary, parentFolderServerRelativeUrl, targetFolderName);
      }
    }

    var targetFolderPath = sppr_onprem_joinServerRelativePath(parentFolderServerRelativeUrl, targetFolderName);
    var contextInfo = spop_getContextDigest(resolvedLibrary.siteBaseUrl, resolvedLibrary.auth);
    var headers = {
      "X-RequestDigest": contextInfo.formDigestValue
    };
    var createFolderResult = spop_httpRequest(
      "POST",
      resolvedLibrary.siteBaseUrl + "/_api/web/folders",
      resolvedLibrary.auth,
      new java.lang.String(JSON.stringify({
        "__metadata": {
          type: "SP.Folder"
        },
        "ServerRelativeUrl": targetFolderPath
      })).getBytes("UTF-8"),
      "application/json;odata=verbose;charset=utf-8",
      "application/json;odata=verbose",
      headers,
      false
    );
    spop_expectSuccess(createFolderResult, "Create drive folder");

    var createdItem = sppr_onprem_findDriveItemByFileRef(resolvedLibrary, targetFolderPath, includeRaw);
    if (createdItem === null) {
      throw new java.lang.RuntimeException("Unable to resolve created folder after creation (" + sppr_onprem_debugResultSummary(createFolderResult) + ")");
    }

    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      parentItemId: normalizedParentId,
      folderName: spop_safeString(sppr_onprem_get(parametersJS, "folderName", "")),
      resolvedFolderName: targetFolderName,
      conflictBehavior: conflictMode,
      returnCreatedItem: sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "returnCreatedItem", "true"), true),
      selectFields: sppr_onprem_getSelectFields(parametersJS),
      includeRawItem: includeRaw
    });
    responseData.item = createdItem;
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onprem_uploadDriveItemCommon(parametersJS, graphOperation) {
  var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
  var explicitItemId = spop_safeString(sppr_onprem_get(parametersJS, "itemId", ""));
  var requestedParentId = spop_safeString(sppr_onprem_get(parametersJS, "parentItemId", ""));
  var normalizedParentId = spop_isBlank(requestedParentId) ? "root" : requestedParentId;
  var targetFileName = spop_safeString(sppr_onprem_get(parametersJS, "fileName", ""));
  var binaryContent = spop_decodeBase64("contentBase64", sppr_onprem_get(parametersJS, "contentBase64", ""));
  var resolvedContentType = spop_defaultString(sppr_onprem_get(parametersJS, "contentType", "application/octet-stream"), "application/octet-stream");
  var conflictMode = spop_defaultString(sppr_onprem_get(parametersJS, "conflictBehavior", "replace"), "replace").toLowerCase();
  var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
  var contextInfo = spop_getContextDigest(resolvedLibrary.siteBaseUrl, resolvedLibrary.auth);
  var headers = {
    "X-RequestDigest": contextInfo.formDigestValue
  };
  var uploadTarget = "";
  var uploadedItem = null;

  if (!spop_isBlank(explicitItemId)) {
    explicitItemId = sppr_onprem_requireChildItemId(explicitItemId, "Uploading on drive root is not allowed. Provide a file item id.");
    uploadTarget = "item";
    var currentItemContext = sppr_onprem_loadDriveItemContext(resolvedLibrary, explicitItemId, false);
    if (currentItemContext.itemType === "1") {
      throw new java.lang.IllegalArgumentException("Uploading content on a folder is not allowed. Provide a file item id.");
    }
    targetFileName = currentItemContext.itemName;
    headers["IF-MATCH"] = "*";
    headers["X-HTTP-Method"] = "PUT";
    spop_expectSuccess(
      spop_httpRequest(
        "POST",
        spop_buildGetFileEndpoint(resolvedLibrary.siteBaseUrl, currentItemContext.itemPath) + "/$value",
        resolvedLibrary.auth,
        binaryContent,
        resolvedContentType,
        "application/json;odata=verbose",
        headers,
        false
      ),
      "Upload file content"
    );
    uploadedItem = sppr_onprem_expectOk(
      sppr_onPremGetDriveItem(sppr_onprem_withResolvedLibraryRequest({
        itemId: explicitItemId,
        includeRawItem: includeRaw === true ? "true" : "false"
      }, resolvedLibrary)),
      "onprem_get_drive_item"
    ).item;
  } else {
    uploadTarget = "parent";
    targetFileName = spop_require("fileName", targetFileName).trim();
    if (targetFileName.length === 0) {
      throw new java.lang.IllegalArgumentException("fileName cannot be empty");
    }
    var parentFolderServerRelativeUrl = sppr_onprem_resolveParentFolderServerRelativeUrl(resolvedLibrary, normalizedParentId);
    var existingItem = sppr_onprem_findDriveItemByParentAndName(resolvedLibrary, parentFolderServerRelativeUrl, targetFileName, false);
    if (existingItem !== null) {
      if (conflictMode === "fail") {
        throw new java.lang.RuntimeException("A child named " + targetFileName + " already exists under " + parentFolderServerRelativeUrl);
      }
      if (conflictMode === "rename") {
        targetFileName = sppr_onprem_makeAvailableChildName(resolvedLibrary, parentFolderServerRelativeUrl, targetFileName);
      }
    }
    var uploadResult = spop_httpRequest(
      "POST",
      spop_buildGetFolderEndpoint(resolvedLibrary.siteBaseUrl, parentFolderServerRelativeUrl)
        + "/Files/add(url='" + spop_encodeODataStringLiteral(targetFileName) + "',overwrite=true)",
      resolvedLibrary.auth,
      binaryContent,
      resolvedContentType,
      "application/json;odata=verbose",
      headers,
      false
    );
    spop_expectSuccess(uploadResult, "Upload file content");
    uploadedItem = sppr_onprem_findDriveItemByFileRef(
      resolvedLibrary,
      sppr_onprem_joinServerRelativePath(parentFolderServerRelativeUrl, targetFileName),
      includeRaw
    );
    if (uploadedItem === null) {
      throw new java.lang.RuntimeException("Unable to resolve uploaded item after upload (" + sppr_onprem_debugResultSummary(uploadResult) + ")");
    }
  }

  var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
    uploadTarget: uploadTarget,
    itemId: explicitItemId,
    parentItemId: normalizedParentId,
    fileName: targetFileName,
    contentType: resolvedContentType,
    contentBytes: binaryContent.length,
    returnUploadedItem: sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "returnUploadedItem", "true"), true),
    selectFields: sppr_onprem_getSelectFields(parametersJS),
    includeRawItem: includeRaw
  });
  responseData.item = uploadedItem;
  return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, "SharePoint on-prem list write", responseData);
}

function sppr_onPremUploadDriveItemContent(parametersJS) {
  return sppr_onprem_uploadDriveItemCommon(parametersJS, "onprem_upload_drive_item_content");
}

function sppr_onPremUploadDriveItemLargeContent(parametersJS) {
  return sppr_onprem_uploadDriveItemCommon(parametersJS, "onprem_upload_drive_item_large_content");
}

function sppr_onPremDownloadDriveItemContent(parametersJS) {
  var graphOperation = "onprem_download_drive_item_content";
  var requiredPermission = "SharePoint on-prem list read";

  try {
    var targetItemId = sppr_onprem_requireChildItemId(
      sppr_onprem_get(parametersJS, "itemId", ""),
      "Downloading drive root is not allowed. Provide a file item id."
    );
    var includeBase64 = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeContentBase64", "true"), true);
    var includeMetadata = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeMetadata", "true"), true);
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var itemContext = sppr_onprem_loadDriveItemContext(resolvedLibrary, targetItemId, includeRaw);
    if (itemContext.itemType === "1") {
      throw new java.lang.IllegalArgumentException("Downloading folder content is not allowed. Provide a file item id.");
    }

    var downloadResult = spop_httpRequest(
      "GET",
      spop_buildGetFileEndpoint(resolvedLibrary.siteBaseUrl, itemContext.itemPath) + "/$value",
      resolvedLibrary.auth,
      null,
      null,
      "*/*",
      null,
      true
    );
    spop_expectSuccess(downloadResult, "Download file content");

    var contentBytes = downloadResult.bytes === null || downloadResult.bytes === undefined ? sppr_onprem_emptyBytes() : downloadResult.bytes;
    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      itemId: targetItemId,
      includeContentBase64: includeBase64,
      includeMetadata: includeMetadata,
      includeRawItem: includeRaw
    });
    responseData.contentLength = contentBytes.length;
    responseData.contentType = spop_safeString(downloadResult.contentType);
    responseData.contentDisposition = spop_safeString(downloadResult.contentDisposition);
    responseData.contentLengthHeader = spop_safeString(downloadResult.contentLengthHeader);
    responseData.contentBase64 = includeBase64 ? spop_encodeBase64(contentBytes) : "";
    responseData.item = includeMetadata ? sppr_onprem_mapListItemToDriveItem(resolvedLibrary, itemContext.item, includeRaw) : {};
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onprem_moveOrRenameDriveItem(parametersJS, graphOperation) {
  var requiredPermission = "SharePoint on-prem list write";

  try {
    var targetItemId = sppr_onprem_requireChildItemId(
      sppr_onprem_get(parametersJS, "itemId", ""),
      "Updating drive root is not allowed. Provide a child item id."
    );
    var requestedDestinationParent = spop_safeString(sppr_onprem_get(parametersJS, "destinationParentItemId", ""));
    var requestedNewName = spop_safeString(sppr_onprem_get(parametersJS, "newName", ""));
    var requestedIfMatch = spop_safeString(sppr_onprem_get(parametersJS, "ifMatch", ""));
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var currentItemContext = sppr_onprem_loadDriveItemContext(resolvedLibrary, targetItemId, false);
    var targetInfo = sppr_onprem_resolveDriveTargetPath(
      resolvedLibrary,
      currentItemContext.parentPath,
      currentItemContext.itemName,
      requestedDestinationParent,
      requestedNewName,
      false
    );
    var targetName = targetInfo.targetName;
    if (spop_isBlank(targetName)) {
      throw new java.lang.IllegalArgumentException("newName cannot be empty");
    }
    if (spop_isBlank(requestedDestinationParent) && targetName === currentItemContext.itemName) {
      throw new java.lang.IllegalArgumentException("Provide destinationParentItemId and/or newName");
    }

    var contextInfo = spop_getContextDigest(resolvedLibrary.siteBaseUrl, resolvedLibrary.auth);
    var headers = {
      "X-RequestDigest": contextInfo.formDigestValue,
      "IF-MATCH": spop_isBlank(requestedIfMatch) ? "*" : requestedIfMatch
    };
    var itemEndpoint = currentItemContext.itemType === "1"
      ? spop_buildGetFolderEndpoint(resolvedLibrary.siteBaseUrl, currentItemContext.itemPath)
      : spop_buildGetFileEndpoint(resolvedLibrary.siteBaseUrl, currentItemContext.itemPath);
    spop_expectSuccess(
      spop_httpRequest(
        "POST",
        itemEndpoint + "/moveto(newurl='" + spop_encodeODataStringLiteral(targetInfo.targetPath) + "',flags=1)",
        resolvedLibrary.auth,
        sppr_onprem_emptyBytes(),
        "application/json;odata=verbose;charset=utf-8",
        "application/json;odata=verbose",
        headers,
        false
      ),
      "Move drive item"
    );

    var movedItem = sppr_onprem_findDriveItemByFileRef(resolvedLibrary, targetInfo.targetPath, includeRaw);
    if (movedItem === null) {
      throw new java.lang.RuntimeException("Unable to resolve moved item after move");
    }

    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      itemId: targetItemId,
      destinationParentItemId: requestedDestinationParent,
      resolvedDestinationParentItemId: targetInfo.resolvedDestinationParentItemId,
      newName: requestedNewName,
      ifMatchProvided: !spop_isBlank(requestedIfMatch),
      returnMovedItem: sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "returnMovedItem", "true"), true),
      selectFields: sppr_onprem_getSelectFields(parametersJS),
      includeRawItem: includeRaw
    });
    responseData.item = movedItem;
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremUpdateDriveItem(parametersJS) {
  try {
    var patchObject = {};
    var rawUpdateJson = spop_require("updateJson", sppr_onprem_get(parametersJS, "updateJson", ""));
    patchObject = JSON.parse(String(rawUpdateJson));
    if (patchObject === null || patchObject === undefined || patchObject instanceof Array) {
      throw new java.lang.IllegalArgumentException("updateJson must be a JSON object");
    }
    var forwarded = {};
    var key;
    for (key in parametersJS) {
      if (Object.prototype.hasOwnProperty.call(parametersJS, key)) {
        forwarded[key] = parametersJS[key];
      }
    }
    if (patchObject.name !== undefined && patchObject.name !== null) {
      forwarded.newName = patchObject.name;
    }
    if (patchObject.parentReference && patchObject.parentReference.id !== undefined && patchObject.parentReference.id !== null) {
      forwarded.destinationParentItemId = patchObject.parentReference.id;
    }
    if (spop_isBlank(spop_safeString(forwarded.newName)) && spop_isBlank(spop_safeString(forwarded.destinationParentItemId))) {
      throw new java.lang.IllegalArgumentException("Only updateJson.name and updateJson.parentReference.id are supported in on-prem mode");
    }
    return sppr_onprem_moveOrRenameDriveItem(forwarded, "onprem_update_drive_item");
  } catch (e) {
    return spop_responseError("onprem_update_drive_item", "onprem_unknown", "SharePoint on-prem list write", e);
  }
}

function sppr_onPremMoveDriveItem(parametersJS) {
  return sppr_onprem_moveOrRenameDriveItem(parametersJS, "onprem_move_drive_item");
}

function sppr_onPremCopyDriveItem(parametersJS) {
  var graphOperation = "onprem_copy_drive_item";
  var requiredPermission = "SharePoint on-prem list write";

  try {
    var targetItemId = sppr_onprem_requireChildItemId(
      sppr_onprem_get(parametersJS, "itemId", ""),
      "Copying drive root is not allowed. Provide a child item id."
    );
    var requestedDestinationParent = spop_safeString(sppr_onprem_get(parametersJS, "destinationParentItemId", ""));
    var requestedNewName = spop_safeString(sppr_onprem_get(parametersJS, "newName", ""));
    var shouldIncludeMonitor = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeMonitorResponse", "false"), false);
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var currentItemContext = sppr_onprem_loadDriveItemContext(resolvedLibrary, targetItemId, false);
    var targetInfo = sppr_onprem_resolveDriveTargetPath(
      resolvedLibrary,
      currentItemContext.parentPath,
      currentItemContext.itemName,
      requestedDestinationParent,
      requestedNewName,
      true
    );
    var normalizedName = targetInfo.targetName;
    if (spop_isBlank(normalizedName)) {
      throw new java.lang.IllegalArgumentException("newName cannot be empty");
    }
    var contextInfo = spop_getContextDigest(resolvedLibrary.siteBaseUrl, resolvedLibrary.auth);
    var headers = {
      "X-RequestDigest": contextInfo.formDigestValue
    };
    var itemEndpoint = currentItemContext.itemType === "1"
      ? spop_buildGetFolderEndpoint(resolvedLibrary.siteBaseUrl, currentItemContext.itemPath)
      : spop_buildGetFileEndpoint(resolvedLibrary.siteBaseUrl, currentItemContext.itemPath);
    spop_expectSuccess(
      spop_httpRequest(
        "POST",
        itemEndpoint + "/copyto(strnewurl='" + spop_encodeODataStringLiteral(targetInfo.targetPath) + "',boverwrite=true)",
        resolvedLibrary.auth,
        sppr_onprem_emptyBytes(),
        "application/json;odata=verbose;charset=utf-8",
        "application/json;odata=verbose",
        headers,
        false
      ),
      "Copy drive item"
    );

    var copiedItem = sppr_onprem_findDriveItemByFileRef(resolvedLibrary, targetInfo.targetPath, false);
    var monitorUrl = sppr_onprem_buildCopyMonitorUrl(copiedItem === null ? "" : spop_safeString(copiedItem.itemId), normalizedName);
    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      itemId: targetItemId,
      destinationParentItemId: requestedDestinationParent,
      resolvedDestinationParentItemId: targetInfo.resolvedDestinationParentItemId,
      newName: requestedNewName,
      includeMonitorResponse: shouldIncludeMonitor
    });
    responseData.copyRequest = {
      statusCode: 200,
      monitorUrl: monitorUrl,
      retryAfter: "",
      requestId: "",
      clientRequestId: "",
      normalizedName: normalizedName,
      accepted: true,
      copiedItemId: copiedItem === null ? "" : spop_safeString(copiedItem.itemId)
    };
    responseData.monitorUrl = monitorUrl;
    responseData.operationLocation = monitorUrl;
    responseData.monitor = shouldIncludeMonitor ? {
      monitorUrl: monitorUrl,
      statusCode: 200,
      inProgress: false,
      completed: true,
      operationState: "completed",
      retryAfter: "",
      location: "",
      requestId: "",
      clientRequestId: "",
      body: {
        status: "completed",
        provider: "onprem",
        synchronous: true,
        copiedItemId: copiedItem === null ? "" : spop_safeString(copiedItem.itemId),
        normalizedName: normalizedName
      },
      bodyText: ""
    } : {};
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremDeleteDriveItem(parametersJS) {
  var graphOperation = "onprem_delete_drive_item";
  var requiredPermission = "SharePoint on-prem list write";

  try {
    var targetItemId = sppr_onprem_requireChildItemId(
      sppr_onprem_get(parametersJS, "itemId", ""),
      "Deleting drive root is not allowed. Provide a child item id."
    );
    var includeSnapshot = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeDeletedItemSnapshot", "false"), false);
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var snapshotItem = null;
    if (includeSnapshot) {
      snapshotItem = sppr_onprem_expectOk(
        sppr_onPremGetDriveItem(sppr_onprem_withResolvedLibraryRequest({
          itemId: targetItemId,
          includeRawItem: includeRaw === true ? "true" : "false"
        }, resolvedLibrary)),
        "onprem_get_drive_item"
      ).item;
    }

    sppr_onprem_expectOk(
      sppr_onPremDeleteListItem(sppr_onprem_withResolvedLibraryRequest({
        itemId: targetItemId,
        includeDeletedItemSnapshot: "false",
        includeRawItem: "false"
      }, resolvedLibrary)),
      "onprem_delete_list_item"
    );

    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      itemId: targetItemId,
      includeDeletedItemSnapshot: includeSnapshot,
      selectFields: sppr_onprem_getSelectFields(parametersJS),
      includeRawItem: includeRaw
    });
    responseData.deleted = true;
    responseData.deleteHttpStatus = 200;
    responseData.item = includeSnapshot ? snapshotItem : {
      itemId: String(targetItemId)
    };
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremGetCopyDriveItemOperation(parametersJS) {
  var graphOperation = "onprem_get_copy_operation_status";
  var requiredPermission = "SharePoint on-prem list write";

  try {
    var requestedMonitorUrl = spop_require("monitorUrl", sppr_onprem_get(parametersJS, "monitorUrl", ""));
    if (requestedMonitorUrl.indexOf("onprem-copy://") !== 0) {
      throw new java.lang.IllegalArgumentException("Unsupported on-prem monitorUrl format");
    }
    var includeBody = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeMonitorBody", "true"), true);
    var parsedMonitor = sppr_onprem_parseCopyMonitorUrl(requestedMonitorUrl);
    var body = includeBody ? {
      status: "completed",
      provider: "onprem",
      synchronous: true,
      copiedItemId: parsedMonitor.copiedItemId,
      normalizedName: parsedMonitor.normalizedName
    } : {};

    return spop_responseOk(graphOperation, "onprem", requiredPermission, {
      monitorUrl: requestedMonitorUrl,
      statusCode: 200,
      inProgress: false,
      completed: true,
      operationState: "completed",
      retryAfter: "",
      location: "",
      requestId: "",
      clientRequestId: "",
      body: body,
      bodyText: ""
    });
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremListDriveItemVersions(parametersJS) {
  var graphOperation = "onprem_list_drive_item_versions";
  var requiredPermission = "SharePoint on-prem list read";

  try {
    var targetItemId = sppr_onprem_requireChildItemId(
      sppr_onprem_get(parametersJS, "itemId", ""),
      "Versions are available on file items, not root"
    );
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawVersion", "false"), false);
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var itemContext = sppr_onprem_loadDriveItemContext(resolvedLibrary, targetItemId, false);
    if (itemContext.itemType === "1") {
      throw new java.lang.IllegalArgumentException("Versions are available on file items, not folders");
    }

    var versionsPayload = spop_httpRequestJson(
      "GET",
      spop_buildGetFileEndpoint(resolvedLibrary.siteBaseUrl, itemContext.itemPath) + "/Versions",
      resolvedLibrary.auth,
      null,
      null
    );
    var versionsRoot = spop_parseVerbose(versionsPayload);
    var rawVersions = spop_extractArray(versionsRoot);
    var versions = [];
    var i;
    for (i = 0; i < rawVersions.length; i++) {
      var rawVersion = rawVersions[i];
      var normalizedVersion = {
        versionId: spop_safeString(rawVersion.VersionLabel),
        id: spop_safeString(rawVersion.VersionLabel),
        lastModifiedDateTime: sppr_onprem_normalizeDateTime(rawVersion.Created),
        size: rawVersion.Size === undefined || rawVersion.Size === null ? sppr_onprem_parseIntegerValue(rawVersion.Length, 0) : rawVersion.Size,
        publication: {
          level: rawVersion.IsCurrentVersion === true ? "current" : "published"
        },
        lastModifiedBy: {}
      };
      if (includeRaw) {
        normalizedVersion.raw = rawVersion;
      }
      versions.push(normalizedVersion);
    }

    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      top: spop_parseInteger("top", sppr_onprem_get(parametersJS, "top", ""), 200),
      maxPages: spop_parseInteger("maxPages", sppr_onprem_get(parametersJS, "maxPages", ""), 20),
      includeRawVersion: includeRaw
    });
    responseData.itemId = spop_safeString(targetItemId);
    responseData.pageCount = 1;
    responseData.itemCount = versions.length;
    responseData.nextLink = "";
    responseData.versions = versions;
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremRestoreDriveItemVersion(parametersJS) {
  var graphOperation = "onprem_restore_drive_item_version";
  var requiredPermission = "SharePoint on-prem list write";

  try {
    var targetItemId = sppr_onprem_requireChildItemId(
      sppr_onprem_get(parametersJS, "itemId", ""),
      "Restore is available on file items, not root"
    );
    var targetVersionId = spop_require("versionId", sppr_onprem_get(parametersJS, "versionId", ""));
    var includeItemSnapshot = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeItemSnapshot", "false"), false);
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var itemContext = sppr_onprem_loadDriveItemContext(resolvedLibrary, targetItemId, false);
    if (itemContext.itemType === "1") {
      throw new java.lang.IllegalArgumentException("Restore is available on file items, not folders");
    }

    var contextInfo = spop_getContextDigest(resolvedLibrary.siteBaseUrl, resolvedLibrary.auth);
    var restoreResult = spop_httpRequest(
      "POST",
      spop_buildGetFileEndpoint(resolvedLibrary.siteBaseUrl, itemContext.itemPath) + "/Versions/RestoreByLabel('" + spop_encodeODataStringLiteral(targetVersionId) + "')",
      resolvedLibrary.auth,
      sppr_onprem_emptyBytes(),
      "application/json;odata=verbose;charset=utf-8",
      "application/json;odata=verbose",
      {
        "X-RequestDigest": contextInfo.formDigestValue
      },
      false
    );
    spop_expectSuccess(restoreResult, "Restore drive item version");

    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {});
    responseData.itemId = spop_safeString(targetItemId);
    responseData.versionId = spop_safeString(targetVersionId);
    responseData.restoreRequest = {
      statusCode: restoreResult.statusCode,
      requestId: "",
      clientRequestId: ""
    };
    responseData.itemSnapshot = includeItemSnapshot ? sppr_onprem_expectOk(
      sppr_onPremGetDriveItem(sppr_onprem_withResolvedLibraryRequest({
        itemId: targetItemId,
        includeRawItem: includeRaw === true ? "true" : "false"
      }, resolvedLibrary)),
      "onprem_get_drive_item"
    ).item : {};
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremListGetItemsDelta(parametersJS) {
  var graphOperation = "onprem_list_get_items_delta";
  var requiredPermission = "SharePoint on-prem list read";

  try {
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var resolvedList = sppr_onprem_resolveListContext(parametersJS);
    var deltaData = sppr_onprem_queryListChanges(
      resolvedList,
      "",
      sppr_onprem_get(parametersJS, "deltaLink", ""),
      sppr_onprem_get(parametersJS, "deltaToken", ""),
      sppr_onprem_get(parametersJS, "top", ""),
      sppr_onprem_get(parametersJS, "maxPages", ""),
      includeRaw
    );
    var items = [];
    var i;
    for (i = 0; i < deltaData.items.length; i++) {
      items.push(deltaData.items[i]);
    }
    for (i = 0; i < deltaData.removedIds.length; i++) {
      items.push(sppr_onprem_buildRemovedListItem(deltaData.removedIds[i]));
    }

    return spop_responseOk(graphOperation, resolvedList.tokenMode, requiredPermission, {
      siteId: spop_safeString(resolvedList.siteId),
      siteHostname: spop_safeString(resolvedList.siteHostname),
      sitePath: spop_safeString(resolvedList.sitePath),
      listId: spop_safeString(resolvedList.listId),
      listName: spop_safeString(resolvedList.listName),
      listDisplayName: spop_safeString(resolvedList.listDisplayName),
      listWebUrl: spop_safeString(resolvedList.listWebUrl),
      listInfo: resolvedList.listInfo === undefined || resolvedList.listInfo === null ? {} : resolvedList.listInfo,
      listSharepointIds: resolvedList.listSharepointIds === undefined || resolvedList.listSharepointIds === null ? {} : resolvedList.listSharepointIds,
      query: {
        deltaToken: spop_safeString(sppr_onprem_get(parametersJS, "deltaToken", "")),
        deltaLink: spop_safeString(sppr_onprem_get(parametersJS, "deltaLink", "")),
        top: spop_parseInteger("top", sppr_onprem_get(parametersJS, "top", ""), 200),
        maxPages: spop_parseInteger("maxPages", sppr_onprem_get(parametersJS, "maxPages", ""), 20),
        selectFields: sppr_onprem_splitCsv(sppr_onprem_get(parametersJS, "selectFields", "")),
        expandFields: sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "expandFields", "true"), true),
        includeRawItem: includeRaw
      },
      pageCount: deltaData.pageCount,
      itemCount: items.length,
      nextLink: spop_safeString(deltaData.nextLink),
      deltaLink: spop_safeString(deltaData.deltaLink),
      items: items
    });
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}

function sppr_onPremListDriveItemsDelta(parametersJS) {
  var graphOperation = "onprem_list_drive_items_delta";
  var requiredPermission = "SharePoint on-prem list read";

  try {
    var includeRaw = sppr_onprem_toBoolean(sppr_onprem_get(parametersJS, "includeRawItem", "false"), false);
    var requestedParentId = spop_safeString(sppr_onprem_get(parametersJS, "parentItemId", ""));
    var normalizedParentId = spop_isBlank(requestedParentId) ? "root" : requestedParentId;
    var resolvedLibrary = sppr_onprem_resolveLibraryContext(parametersJS);
    var parentFolderServerRelativeUrl = sppr_onprem_resolveParentFolderServerRelativeUrl(resolvedLibrary, normalizedParentId);
    var deltaData = sppr_onprem_queryListChanges(
      resolvedLibrary,
      parentFolderServerRelativeUrl,
      sppr_onprem_get(parametersJS, "deltaLink", ""),
      sppr_onprem_get(parametersJS, "deltaToken", ""),
      sppr_onprem_get(parametersJS, "top", ""),
      sppr_onprem_get(parametersJS, "maxPages", ""),
      includeRaw
    );
    var items = [];
    var i;
    for (i = 0; i < deltaData.items.length; i++) {
      items.push(sppr_onprem_mapListItemToDriveItem(resolvedLibrary, deltaData.items[i], includeRaw));
    }
    for (i = 0; i < deltaData.removedIds.length; i++) {
      items.push(sppr_onprem_buildRemovedDriveItem(resolvedLibrary, deltaData.removedIds[i]));
    }

    var responseData = sppr_onprem_buildDriveResponse(resolvedLibrary, {
      parentItemId: normalizedParentId,
      deltaToken: spop_safeString(sppr_onprem_get(parametersJS, "deltaToken", "")),
      deltaLink: spop_safeString(sppr_onprem_get(parametersJS, "deltaLink", "")),
      top: spop_parseInteger("top", sppr_onprem_get(parametersJS, "top", ""), 200),
      maxPages: spop_parseInteger("maxPages", sppr_onprem_get(parametersJS, "maxPages", ""), 20),
      selectFields: sppr_onprem_getSelectFields(parametersJS),
      includeRawItem: includeRaw
    });
    responseData.parentItemId = normalizedParentId;
    sppr_onprem_applyPagedItems(responseData, deltaData.pageCount, items.length, deltaData.nextLink, items, deltaData.deltaLink);
    return spop_responseOk(graphOperation, resolvedLibrary.tokenMode, requiredPermission, responseData);
  } catch (e) {
    return spop_responseError(graphOperation, "onprem_unknown", requiredPermission, e);
  }
}
