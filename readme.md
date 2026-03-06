


# lib_Microsoft_Sharepoint

SharePoint connector for Convertigo (Online + On-Prem).
Online mode uses Microsoft Graph for site/list/drive operations.
On-prem mode exposes SharePoint REST APIs for site/list/file operations with digest-based writes.

[![Build, Deploy and Tests](https://github.com/convertigo/c8oprj-lib-sharepoint/actions/workflows/build-and-release.yml/badge.svg?branch=8.0.0.0)](https://github.com/convertigo/c8oprj-lib-sharepoint/actions/workflows/build-and-release.yml)

For more technical informations : [documentation](./project.md)

- [Installation](#installation)
- [Configuration Symbols](#configuration-symbols)
    - [On-Prem Symbols](#on-prem-symbols)
- [Authentication Model](#authentication-model)
- [On-Prem API Coverage](#on-prem-api-coverage)
- [Endpoint-Only Test Calls](#endpoint-only-test-calls)
- [Logical Test Plan Script](#logical-test-plan-script)
- [Typical Request Patterns](#typical-request-patterns)
- [Required Azure Permissions](#required-azure-permissions)
- [Permissions by Sequence](#permissions-by-sequence)
- [Known Limitations](#known-limitations)
- [Payload Examples](#payload-examples)
- [Sequences](#sequences)
    - [BuildGraphFlatJar](#buildgraphflatjar)
    - [CopyItem](#copyitem)
    - [CreateFolder](#createfolder)
    - [CreateGraphSubscription](#creategraphsubscription)
    - [CreateListItem](#createlistitem)
    - [CreateShareLink](#createsharelink)
    - [DeleteGraphSubscription](#deletegraphsubscription)
    - [DeleteItem](#deleteitem)
    - [DeleteItemPermission](#deleteitempermission)
    - [DeleteListItem](#deletelistitem)
    - [DownloadItemContent](#downloaditemcontent)
    - [ExecuteGraphBatch](#executegraphbatch)
    - [GetCopyItemOperation](#getcopyitemoperation)
    - [GetGraphAccessToken](#getgraphaccesstoken)
    - [GetItem](#getitem)
    - [GetListItem](#getlistitem)
    - [InviteItemRecipients](#inviteitemrecipients)
    - [ListGetItems](#listgetitems)
    - [ListGetItemsDelta](#listgetitemsdelta)
    - [ListItemPermissions](#listitempermissions)
    - [ListItems](#listitems)
    - [ListItemsDelta](#listitemsdelta)
    - [ListItemVersions](#listitemversions)
    - [ListSiteDrives](#listsitedrives)
    - [ListSiteLists](#listsitelists)
    - [MoveItem](#moveitem)
    - [ResolveLibrary](#resolvelibrary)
    - [ResolveList](#resolvelist)
    - [ResolveSite](#resolvesite)
    - [RestoreItemVersion](#restoreitemversion)
    - [UpdateItem](#updateitem)
    - [UpdateListItem](#updatelistitem)
    - [UploadItemContent](#uploaditemcontent)
    - [UploadItemLargeContent](#uploaditemlargecontent)


## Installation

1. In your Convertigo Studio click on ![](https://github.com/convertigo/convertigo/blob/develop/eclipse-plugin-studio/icons/studio/project_import.gif?raw=true "Import a project in treeview") to import a project in the treeview
2. In the import wizard

   ![](https://github.com/convertigo/convertigo/blob/develop/eclipse-plugin-studio/tomcat/webapps/convertigo/templates/ftl/project_import_wzd.png?raw=true "Import Project")
   
   paste the text below into the `Project remote URL` field:
   <table>
     <tr><td>Usage</td><td>Click the copy button at the end of the line</td></tr>
     <tr><td>To contribute</td><td>

     ```
     lib_Microsoft_Sharepoint=https://github.com/convertigo/c8oprj-lib-sharepoint.git:branch=8.0.0.0
     ```
     </td></tr>
     <tr><td>To simply use</td><td>

     ```
     lib_Microsoft_Sharepoint=https://github.com/convertigo/c8oprj-lib-sharepoint/archive/8.0.0.0.zip
     ```
     </td></tr>
    </table>
3. Click the `Finish` button. This will automatically import the __lib_Microsoft_Sharepoint__ project


## Configuration Symbols

These symbols can be set at project level and reused by all sequences.
In a standard deployment, they are configured once on the server and not passed on each request.

<table>
<tr><th>Symbol</th><th>Required</th><th>Secret</th><th>Purpose</th></tr>
<tr><td><code>${Microsoft_AzGraph.tenantId}</code></td><td>Yes (app-only)</td><td>No</td><td>Azure Entra tenant ID.</td></tr>
<tr><td><code>${Microsoft_AzGraph.clientId}</code></td><td>Yes (app-only)</td><td>No</td><td>Application (client) ID.</td></tr>
<tr><td><code>${Microsoft_AzGraph.clientSecret.secret}</code></td><td>Yes (app-only)</td><td>Yes</td><td>Application client secret.</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.provider}</code></td><td>No</td><td>No</td><td>Default backend provider for routed sequences: <code>graph</code> or <code>onprem</code> (default fallback is <code>graph</code>).</td></tr>
</table>

### On-Prem Symbols

These symbols are optional and used by routed sequences (`provider=onprem`) when values are not passed in the request.

<table>
<tr><th>Symbol</th><th>Required</th><th>Secret</th><th>Purpose</th></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.siteBaseUrl}</code></td><td>Recommended</td><td>No</td><td>Base URL of on-prem site (example: <code>https://sharepoint.local/sites/intranet</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.protocol}</code></td><td>Optional</td><td>No</td><td>Fallback protocol used by routed sequences when only host/path are provided (default <code>https</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.sitePath}</code></td><td>Optional</td><td>No</td><td>Default on-prem site path fallback used by routed sequences when <code>sitePath</code> is not passed (example: <code>/sites/test</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.listName}</code></td><td>Optional</td><td>No</td><td>Default on-prem list name fallback used by routed sequences when <code>listName</code> is not passed (example: <code>TestList</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.driveName}</code></td><td>Optional</td><td>No</td><td>Default on-prem library/drive name fallback used by routed sequences when <code>driveName</code> is not passed (example: <code>TestLibrary</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.username}</code></td><td>Optional</td><td>No</td><td>Technical username for on-prem auth (connector Basic/NTLM and sequence-level Basic fallback).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.password.secret}</code></td><td>Optional</td><td>Yes</td><td>Technical password for on-prem auth (connector Basic/NTLM and sequence-level Basic fallback).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.cookieHeader.secret}</code></td><td>Optional</td><td>Yes</td><td>Cookie header for forms auth (FedAuth/rtFa).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.server}</code></td><td>Optional</td><td>No</td><td>On-prem HTTP connector host (default <code>sharepoint.local</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.port}</code></td><td>Optional</td><td>No</td><td>On-prem HTTP connector port (default <code>443</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.https}</code></td><td>Optional</td><td>No</td><td>Use HTTPS for on-prem HTTP connector (default <code>true</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.baseDir}</code></td><td>Optional</td><td>No</td><td>On-prem HTTP connector base path (default <code>/</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.trustAll}</code></td><td>Optional</td><td>No</td><td>Trust all TLS certificates in on-prem HTTP connector (default <code>true</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.defaultSubDir}</code></td><td>Optional</td><td>No</td><td>Default transaction sub path for connector transactions (default <code>/_api/web</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.authenticationType}</code></td><td>Optional</td><td>No</td><td>Connector auth mode: <code>None</code>, <code>Basic</code>, <code>BasicPreemptive</code>, <code>NTLM</code> (default <code>None</code>).</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.ntlmDomain}</code></td><td>Optional</td><td>No</td><td>NTLM domain when <code>authenticationType=NTLM</code>.</td></tr>
<tr><td><code>${lib_Microsoft_Sharepoint.onPrem.http.ntlmTransportMode}</code></td><td>Optional</td><td>No</td><td>NTLM transport strategy: <code>connectorOnly</code>, <code>connectorThenHttpClient</code> (default), <code>httpClientThenConnector</code>, <code>httpClientOnly</code>.</td></tr>
</table>
Connector `sharepointOnPremHttp` uses shared credentials symbols `onPrem.username` and `onPrem.password.secret`.
For Claims/NTLM farms, set `onPrem.http.authenticationType=NTLM` and `onPrem.http.ntlmDomain` with those same shared credentials symbols.
When `onPrem.http.authenticationType` is `NTLM` (or `Basic`/`BasicPreemptive`), routed on-prem helper calls delegate authentication to the connector and do not force a sequence-level `Authorization: Basic` header.

## Authentication Model

- Default mode (recommended for backend use): rely on server-side symbols (`tenantId`, `clientId`, `clientSecret`) and do not pass credentials in calls.
- Delegated mode: pass `accessToken`; tenant/client/secret are ignored.
- Application override: leave `accessToken` empty and pass tenant/client/secret explicitly when needed.
- Routed public sequences switch backend with `provider` (`graph` or `onprem`), defaulting to `graph`.
- Graph implementation sequences are `Hidden` and `authenticatedContextRequired=true`.
- Sequence responses expose `tokenMode` (`delegated` or `application`) for diagnostics.
- On-prem routed operations support bearer token, basic auth, or cookie-based auth, and write operations automatically request SharePoint FormDigest via `/_api/contextinfo`.
- On-prem logic is executed directly from routed public sequences through shared JS helpers (no internal `OnPrem*` sequences).
- Connector `sharepointOnPremHttp` is used first by shared on-prem helper `spop_httpRequest` for routed NTLM calls, with optional HttpClient5 fallback depending on `onPrem.http.ntlmTransportMode`.
- On-prem HTTP execution strategy is configurable: `connectorOnly`, `connectorThenHttpClient` (default), `httpClientThenConnector`, `httpClientOnly`.

## On-Prem API Coverage

- Site/list resolvers: `ResolveSite`, `ResolveList`, `ResolveLibrary` with `provider=onprem`.
- On-prem discovery lists: `ListSiteLists`, `ListSiteDrives` with `provider=onprem`.
- List CRUD: `ListGetItems`, `GetListItem`, `CreateListItem`, `UpdateListItem`, `DeleteListItem` with `provider=onprem`.
- File/folder APIs: `ListItems`, `UploadItemContent`, `DownloadItemContent`, `DeleteItem`, `ListItemVersions`, `RestoreItemVersion`, `MoveItem`, `CopyItem`, `CreateFolder` with `provider=onprem`.
- Share block (`CreateShareLink`, `InviteItemRecipients`, `ListItemPermissions`, `DeleteItemPermission`) stays routed but returns explicit `not_supported_onprem` fallback in on-prem mode.
- Not covered in on-prem mode: Graph batch, Graph subscriptions.

## Endpoint-Only Test Calls

Minimal example with testcase injection:
```bash
curl 'http://localhost:18080/convertigo/projects/lib_Microsoft_Sharepoint/.json' \
  -H 'Content-Type: application/x-www-form-urlencoded' \
  --data-raw '__sequence=ResolveSite&__testcase=TC_ResolveSite'
```

Typical authenticated admin call (Convertigo Studio session):
```bash
curl 'http://localhost:18080/convertigo/projects/lib_Microsoft_Sharepoint/.json' \
  -H 'Content-Type: application/x-www-form-urlencoded' \
  -H 'Admin-Instance: <admin-instance-id>' \
  -H 'x-xsrf-token: <xsrf-token>' \
  -b 'JSESSIONID=<session-id>' \
  --data-raw '__sequence=ListGetItems&__testcase=TC_ListGetItems'
```

## Logical Test Plan Script

The project provides a full scenario runner that chains all SharePoint sequences and writes a JSON report.

Minimal usage:
```bash
python3 ./scripts/run_testcases.py
```

Useful overrides:
- `C8O_BASE_URL` to target another Convertigo endpoint.
- `C8O_PROJECT` to target another project name.
- `SITE_HOSTNAME`, `SITE_PATH`, `LIST_NAME`, `DRIVE_NAME` to target another SharePoint site context.
- `ACCESS_TOKEN` (delegated mode) or `AZ_TENANT_ID` + `AZ_CLIENT_ID` + `AZ_CLIENT_SECRET` (application override).
- `RUN_SHARE_OPERATIONS=false` to skip share/invite/permission deletion calls.
- `RUN_DESTRUCTIVE_CLEANUP=false` to skip delete calls.
- `REPORT_FILE` to change report output path (default `build/logical-test-plan-report.json`).

## Typical Request Patterns

- Site-first pattern: provide `siteHostname` + `sitePath`, and keep `siteId` empty.
- ID-first pattern: provide `siteId` / `listId` / `driveId` directly to skip resolver lookups.
- Drive resolution priority: `driveId` first, then list-based resolution (`listId`/`listName`), then `driveName`.
- List operations use list item identifiers (`itemId` numeric string), while drive operations use Graph drive item identifiers (opaque string).
- For large binary uploads, prefer `UploadItemLargeContent`; for smaller payloads, `UploadItemContent` is usually simpler.
- For routed on-prem calls (`provider=onprem`), drive-item id lookups are not available from Graph IDs: pass `onPremFileServerRelativeUrl` (or a server-relative `itemId`) for version and restore endpoints.

## Required Azure Permissions

Grant Microsoft Graph **Application** permissions, then click `Grant admin consent`.

<table>
<tr><th>Functional scope</th><th>Sequences</th><th>Graph permissions (Application)</th><th>Notes</th></tr>
<tr><td>Site and list discovery/read</td><td><code>ResolveSite</code>, <code>ResolveList</code>, <code>ListGetItems</code>, <code>GetListItem</code></td><td><code>Sites.Read.All</code> (or <code>Sites.ReadWrite.All</code>)</td><td>Use <code>Sites.ReadWrite.All</code> when list write operations are required.</td></tr>
<tr><td>Drive discovery/read</td><td><code>ResolveLibrary</code>, <code>ListItems</code>, <code>GetItem</code>, <code>DownloadItemContent</code>, <code>ListItemPermissions</code></td><td><code>Files.Read.All</code> + <code>Sites.Read.All</code> (or write variants)</td><td>Some tenants require both Files and Sites scopes for drive metadata traversal.</td></tr>
<tr><td>Delta and versions read</td><td><code>ListSiteLists</code>, <code>ListSiteDrives</code>, <code>ListGetItemsDelta</code>, <code>ListItemsDelta</code>, <code>ListItemVersions</code>, <code>GetCopyItemOperation</code></td><td><code>Sites.Read.All</code> and/or <code>Files.Read.All</code> (or write variants)</td><td>Delta links are incremental cursors that must be persisted by caller code.</td></tr>
<tr><td>List write operations</td><td><code>CreateListItem</code>, <code>UpdateListItem</code>, <code>DeleteListItem</code></td><td><code>Sites.ReadWrite.All</code></td><td>Targets SharePoint list items through <code>/sites/{siteId}/lists/{listId}</code>.</td></tr>
<tr><td>Drive write operations</td><td><code>CreateFolder</code>, <code>UpdateItem</code>, <code>DeleteItem</code>, <code>MoveItem</code>, <code>CopyItem</code>, <code>UploadItemContent</code>, <code>UploadItemLargeContent</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td><td><code>CopyItem</code> is asynchronous and returns a monitor URL.</td></tr>
<tr><td>Version restore</td><td><code>RestoreItemVersion</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td><td>Restoring versions can create additional versions depending on library retention policies.</td></tr>
<tr><td>Sharing and permission updates</td><td><code>CreateShareLink</code>, <code>InviteItemRecipients</code>, <code>DeleteItemPermission</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td><td>Invite and link creation can also be constrained by SharePoint external sharing policy.</td></tr>
<tr><td>Graph subscriptions and batch</td><td><code>CreateGraphSubscription</code>, <code>DeleteGraphSubscription</code>, <code>ExecuteGraphBatch</code></td><td>Depends on subscribed/batched resources</td><td>Subscription creation requires a reachable HTTPS notification endpoint.</td></tr>
<tr><td>Token helper</td><td><code>GetGraphAccessToken</code></td><td>No direct Graph API call</td><td>Acquires token from Entra ID; downstream sequence still needs Graph roles.</td></tr>
<tr><td>Local tooling</td><td><code>BuildGraphFlatJar</code></td><td>None</td><td>Builds local SDK JAR, no online call.</td></tr>
</table>

## Permissions by Sequence

<table>
<tr><th>Sequence</th><th>Graph permissions (Application)</th></tr>
<tr><td><code>BuildGraphFlatJar</code></td><td>None</td></tr>
<tr><td><code>GetGraphAccessToken</code></td><td>None direct; downstream usually <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ResolveSite</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ResolveList</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ResolveLibrary</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code> + <code>Files.Read.All</code> or <code>Files.ReadWrite.All</code></td></tr>
<tr><td><code>ListGetItems</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>GetListItem</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CreateListItem</code></td><td><code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>UpdateListItem</code></td><td><code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>DeleteListItem</code></td><td><code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListItems</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>GetItem</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>DownloadItemContent</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CreateFolder</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>UpdateItem</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>DeleteItem</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>MoveItem</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CopyItem</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>UploadItemContent</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>UploadItemLargeContent</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CreateShareLink</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>InviteItemRecipients</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListItemPermissions</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>DeleteItemPermission</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListSiteLists</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListSiteDrives</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListGetItemsDelta</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListItemsDelta</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListItemVersions</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>RestoreItemVersion</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>GetCopyItemOperation</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CreateGraphSubscription</code></td><td>Depends on subscribed resource</td></tr>
<tr><td><code>DeleteGraphSubscription</code></td><td>Depends on subscribed resource</td></tr>
<tr><td><code>ExecuteGraphBatch</code></td><td>Depends on batched requests</td></tr>
</table>

## Known Limitations

- If app roles do not include SharePoint scopes, Graph returns `accessDenied` even if token acquisition succeeds.
- `Sites.Selected` requires explicit grant on target sites; otherwise all calls return `403`.
- List item identifiers (usually numeric strings) and drive item identifiers (Graph opaque ids) are different and not interchangeable.
- `CopyItem` returns an asynchronous monitor URL; completion must be polled by client code.
- `CreateGraphSubscription` requires a publicly reachable HTTPS callback endpoint and Graph webhook validation handling.
- `ExecuteGraphBatch` permissions are the union of all operations included in `batchJson.requests`.

## Payload Examples

`fieldsJson` example for `CreateListItem` / `UpdateListItem`:
```json
{
  "Title": "Updated from Convertigo",
  "CustomText": "Hello from API"
}
```

`updateJson` example for `UpdateItem`:
```json
{
  "name": "Renamed_document.docx",
  "description": "Updated by Convertigo sequence"
}
```

`recipientsJson` example for `InviteItemRecipients`:
```json
[
  { "email": "user1@contoso.com" },
  { "email": "user2@contoso.com", "alias": "User Two" }
]
```
## Sequences

### BuildGraphFlatJar

Builds a flat JAR for Microsoft Graph Java SDK and stores it under .//libs (no Microsoft Graph permission required)

### CopyItem

Starts one SharePoint Online drive item copy operation through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>destinationParentItemId</td><td>Optional destination parent item id. Use root to copy under drive root.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeMonitorResponse</td><td>true performs one immediate GET on monitorUrl and includes returned payload/status.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier to copy.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>newName</td><td>Optional target file/folder name for the copied item.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### CreateFolder

Creates one SharePoint Online drive folder through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>conflictBehavior</td><td>Conflict behavior when folder exists. Allowed values are rename, replace or fail.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>folderName</td><td>Name of the folder to create.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under response.data.item.raw.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>parentItemId</td><td>Parent drive item id where folder will be created. Use root or leave empty for drive root.</td>
</tr>
<tr>
<td>returnCreatedItem</td><td>true reloads created item with optional field projection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### CreateGraphSubscription

Creates a Microsoft Graph webhook subscription (resource-dependent permissions).

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>changeType</td><td>Change type list, for example created,updated,deleted.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientState</td><td>Optional shared secret returned by Graph in notifications.</td>
</tr>
<tr>
<td>encryptionCertificate</td><td>Base64-encoded X.509 certificate public key used when includeResourceData=true.</td>
</tr>
<tr>
<td>encryptionCertificateId</td><td>Client-generated certificate id used when includeResourceData=true.</td>
</tr>
<tr>
<td>expirationDateTime</td><td>Subscription expiration in UTC ISO-8601 format.</td>
</tr>
<tr>
<td>includeResourceData</td><td>true enables encrypted resource data notifications (requires certificate inputs).</td>
</tr>
<tr>
<td>latestSupportedTlsVersion</td><td>Optional TLS version hint, for example v1_2.</td>
</tr>
<tr>
<td>notificationUrl</td><td>HTTPS endpoint receiving Graph validation and notifications.</td>
</tr>
<tr>
<td>resource</td><td>Graph resource to monitor, for example /sites/{site-id}/lists/{list-id}/items.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### CreateListItem

Creates one SharePoint list item through Microsoft Graph (online) or SharePoint REST (on-prem). (Graph permissions Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>contentTypeId</td><td>Optional SharePoint content type id sent in payload.contentType.id.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>fieldsJson</td><td>JSON object payload used as Graph fields map for list item creation.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph payload under item.raw.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier. If provided, listName is ignored.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve listId.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>returnCreatedItem</td><td>true reloads the created item with optional field projection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in returned item.fields.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### CreateShareLink

Creates one SharePoint drive item sharing link through Microsoft Graph (online), with graceful fallback when unsupported on-prem. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>expirationDateTime</td><td>Optional link expiration datetime in ISO-8601 format.</td>
</tr>
<tr>
<td>includeRawPermission</td><td>true includes raw Graph permission payload under response.data.permission.raw.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier.</td>
</tr>
<tr>
<td>linkType</td><td>Sharing link type (view, edit, embed, blocksDownload).</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>password</td><td>Optional password used when scope/type supports it.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>retainInheritedPermissions</td><td>Retains inherited permissions on first share when true.</td>
</tr>
<tr>
<td>scope</td><td>Sharing link scope (anonymous, organization, users).</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### DeleteGraphSubscription

Deletes one Microsoft Graph webhook subscription (resource-dependent permissions).

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>subscriptionId</td><td>Subscription identifier returned by CreateGraphSubscription.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### DeleteItem

Deletes one SharePoint Online drive item through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeDeletedItemSnapshot</td><td>true reads item before deletion and returns it in response.data.item.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under item.raw when snapshot is enabled.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier to delete.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select in returned snapshot.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### DeleteItemPermission

Deletes one SharePoint drive item permission through Microsoft Graph (online), with graceful fallback when unsupported on-prem. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeDeletedPermissionSnapshot</td><td>true reads permission before deletion and returns it in response.data.permission.</td>
</tr>
<tr>
<td>includeRawPermission</td><td>true includes raw Graph permission payload under permission.raw when snapshot is enabled.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier owning the permission to delete.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>permissionId</td><td>Target permission identifier to delete.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### DeleteListItem

Deletes one SharePoint list item through Microsoft Graph (online) or SharePoint REST (on-prem). (Graph permissions Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>includeDeletedItemSnapshot</td><td>true reads item before deletion and returns it in response.data.item.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph payload under item.raw when snapshot is enabled.</td>
</tr>
<tr>
<td>itemId</td><td>Target SharePoint list item identifier.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier. If provided, listName is ignored.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve listId.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in returned snapshot fields.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### DownloadItemContent

Downloads one SharePoint Online drive file content through Microsoft Graph. (Graph permissions Files.Read.All|Files.ReadWrite.All|Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeContentBase64</td><td>true includes downloaded bytes as base64 in response.data.contentBase64.</td>
</tr>
<tr>
<td>includeMetadata</td><td>true loads metadata in response.data.item.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under response.data.item.raw when metadata is loaded.</td>
</tr>
<tr>
<td>itemId</td><td>Target file drive item identifier to download.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### ExecuteGraphBatch

Executes one Microsoft Graph JSON batch request against /$batch endpoint (permissions depend on contained requests).

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>batchJson</td><td>Graph batch payload as JSON object with requests array.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>includeRawResponse</td><td>true includes full raw Graph batch response payload.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### GetCopyItemOperation

Reads status of one asynchronous copy/move monitor URL returned by Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>includeMonitorBody</td><td>true includes parsed monitor response body when present.</td>
</tr>
<tr>
<td>monitorUrl</td><td>Operation monitor URL returned by CopyDriveItem, usually from data.copyRequest.monitorUrl.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### GetGraphAccessToken

Resolves a delegated or app-only Microsoft Graph access token for SharePoint Online operations. (Graph permissions for downstream calls Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>includeTokenPayload</td><td>true decodes token payload claims for troubleshooting.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### GetItem

Retrieves one SharePoint Online drive item through Microsoft Graph. (Graph permissions Files.Read.All|Files.ReadWrite.All|Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under response.data.item.raw.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier. Use root or leave empty to retrieve drive root metadata.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### GetListItem

Retrieves one SharePoint list item through Microsoft Graph (online) or SharePoint REST (on-prem). (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph item payload under item.raw.</td>
</tr>
<tr>
<td>itemId</td><td>Target SharePoint list item identifier.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier. If provided, listName is ignored.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve listId.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in fields object.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### InviteItemRecipients

Invites recipients to one SharePoint drive item through Microsoft Graph (online), with graceful fallback when unsupported on-prem. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>expirationDateTime</td><td>Optional invitation expiration datetime in ISO-8601 format.</td>
</tr>
<tr>
<td>includeRawPermission</td><td>true includes raw Graph permission payload under response.data.permissions[*].raw.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>message</td><td>Optional invitation message sent by Graph.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>recipientEmails</td><td>Optional comma-separated list of recipient email addresses.</td>
</tr>
<tr>
<td>recipientsJson</td><td>Optional JSON array of recipients (objects with email/alias or simple email strings).</td>
</tr>
<tr>
<td>requireSignIn</td><td>true requires invited recipients to sign in.</td>
</tr>
<tr>
<td>retainInheritedPermissions</td><td>Retains inherited permissions on first invite when true.</td>
</tr>
<tr>
<td>rolesCsv</td><td>Roles as comma-separated values (read,write). Defaults to read.</td>
</tr>
<tr>
<td>sendInvitation</td><td>true asks Graph to send email invitations.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### ListGetItems

Lists SharePoint list items through Microsoft Graph (online) or SharePoint REST (on-prem), with filtering/pagination/projection. (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>expandFields</td><td>true includes fields expansion in Graph query.</td>
</tr>
<tr>
<td>filter</td><td>Optional OData filter expression.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph item payload under each result item.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier. If provided, listName is ignored.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve listId.</td>
</tr>
<tr>
<td>maxPages</td><td>Maximum number of Graph pages fetched before stopping pagination.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>orderBy</td><td>Optional OData order by expression.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in fields object.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>top</td><td>Maximum number of items requested per Graph page.</td>
</tr>
</table>

### ListGetItemsDelta

Lists list item changes through Microsoft Graph delta API. (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>deltaLink</td><td>Optional full @odata.deltaLink from previous call. If provided, other query parameters are ignored.</td>
</tr>
<tr>
<td>deltaToken</td><td>Optional delta token returned by previous call.</td>
</tr>
<tr>
<td>expandFields</td><td>true appends $expand=fields to include columns payload in each item.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph listItem payload under each result item.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier. If provided, listName is ignored.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve listId.</td>
</tr>
<tr>
<td>maxPages</td><td>Maximum number of pages fetched before stopping pagination.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph listItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>top</td><td>Maximum number of items requested per Graph page.</td>
</tr>
</table>

### ListItemPermissions

Lists permissions of one SharePoint drive item through Microsoft Graph (online), with graceful fallback when unsupported on-prem. (Graph permissions Files.Read.All|Files.ReadWrite.All|Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeRawPermission</td><td>true includes raw Graph permission payload under response.data.permissions[*].raw.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>maxPages</td><td>Maximum number of Graph pages fetched before stopping pagination.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph permission properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>top</td><td>Maximum number of permissions requested per Graph page.</td>
</tr>
</table>

### ListItems

Lists SharePoint Online drive items through Microsoft Graph with pagination and projection. (Graph permissions Files.Read.All|Files.ReadWrite.All|Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>filter</td><td>Optional OData filter expression.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under each result item.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>maxPages</td><td>Maximum number of Graph pages fetched before stopping pagination.</td>
</tr>
<tr>
<td>orderBy</td><td>Optional OData order by expression.</td>
</tr>
<tr>
<td>parentItemId</td><td>Parent drive item id. Use root or leave empty to list root children.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>top</td><td>Maximum number of items requested per Graph page.</td>
</tr>
</table>

### ListItemsDelta

Lists drive item changes through Microsoft Graph delta API. (Graph permissions Files.Read.All|Files.ReadWrite.All|Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>deltaLink</td><td>Optional full @odata.deltaLink from previous call. If provided, other query parameters are ignored.</td>
</tr>
<tr>
<td>deltaToken</td><td>Optional delta token returned by previous call.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under each result item.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>maxPages</td><td>Maximum number of pages fetched before stopping pagination.</td>
</tr>
<tr>
<td>parentItemId</td><td>Parent drive item id for scoped delta. Use root for drive root.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>top</td><td>Maximum number of items requested per Graph page.</td>
</tr>
</table>

### ListItemVersions

Lists versions of one drive item through Microsoft Graph with pagination. (Graph permissions Files.Read.All|Files.ReadWrite.All|Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeRawVersion</td><td>true includes raw Graph driveItemVersion payload under each result item.</td>
</tr>
<tr>
<td>itemId</td><td>Drive item identifier to inspect version history.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>maxPages</td><td>Maximum number of pages fetched before stopping pagination.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>top</td><td>Maximum number of versions requested per Graph page.</td>
</tr>
</table>

### ListSiteDrives

Lists SharePoint site drives through Microsoft Graph (online) or SharePoint REST (on-prem). (Graph permissions Files.Read.All|Files.ReadWrite.All|Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>includeRawDrive</td><td>true includes raw Graph drive payload under each result item.</td>
</tr>
<tr>
<td>maxPages</td><td>Maximum number of pages fetched before stopping pagination.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list properties to select.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>top</td><td>Maximum number of drives requested per Graph page.</td>
</tr>
</table>

### ListSiteLists

Lists SharePoint site lists through Microsoft Graph (online) or SharePoint REST (on-prem). (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>filter</td><td>Optional OData filter expression.</td>
</tr>
<tr>
<td>includeRawList</td><td>true includes raw Graph list payload under each result item.</td>
</tr>
<tr>
<td>maxPages</td><td>Maximum number of pages fetched before stopping pagination.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>orderBy</td><td>Optional OData order by expression.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>search</td><td>Optional OData search term.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list properties to select.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>top</td><td>Maximum number of lists requested per Graph page.</td>
</tr>
</table>

### MoveItem

Moves or renames one SharePoint Online drive item through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>destinationParentItemId</td><td>Optional destination parent item id. Use root to move under drive root.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>ifMatch</td><td>Optional ETag precondition. When set, move/rename occurs only if item ETag matches.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under response.data.item.raw.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier to move or rename.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>newName</td><td>Optional new file/folder name.</td>
</tr>
<tr>
<td>returnMovedItem</td><td>true reloads moved item with optional property selection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### ResolveLibrary

Resolves a SharePoint Online drive by site and drive identifiers using Microsoft Graph. (Graph permissions Sites.Read.All|Sites.ReadWrite.All|Files.Read.All|Files.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve a drive when driveId is not provided.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve the associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve the associated drive.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### ResolveList

Resolves a SharePoint list by site and list identifier/name using Microsoft Graph (online) or SharePoint REST (on-prem). (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier. If provided, listName is ignored.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve listId.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### ResolveSite

Resolves a SharePoint site by hostname and path using Microsoft Graph (online) or SharePoint REST (on-prem). (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname for target site, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path under hostname, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### RestoreItemVersion

Restores one historical version of a drive item through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeItemSnapshot</td><td>true fetches current item snapshot after restore operation.</td>
</tr>
<tr>
<td>itemId</td><td>Drive item identifier to restore from version history.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>versionId</td><td>Version identifier returned by ListDriveItemVersions.</td>
</tr>
</table>

### UpdateItem

Updates one SharePoint Online drive item through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under response.data.item.raw.</td>
</tr>
<tr>
<td>itemId</td><td>Target drive item identifier to update.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>returnUpdatedItem</td><td>true reloads updated item with optional property selection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
<tr>
<td>updateJson</td><td>JSON object payload used to update drive item metadata (for example name or parentReference).</td>
</tr>
</table>

### UpdateListItem

Updates one SharePoint list item through Microsoft Graph (online) or SharePoint REST (on-prem). (Graph permissions Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>cookieHeader</td><td>Optional Cookie header for on-prem forms authentication.</td>
</tr>
<tr>
<td>fieldsJson</td><td>JSON object payload used to patch list item fields.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph payload under item.raw.</td>
</tr>
<tr>
<td>itemId</td><td>Target SharePoint list item identifier.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier. If provided, listName is ignored.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve listId.</td>
</tr>
<tr>
<td>onPremPassword</td><td>Optional technical password for on-prem basic authentication.</td>
</tr>
<tr>
<td>onPremProtocol</td><td>URL scheme fallback for on-prem mode when siteBaseUrl is not provided.</td>
</tr>
<tr>
<td>onPremUsername</td><td>Optional technical username for on-prem basic authentication.</td>
</tr>
<tr>
<td>provider</td><td>Backend provider selection graph or onprem.</td>
</tr>
<tr>
<td>returnUpdatedItem</td><td>true reloads updated item with optional field projection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in returned item.fields.</td>
</tr>
<tr>
<td>siteBaseUrl</td><td>Optional SharePoint on-prem site base URL, for example https://sharepoint.local/sites/intranet.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### UploadItemContent

Uploads one SharePoint Online drive file content through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>contentBase64</td><td>Base64 content to upload.</td>
</tr>
<tr>
<td>contentType</td><td>Content type sent to Graph content endpoint.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>fileName</td><td>File name used when itemId is empty.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under response.data.item.raw.</td>
</tr>
<tr>
<td>itemId</td><td>Optional existing drive item id when uploading new content revision.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>parentItemId</td><td>Parent drive item id used when itemId is empty. Use root or leave empty for drive root.</td>
</tr>
<tr>
<td>returnUploadedItem</td><td>true reloads uploaded item with optional property selection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### UploadItemLargeContent

Uploads one SharePoint Online drive file using Graph upload session and chunk transfer. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>accessToken</td><td>Optional delegated bearer token. If provided, tenant/client/secret are ignored.</td>
</tr>
<tr>
<td>chunkSize</td><td>Chunk size in bytes (multiple of 327680). Default is 3276800.</td>
</tr>
<tr>
<td>clientId</td><td>Azure Entra application client id used for app-only token acquisition.</td>
</tr>
<tr>
<td>clientSecret</td><td>Azure Entra application client secret used for app-only token acquisition.</td>
</tr>
<tr>
<td>conflictBehavior</td><td>Conflict behavior when creating a new file by path (replace, rename, fail).</td>
</tr>
<tr>
<td>contentBase64</td><td>Base64 content uploaded through a Graph upload session.</td>
</tr>
<tr>
<td>contentType</td><td>Content type sent for each upload chunk.</td>
</tr>
<tr>
<td>driveId</td><td>Optional drive identifier. If provided, driveName/listId/listName are ignored.</td>
</tr>
<tr>
<td>driveName</td><td>Drive name used to resolve driveId when driveId is not provided.</td>
</tr>
<tr>
<td>fileName</td><td>File name used when itemId is empty.</td>
</tr>
<tr>
<td>includeRawItem</td><td>true includes raw Graph driveItem payload under response.data.item.raw.</td>
</tr>
<tr>
<td>includeSessionDetails</td><td>true includes uploadUrl and upload session metadata in the response.</td>
</tr>
<tr>
<td>itemId</td><td>Optional existing drive item id when uploading new content revision.</td>
</tr>
<tr>
<td>listId</td><td>Optional list identifier used to resolve associated drive.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve associated drive.</td>
</tr>
<tr>
<td>parentItemId</td><td>Parent drive item id used when itemId is empty. Use root or leave empty for drive root.</td>
</tr>
<tr>
<td>returnUploadedItem</td><td>true reloads uploaded item with optional property selection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of Graph driveItem properties to select.</td>
</tr>
<tr>
<td>siteHostname</td><td>SharePoint Online hostname when resolving siteId, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>siteId</td><td>Optional site identifier. If empty, siteHostname and sitePath are used to resolve it.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path when resolving siteId, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

