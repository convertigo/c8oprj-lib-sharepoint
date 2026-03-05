


# lib_Microsoft_Sharepoint

SharePoint Online connector for Convertigo.
Uses Microsoft Graph to resolve sites, lists and drives, then perform read/write/share operations on SharePoint Online content.

	[![Build, Deploy and Tests](https://github.com/convertigo/c8oprj-lib-sharepoint/actions/workflows/build-and-release.yml/badge.svg?branch=8.0.0.0)](https://github.com/convertigo/c8oprj-lib-sharepoint/actions/workflows/build-and-release.yml)

For more technical informations : [documentation](./project.md)

- [Installation](#installation)
- [Configuration Symbols](#configuration-symbols)
- [Authentication Model](#authentication-model)
- [Endpoint-Only Test Calls](#endpoint-only-test-calls)
- [Logical Test Plan Script](#logical-test-plan-script)
- [Typical Request Patterns](#typical-request-patterns)
- [Required Azure Permissions](#required-azure-permissions)
- [Permissions by Sequence](#permissions-by-sequence)
- [Known Limitations](#known-limitations)
- [Payload Examples](#payload-examples)
- [Sequences](#sequences)
    - [BuildGraphFlatJar](#buildgraphflatjar)
    - [CopyDriveItem](#copydriveitem)
    - [CreateDriveFolder](#createdrivefolder)
    - [CreateDriveItemShareLink](#createdriveitemsharelink)
    - [CreateListItem](#createlistitem)
    - [DeleteDriveItem](#deletedriveitem)
    - [DeleteDriveItemPermission](#deletedriveitempermission)
    - [DeleteListItem](#deletelistitem)
    - [DownloadDriveItemContent](#downloaddriveitemcontent)
    - [GetDriveItem](#getdriveitem)
    - [GetGraphAccessToken](#getgraphaccesstoken)
    - [GetListItem](#getlistitem)
    - [InviteDriveItemRecipients](#invitedriveitemrecipients)
    - [ListDriveItemPermissions](#listdriveitempermissions)
    - [ListDriveItems](#listdriveitems)
    - [ListGetItems](#listgetitems)
    - [MoveDriveItem](#movedriveitem)
    - [ResolveDrive](#resolvedrive)
    - [ResolveList](#resolvelist)
    - [ResolveSite](#resolvesite)
    - [UpdateDriveItem](#updatedriveitem)
    - [UpdateListItem](#updatelistitem)
    - [UploadDriveItemContent](#uploaddriveitemcontent)
    - [UploadDriveItemLargeContent](#uploaddriveitemlargecontent)


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
</table>

## Authentication Model

- Default mode (recommended for backend use): rely on server-side symbols (`tenantId`, `clientId`, `clientSecret`) and do not pass credentials in calls.
- Delegated mode: pass `accessToken`; tenant/client/secret are ignored.
- Application override: leave `accessToken` empty and pass tenant/client/secret explicitly when needed.
- All online Graph sequences are `Hidden` and `authenticatedContextRequired=true`.
- Sequence responses expose `tokenMode` (`delegated` or `application`) for diagnostics.

## Endpoint-Only Test Calls

Minimal example with testcase injection:
```bash
curl 'http://localhost:18080/convertigo/projects/lib_Sharepoint/.json' \
  -H 'Content-Type: application/x-www-form-urlencoded' \
  --data-raw '__sequence=ResolveSite&__testcase=TC_ResolveSite'
```

Typical authenticated admin call (Convertigo Studio session):
```bash
curl 'http://localhost:18080/convertigo/projects/lib_Sharepoint/.json' \
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
- `RUN_BUILD_JAR=false` to skip `BuildGraphFlatJar`.
- `RUN_SHARE_OPERATIONS=false` to skip share/invite/permission deletion calls.
- `RUN_DESTRUCTIVE_CLEANUP=false` to skip delete calls.
- `REPORT_FILE` to change report output path (default `build/logical-test-plan-report.json`).

## Typical Request Patterns

- Site-first pattern: provide `siteHostname` + `sitePath`, and keep `siteId` empty.
- ID-first pattern: provide `siteId` / `listId` / `driveId` directly to skip resolver lookups.
- Drive resolution priority: `driveId` first, then list-based resolution (`listId`/`listName`), then `driveName`.
- List operations use list item identifiers (`itemId` numeric string), while drive operations use Graph drive item identifiers (opaque string).
- For large binary uploads, prefer `UploadDriveItemLargeContent`; for smaller payloads, `UploadDriveItemContent` is usually simpler.

## Required Azure Permissions

Grant Microsoft Graph **Application** permissions, then click `Grant admin consent`.

<table>
<tr><th>Functional scope</th><th>Sequences</th><th>Graph permissions (Application)</th><th>Notes</th></tr>
<tr><td>Site and list discovery/read</td><td><code>ResolveSite</code>, <code>ResolveList</code>, <code>ListGetItems</code>, <code>GetListItem</code></td><td><code>Sites.Read.All</code> (or <code>Sites.ReadWrite.All</code>)</td><td>Use <code>Sites.ReadWrite.All</code> when list write operations are required.</td></tr>
<tr><td>Drive discovery/read</td><td><code>ResolveDrive</code>, <code>ListDriveItems</code>, <code>GetDriveItem</code>, <code>DownloadDriveItemContent</code>, <code>ListDriveItemPermissions</code></td><td><code>Files.Read.All</code> + <code>Sites.Read.All</code> (or write variants)</td><td>Some tenants require both Files and Sites scopes for drive metadata traversal.</td></tr>
<tr><td>List write operations</td><td><code>CreateListItem</code>, <code>UpdateListItem</code>, <code>DeleteListItem</code></td><td><code>Sites.ReadWrite.All</code></td><td>Targets SharePoint list items through <code>/sites/{siteId}/lists/{listId}</code>.</td></tr>
<tr><td>Drive write operations</td><td><code>CreateDriveFolder</code>, <code>UpdateDriveItem</code>, <code>DeleteDriveItem</code>, <code>MoveDriveItem</code>, <code>CopyDriveItem</code>, <code>UploadDriveItemContent</code>, <code>UploadDriveItemLargeContent</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td><td><code>CopyDriveItem</code> is asynchronous and returns a monitor URL.</td></tr>
<tr><td>Sharing and permission updates</td><td><code>CreateDriveItemShareLink</code>, <code>InviteDriveItemRecipients</code>, <code>DeleteDriveItemPermission</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td><td>Invite and link creation can also be constrained by SharePoint external sharing policy.</td></tr>
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
<tr><td><code>ResolveDrive</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code> + <code>Files.Read.All</code> or <code>Files.ReadWrite.All</code></td></tr>
<tr><td><code>ListGetItems</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>GetListItem</code></td><td><code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CreateListItem</code></td><td><code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>UpdateListItem</code></td><td><code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>DeleteListItem</code></td><td><code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListDriveItems</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>GetDriveItem</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>DownloadDriveItemContent</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CreateDriveFolder</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>UpdateDriveItem</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>DeleteDriveItem</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>MoveDriveItem</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CopyDriveItem</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>UploadDriveItemContent</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>UploadDriveItemLargeContent</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>CreateDriveItemShareLink</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>InviteDriveItemRecipients</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>ListDriveItemPermissions</code></td><td><code>Files.Read.All</code> or <code>Files.ReadWrite.All</code> + <code>Sites.Read.All</code> or <code>Sites.ReadWrite.All</code></td></tr>
<tr><td><code>DeleteDriveItemPermission</code></td><td><code>Files.ReadWrite.All</code> + <code>Sites.ReadWrite.All</code></td></tr>
</table>

## Known Limitations

- If app roles do not include SharePoint scopes, Graph returns `accessDenied` even if token acquisition succeeds.
- `Sites.Selected` requires explicit grant on target sites; otherwise all calls return `403`.
- List item identifiers (usually numeric strings) and drive item identifiers (Graph opaque ids) are different and not interchangeable.
- `CopyDriveItem` returns an asynchronous monitor URL; completion must be polled by client code.

## Payload Examples

`fieldsJson` example for `CreateListItem` / `UpdateListItem`:
```json
{
  "Title": "Updated from Convertigo",
  "CustomText": "Hello from API"
}
```

`updateJson` example for `UpdateDriveItem`:
```json
{
  "name": "Renamed_document.docx",
  "description": "Updated by Convertigo sequence"
}
```

`recipientsJson` example for `InviteDriveItemRecipients`:
```json
[
  { "email": "user1@contoso.com" },
  { "email": "user2@contoso.com", "alias": "User Two" }
]
```
## Sequences

### BuildGraphFlatJar

Builds a flat JAR for Microsoft Graph Java SDK and stores it under .//libs (no Microsoft Graph permission required)

### CopyDriveItem

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

### CreateDriveFolder

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

### CreateDriveItemShareLink

Creates one SharePoint Online drive item sharing link through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

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
<td>retainInheritedPermissions</td><td>Retains inherited permissions on first share when true.</td>
</tr>
<tr>
<td>scope</td><td>Sharing link scope (anonymous, organization, users).</td>
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

### CreateListItem

Creates a SharePoint Online list item through Microsoft Graph. (Graph permissions Sites.ReadWrite.All)

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
<td>returnCreatedItem</td><td>true reloads the created item with optional field projection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in returned item.fields.</td>
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

### DeleteDriveItem

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

### DeleteDriveItemPermission

Deletes one SharePoint Online drive item permission through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

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

### DeleteListItem

Deletes one SharePoint Online list item through Microsoft Graph. (Graph permissions Sites.ReadWrite.All)

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
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in returned snapshot fields.</td>
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

### DownloadDriveItemContent

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

### GetDriveItem

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

### GetListItem

Retrieves one SharePoint Online list item through Microsoft Graph. (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

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
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in fields object.</td>
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

### InviteDriveItemRecipients

Invites recipients to one SharePoint Online drive item through Microsoft Graph. (Graph permissions Files.ReadWrite.All|Sites.ReadWrite.All)

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

### ListDriveItemPermissions

Lists permissions of one SharePoint Online drive item through Microsoft Graph. (Graph permissions Files.Read.All|Files.ReadWrite.All|Sites.Read.All|Sites.ReadWrite.All)

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
<td>selectFields</td><td>Optional comma-separated list of Graph permission properties to select.</td>
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
<td>top</td><td>Maximum number of permissions requested per Graph page.</td>
</tr>
</table>

### ListDriveItems

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

### ListGetItems

Lists SharePoint Online list items through Microsoft Graph with OData pagination and field projection. (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

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
<td>orderBy</td><td>Optional OData order by expression.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in fields object.</td>
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

### MoveDriveItem

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

### ResolveDrive

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

Resolves a SharePoint Online list by site and list name using Microsoft Graph. (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

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
<td>listId</td><td>Optional list identifier. If provided, listName is ignored.</td>
</tr>
<tr>
<td>listName</td><td>List display name or internal name used to resolve listId.</td>
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

### ResolveSite

Resolves a SharePoint Online site by hostname and site path using Microsoft Graph. (Graph permissions Sites.Read.All|Sites.ReadWrite.All)

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
<td>siteHostname</td><td>SharePoint Online hostname for target site, for example contoso.sharepoint.com.</td>
</tr>
<tr>
<td>sitePath</td><td>Site path under hostname, for example /sites/engineering.</td>
</tr>
<tr>
<td>tenantId</td><td>Azure Entra tenant id used for app-only token acquisition.</td>
</tr>
</table>

### UpdateDriveItem

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

Updates one SharePoint Online list item through Microsoft Graph. (Graph permissions Sites.ReadWrite.All)

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
<td>returnUpdatedItem</td><td>true reloads updated item with optional field projection before returning.</td>
</tr>
<tr>
<td>selectFields</td><td>Optional comma-separated list of field internal names projected in returned item.fields.</td>
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

### UploadDriveItemContent

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

### UploadDriveItemLargeContent

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


