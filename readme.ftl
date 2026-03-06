<#-- This FTL template helps generating the readme.md file of your project -->
<#-- see FTL language documentation : https://freemarker.apache.org/docs/index.html -->

<#-- GLOBALS -->
<#global lineBreak = settings.lineBreak />
<#global locale = "US" />
<#global dictionnary = {
		"installation":	{"US": "Installation"			, "FR": "Installation"},
		"more.info": 	{"US": "For more technical informations"	, "FR": "Pour plus d'informations techniques"},
		"connectors": 	{"US": "Connectors"				, "FR": "Connecteurs"},
		"transactions": {"US": "Transactions"			, "FR": "Transactions"},
		"sequences": 	{"US": "Sequences"				, "FR": "Séquences"},
		"references": 	{"US": "References"				, "FR": "Références"},
		"urlmapper": 	{"US": "Rest Web Service"		, "FR": "Service Web REST"},
		"mappings": 	{"US": "Mappings"				, "FR": "Mappages"},
		"operations": 	{"US": "Operations"				, "FR": "Operations"},
		"parameters": 	{"US": "Parameters"				, "FR": "Paramètres"},
		"mobileapp": 	{"US": "Mobile Application"		, "FR": "Application Mobile"},
		"mobilelib": 	{"US": "Mobile Library"			, "FR": "Librairie Mobile"},
		"pages": 		{"US": "Pages"					, "FR": "Pages"},
		"actions": 		{"US": "Shared Actions"			, "FR": "Actions partagées"},
		"components": 	{"US": "Shared Components"		, "FR": "Composants partagés"},
		"variables": 	{"US": "variables"				, "FR": "variables"},
		"events": 		{"US": "events"					, "FR": "évènements"}
	}
/>
<#-- please modify the global show values as needed -->
<#global show = {
	"toc"			: true,
	"installation"	: true,
	
	"connectors"	: false,
	"transactions"	: true,
	"sequences"		: !has(project, "urlmapper") && !has(project, "mobileapp"),
	
	"references"	: false,
	
	"urlmapper"		: true,
	"mappings"		: true,
	"operations"	: true,
	"parameters"	: true,
	
	"mobileapp"		: true,
	"pages"			: !project.name?starts_with("lib_"),
	"actions"		: true,
	"components"	: true,
	
	"variables"		: true,
	"events"		: true
	} 
/>

<#-- FUNCTIONS -->
<#-- on: returns the show flag for the given key -->
<#function on key>
  <#return show[key]?? && show[key]>
</#function>

<#-- on: test if given dbo has the given key with non empty size -->
<#function has dbo key>
  <#return dbo[key]?? && (dbo[key]?size > 0) >
</#function>

<#-- anchor: generates an anchor link for the given text -->
<#function anchor anchors text>
  <#assign a = ""+ text?lower_case?replace(" ", "-")?replace("/", "")>
  <#if anchors?seq_contains(a)>
  	<#assign f = anchors?filter(s -> s?matches(""+ a + "-(\\d+)"))>
  	<#assign a = ""+ a + "-" + (f?size+1)>
  </#if>
  <#assign anchors += [""+a]>
  <#return a>
</#function>

<#-- on: returns the dictionnary value for the given key -->
<#function help key>
  <#if has(dictionnary, key)>
    <#return dictionnary[key][locale]!key>
  </#if>
  <#return key>
</#function>

<#-- toHttpsRemoteUrl: rewrites git SSH GitHub remote URLs to HTTPS -->
<#function toHttpsRemoteUrl url>
  <#assign value = (url!"")?trim>
  <#assign value = value?replace("=ssh://git@github.com/", "=https://github.com/")>
  <#assign value = value?replace("=git@github.com:", "=https://github.com/")>
  <#assign value = value?replace("ssh://git@github.com/", "https://github.com/")>
  <#assign value = value?replace("git@github.com:", "https://github.com/")>
  <#return value>
</#function>

<#-- MACROS -->
<#-- header: generates a header with given text as heading and add it to TOC with its anchor link -->
<#macro header toc anchors heading text>
${heading} ${text}${lineBreak}
<#assign a = anchor(anchors, text)>
<#if (heading?keep_before_last("#")?length > 0)>
<#assign toc += "" + heading?keep_before_last("##")?replace("#","    ") + "-" + " ["+text+"](#"+ a +")" + lineBreak>
</#if>
</#macro>

<#-- comment: add given text -->
<#macro comment text>
<#if (text?length > 0) >
${text}${lineBreak}
</#if>
</#macro>

<#-- table: generates a table with given headers and rows -->
<#macro table title headers rows>
<#if (rows?size > 0)>
${title}${lineBreak}
<table>
<tr>
<#list headers as header><th>${header}</th></#list>
</tr>
<#list rows as i>
<tr>
<#list headers as header><td>${i[header]}</td></#list>
</tr>
</#list>
</table>${lineBreak}
</#if>
</#macro>

<#-- installation : add project installation instructions if any -->
<#macro installation>
<#if locale == "US">
1. In your Convertigo Studio click on ![](https://github.com/convertigo/convertigo/blob/develop/eclipse-plugin-studio/icons/studio/project_import.gif?raw=true "Import a project in treeview") to import a project in the treeview
2. In the import wizard

   ![](https://github.com/convertigo/convertigo/blob/develop/eclipse-plugin-studio/tomcat/webapps/convertigo/templates/ftl/project_import_wzd.png?raw=true "Import Project")
   
   paste the text below into the `Project remote URL` field:
   <table>
     <tr><td>Usage</td><td>Click the copy button at the end of the line</td></tr>
     <tr><td>To contribute</td><td>${lineBreak}
     ```
     ${toHttpsRemoteUrl(project.contributeUrl)}
     ```
     </td></tr>
     <tr><td>To simply use</td><td>${lineBreak}
     ```
     ${toHttpsRemoteUrl(project.usageUrl)}
     ```
     </td></tr>
    </table>
3. Click the `Finish` button. This will automatically import the __${project.name}__ project
</#if>
<#if locale == "FR">
1. Dans votre Studio Convertigo, cliquez sur ![](https://github.com/convertigo/convertigo/blob/develop/eclipse-plugin-studio/icons/studio/project_import.gif?raw=true "Import a project in treeview") pour importer un projet dans l'arborescence
2. Dans l'assistant d'importation

   ![](https://github.com/convertigo/convertigo/blob/develop/eclipse-plugin-studio/tomcat/webapps/convertigo/templates/ftl/project_import_wzd.png?raw=true "Import Project")
   
   collez le texte ci-dessous dans le champ `Project remote URL`:
   <table>
     <tr><td>Usage</td><td>Cliquez sur le bouton de copie en fin de ligne</td></tr>
     <tr><td>Pour contribuer</td><td>${lineBreak}
     ```
     ${lineBreak}${toHttpsRemoteUrl(project.contributeUrl)}
     ```
     </td></tr>
     <tr><td>Pour simplement utiliser</td><td>${lineBreak}
     ```
     ${lineBreak}${toHttpsRemoteUrl(project.usageUrl)}
     ```
     </td></tr>
    </table>
3. Cliquez sur le bouton `Finish`. Cela importera automatiquement le projet __${project.name}__
</#if>
${lineBreak}
</#macro>

<#-- DEFAULT PROJECT TEMPLATE -->

<#-- anchors variable for TOC : do not modify -->
<#assign anchors = [""]>
<#-- toc variable : do not modify -->
<#assign toc = "">

<#-- Please modify below templates as needed -->

<#-- intro variable : add project header and comment -->
<#assign intro>
	<@header toc=toc anchors=anchors heading="#" text=project.label />
	<@comment text=project.comment />
[![Build, Deploy and Tests](https://github.com/convertigo/c8oprj-lib-sharepoint/actions/workflows/build-and-release.yml/badge.svg?branch=8.0.0.0)](https://github.com/convertigo/c8oprj-lib-sharepoint/actions/workflows/build-and-release.yml)
	<#-- you can add your text or own macro call here to add something -->
	<#--
	This is text i want to add after the project comment
	<@my_own_macro my_var='xxxx xxxxx xxxxx'>
	-->
</#assign>

<#-- content variable : add project sub-beans header and comment -->
<#-- you can add your text or own macro call anywhere -->
<#assign content>
<#if on("installation") && (project.url?length > 0) && (project.url != project.name)>
	<@header toc=toc anchors=anchors heading="##" text=help("installation") />
	<@installation />
</#if>
<#if locale == "US">
	<@header toc=toc anchors=anchors heading="##" text="Configuration Symbols" />
These symbols can be set at project level and reused by all sequences.
In a standard deployment, they are configured once on the server and not passed on each request.

<table>
<tr><th>Symbol</th><th>Required</th><th>Secret</th><th>Purpose</th></tr>
<tr><td><code><#noparse>${Microsoft_AzGraph.tenantId}</#noparse></code></td><td>Yes (app-only)</td><td>No</td><td>Azure Entra tenant ID.</td></tr>
<tr><td><code><#noparse>${Microsoft_AzGraph.clientId}</#noparse></code></td><td>Yes (app-only)</td><td>No</td><td>Application (client) ID.</td></tr>
<tr><td><code><#noparse>${Microsoft_AzGraph.clientSecret.secret}</#noparse></code></td><td>Yes (app-only)</td><td>Yes</td><td>Application client secret.</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.provider}</#noparse></code></td><td>No</td><td>No</td><td>Default backend provider for routed sequences: <code>graph</code> or <code>onprem</code> (default fallback is <code>graph</code>).</td></tr>
</table>

<@header toc=toc anchors=anchors heading="###" text="On-Prem Symbols" />
These symbols are optional and used by routed sequences (`provider=onprem`) when values are not passed in the request.

<table>
<tr><th>Symbol</th><th>Required</th><th>Secret</th><th>Purpose</th></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.siteBaseUrl}</#noparse></code></td><td>Recommended</td><td>No</td><td>Base URL of on-prem site (example: <code>https://sharepoint.local/sites/intranet</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.protocol}</#noparse></code></td><td>Optional</td><td>No</td><td>Fallback protocol used by routed sequences when only host/path are provided (default <code>https</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.sitePath}</#noparse></code></td><td>Optional</td><td>No</td><td>Default on-prem site path fallback used by routed sequences when <code>sitePath</code> is not passed (example: <code>/sites/test</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.listName}</#noparse></code></td><td>Optional</td><td>No</td><td>Default on-prem list name fallback used by routed sequences when <code>listName</code> is not passed (example: <code>TestList</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.driveName}</#noparse></code></td><td>Optional</td><td>No</td><td>Default on-prem library/drive name fallback used by routed sequences when <code>driveName</code> is not passed (example: <code>TestLibrary</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.username}</#noparse></code></td><td>Optional</td><td>No</td><td>Technical username for on-prem auth (connector Basic/NTLM and sequence-level Basic fallback).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.password.secret}</#noparse></code></td><td>Optional</td><td>Yes</td><td>Technical password for on-prem auth (connector Basic/NTLM and sequence-level Basic fallback).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.cookieHeader.secret}</#noparse></code></td><td>Optional</td><td>Yes</td><td>Cookie header for forms auth (FedAuth/rtFa).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.server}</#noparse></code></td><td>Optional</td><td>No</td><td>On-prem HTTP connector host (default <code>sharepoint.local</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.port}</#noparse></code></td><td>Optional</td><td>No</td><td>On-prem HTTP connector port (default <code>443</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.https}</#noparse></code></td><td>Optional</td><td>No</td><td>Use HTTPS for on-prem HTTP connector (default <code>true</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.baseDir}</#noparse></code></td><td>Optional</td><td>No</td><td>On-prem HTTP connector base path (default <code>/</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.trustAll}</#noparse></code></td><td>Optional</td><td>No</td><td>Trust all TLS certificates in on-prem HTTP connector (default <code>true</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.defaultSubDir}</#noparse></code></td><td>Optional</td><td>No</td><td>Default transaction sub path for connector transactions (default <code>/_api/web</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.authenticationType}</#noparse></code></td><td>Optional</td><td>No</td><td>Connector auth mode: <code>None</code>, <code>Basic</code>, <code>BasicPreemptive</code>, <code>NTLM</code> (default <code>None</code>).</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.ntlmDomain}</#noparse></code></td><td>Optional</td><td>No</td><td>NTLM domain when <code>authenticationType=NTLM</code>.</td></tr>
<tr><td><code><#noparse>${lib_Microsoft_Sharepoint.onPrem.http.ntlmTransportMode}</#noparse></code></td><td>Optional</td><td>No</td><td>NTLM transport strategy: <code>connectorOnly</code>, <code>connectorThenHttpClient</code> (default), <code>httpClientThenConnector</code>, <code>httpClientOnly</code>.</td></tr>
</table>
Connector `sharepointOnPremHttp` uses shared credentials symbols `onPrem.username` and `onPrem.password.secret`.
For Claims/NTLM farms, set `onPrem.http.authenticationType=NTLM` and `onPrem.http.ntlmDomain` with those same shared credentials symbols.
When `onPrem.http.authenticationType` is `NTLM` (or `Basic`/`BasicPreemptive`), routed on-prem helper calls delegate authentication to the connector and do not force a sequence-level `Authorization: Basic` header.

	<@header toc=toc anchors=anchors heading="##" text="Authentication Model" />
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

	<@header toc=toc anchors=anchors heading="##" text="On-Prem API Coverage" />
- Site/list resolvers: `ResolveSite`, `ResolveList`, `ResolveLibrary` with `provider=onprem`.
- On-prem discovery lists: `ListSiteLists`, `ListSiteDrives` with `provider=onprem`.
- List CRUD: `ListGetItems`, `GetListItem`, `CreateListItem`, `UpdateListItem`, `DeleteListItem` with `provider=onprem`.
- File/folder APIs: `ListItems`, `UploadItemContent`, `DownloadItemContent`, `DeleteItem`, `ListItemVersions`, `RestoreItemVersion`, `MoveItem`, `CopyItem`, `CreateFolder` with `provider=onprem`.
- Share block (`CreateShareLink`, `InviteItemRecipients`, `ListItemPermissions`, `DeleteItemPermission`) stays routed but returns explicit `not_supported_onprem` fallback in on-prem mode.
- Not covered in on-prem mode: Graph batch, Graph subscriptions.

	<@header toc=toc anchors=anchors heading="##" text="Endpoint-Only Test Calls" />
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

	<@header toc=toc anchors=anchors heading="##" text="Logical Test Plan Script" />
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

	<@header toc=toc anchors=anchors heading="##" text="Typical Request Patterns" />
- Site-first pattern: provide `siteHostname` + `sitePath`, and keep `siteId` empty.
- ID-first pattern: provide `siteId` / `listId` / `driveId` directly to skip resolver lookups.
- Drive resolution priority: `driveId` first, then list-based resolution (`listId`/`listName`), then `driveName`.
- List operations use list item identifiers (`itemId` numeric string), while drive operations use Graph drive item identifiers (opaque string).
- For large binary uploads, prefer `UploadItemLargeContent`; for smaller payloads, `UploadItemContent` is usually simpler.
- For routed on-prem calls (`provider=onprem`), drive-item id lookups are not available from Graph IDs: pass `onPremFileServerRelativeUrl` (or a server-relative `itemId`) for version and restore endpoints.

	<@header toc=toc anchors=anchors heading="##" text="Required Azure Permissions" />
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

	<@header toc=toc anchors=anchors heading="##" text="Permissions by Sequence" />
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

	<@header toc=toc anchors=anchors heading="##" text="Known Limitations" />
- If app roles do not include SharePoint scopes, Graph returns `accessDenied` even if token acquisition succeeds.
- `Sites.Selected` requires explicit grant on target sites; otherwise all calls return `403`.
- List item identifiers (usually numeric strings) and drive item identifiers (Graph opaque ids) are different and not interchangeable.
- `CopyItem` returns an asynchronous monitor URL; completion must be polled by client code.
- `CreateGraphSubscription` requires a publicly reachable HTTPS callback endpoint and Graph webhook validation handling.
- `ExecuteGraphBatch` permissions are the union of all operations included in `batchJson.requests`.

	<@header toc=toc anchors=anchors heading="##" text="Payload Examples" />
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
</#if>
<#if on("references") && has(project,"references")>
  	<@header toc=toc anchors=anchors heading="##" text=help("references") />
  	<#list project.references as reference>
    	<@header toc=toc anchors=anchors heading="###" text=reference.label />
    	<@comment text=reference.comment />
  	</#list>
</#if>
<#if on("sequences") && has(project,"sequences")>
  	<@header toc=toc anchors=anchors heading="##" text=help("sequences") />
  	<#list project.sequences as sequence>
    	<@header toc=toc anchors=anchors heading="###" text=sequence.label />
    	<@comment text=sequence.comment />
    	<#if on("variables") && has(sequence,"variables")>
      		<@table title="**"+help("variables")+"**" headers=["name","comment"] rows=sequence.variables />
    	</#if>
  </#list>
</#if>
<#if on("connectors") && has(project,"connectors")>
  	<@header toc=toc anchors=anchors heading="##" text=help("connectors") />
  	<#list project.connectors as connector>
    	<@header toc=toc anchors=anchors heading="###" text=connector.label />
    	<@comment text=connector.comment />
    	<#if on("transactions") && has(connector,"transactions")>
      		<@header toc=toc anchors=anchors heading="####" text=help("transactions") />
      		<#list connector.transactions as transaction>
        		<@header toc=toc anchors=anchors heading="#####" text=transaction.label />
        		<@comment text=transaction.comment />
        		<#if on("variables") && has(transaction,"variables")>
          			<@table title="**"+help("variables")+"**" headers=["name","comment"] rows=transaction.variables />
        		</#if>
      		</#list>
    	</#if>
  	</#list>
</#if>
<#if on("urlmapper") && has(project,"urlmapper")>
  	<@header toc=toc anchors=anchors heading="##" text=help("urlmapper") />
  	<@comment text=project.urlmapper.comment />
  	<#if on("mappings") && has(project.urlmapper,"mappings")>
	  	<@header toc=toc anchors=anchors heading="###" text=help("mappings") />
	  	<#list project.urlmapper.mappings as mapping>
	    	<@header toc=toc anchors=anchors heading="####" text=mapping.label />
	    	<@comment text=mapping.comment />
	    	<#if on("operations") && has(mapping,"operations")>
	      		<@header toc=toc anchors=anchors heading="#####" text=help("operations") />
	      		<#list mapping.operations as operation>
	        		<@header toc=toc anchors=anchors heading="######" text=operation.label />
	        		<@comment text=operation.comment />
	        		<#if on("parameters") && has(operation,"parameters")>
	          			<@table title="**"+help("parameters")+"**" headers=["name","comment"] rows=operation.parameters />
	        		</#if>
	      		</#list>
	    	</#if>
	  </#list>
	</#if>
</#if>
<#if on("mobileapp") && has(project,"mobileapp")>
	<#assign appname = (project.mobileapp.applicationName?length > 0)
			?string(project.mobileapp.applicationName, (project.name?starts_with("lib_"))?string(help("mobilelib"),help("mobileapp"))) />
  	<@header toc=toc anchors=anchors heading="##" text=appname />
  	<@comment text=project.mobileapp.comment />
  	<#if on("pages") && has(project.mobileapp,"pages")>
	  	<@header toc=toc anchors=anchors heading="###" text=help("pages") />
	  	<#list project.mobileapp.pages as page>
	    	<@header toc=toc anchors=anchors heading="####" text=page.label />
	    	<@comment text=page.comment />
 	  </#list>
	</#if>
  	<#if on("actions") && has(project.mobileapp,"actions")>
	  	<@header toc=toc anchors=anchors heading="###" text=help("actions") />
	  	<#list project.mobileapp.actions as action>
	    	<@header toc=toc anchors=anchors heading="####" text=action.label />
	    	<@comment text=action.comment />
    		<#if on("variables") && has(action,"variables")>
      			<@table title="**"+help("variables")+"**" headers=["name","comment"] rows=action.variables />
    		</#if>
	  </#list>
	</#if>
  	<#if on("components") && has(project.mobileapp,"components")>
	  	<@header toc=toc anchors=anchors heading="###" text=help("components") />
	  	<#list project.mobileapp.components as component>
	    	<@header toc=toc anchors=anchors heading="####" text=component.label />
	    	<@comment text=component.comment />
    		<#if on("variables") && has(component,"variables")>
      			<@table title="**"+help("variables")+"**" headers=["name","comment"] rows=component.variables />
    		</#if>
    		<#if on("events") && has(component,"events")>
      			<@table title="**"+help("events")+"**" headers=["name","comment"] rows=component.events />
    		</#if>
	  </#list>
	</#if>
</#if>
</#assign>


<#-- output project name and comment -->
${intro}
<#-- output project.md link -->
${help("more.info")} : [documentation](./project.md)

<#-- output table of content -->
<#if on("toc")>${toc}</#if>

<#-- output project content -->
${content}
