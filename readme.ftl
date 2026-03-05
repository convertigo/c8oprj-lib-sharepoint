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
</table>

	<@header toc=toc anchors=anchors heading="##" text="Authentication Model" />
- Default mode (recommended for backend use): rely on server-side symbols (`tenantId`, `clientId`, `clientSecret`) and do not pass credentials in calls.
- Delegated mode: pass `accessToken`; tenant/client/secret are ignored.
- Application override: leave `accessToken` empty and pass tenant/client/secret explicitly when needed.
- All online Graph sequences are `Hidden` and `authenticatedContextRequired=true`.
- Sequence responses expose `tokenMode` (`delegated` or `application`) for diagnostics.

	<@header toc=toc anchors=anchors heading="##" text="Endpoint-Only Test Calls" />
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
- For large binary uploads, prefer `UploadDriveItemLargeContent`; for smaller payloads, `UploadDriveItemContent` is usually simpler.

	<@header toc=toc anchors=anchors heading="##" text="Required Azure Permissions" />
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

	<@header toc=toc anchors=anchors heading="##" text="Permissions by Sequence" />
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

	<@header toc=toc anchors=anchors heading="##" text="Known Limitations" />
- If app roles do not include SharePoint scopes, Graph returns `accessDenied` even if token acquisition succeeds.
- `Sites.Selected` requires explicit grant on target sites; otherwise all calls return `403`.
- List item identifiers (usually numeric strings) and drive item identifiers (Graph opaque ids) are different and not interchangeable.
- `CopyDriveItem` returns an asynchronous monitor URL; completion must be polled by client code.

	<@header toc=toc anchors=anchors heading="##" text="Payload Examples" />
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
