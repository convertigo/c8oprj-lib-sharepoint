


# lib_Sharepoint

This is is the Sharepoint Connector for Convertigo. Use this connector to access On premises versions of SharePoint server. This connector will not work to access Azure Cloud Sharepoint Services



For more technical informations : [documentation](./project.md)

- [Installation](#installation)
- [Sequences](#sequences)
    - [ListGetItems](#listgetitems)


## Installation

1. In your Convertigo Studio click on ![](https://github.com/convertigo/convertigo/blob/develop/eclipse-plugin-studio/icons/studio/project_import.gif?raw=true "Import a project in treeview") to import a project in the treeview
2. In the import wizard

   ![](https://github.com/convertigo/convertigo/blob/develop/eclipse-plugin-studio/tomcat/webapps/convertigo/templates/ftl/project_import_wzd.png?raw=true "Import Project")
   
   paste the text below into the `Project remote URL` field:
   <table>
     <tr><td>Usage</td><td>Click the copy button at the end of the line</td></tr>
     <tr><td>To contribute</td><td>

     ```
     lib_Sharepoint=https://github.com/convertigo/c8oprj-lib-sharepoint.git:branch=master
     ```
     </td></tr>
     <tr><td>To simply use</td><td>

     ```
     lib_Sharepoint=https://github.com/convertigo/c8oprj-lib-sharepoint/archive/master.zip
     ```
     </td></tr>
    </table>
3. Click the `Finish` button. This will automatically import the __lib_Sharepoint__ project


## Sequences

### ListGetItems

List Items from a Sharepoint list

**variables**

<table>
<tr>
<th>name</th><th>comment</th>
</tr>
<tr>
<td>domain</td><td>The NT Domain</td>
</tr>
<tr>
<td>password</td><td>User's password</td>
</tr>
<tr>
<td>sharepointBase</td><td>Sharepoint base as http(s)://server/mysite</td>
</tr>
<tr>
<td>username</td><td>User accessing the list</td>
</tr>
</table>



