# spfx-extensions-cli

> :warning: I am no longer actively working on this tool mainly because the Microsoft 365 CLI (formerly Office 365 CLI) is a better cross platform alternative to manage the SPFx extensions among many other things.

> Here is a blogpost which might help you: https://sharepoint.handsontek.net/2018/02/11/mange-spfx-extensions-from-macos-with-office-365-cli/

> and the customaction docs for the M365 CLI: https://pnp.github.io/cli-microsoft365/cmd/spo/customaction/customaction-add/

---

[![NPM](https://nodei.co/npm/spfx-extensions-cli.png?compact=true)](https://nodei.co/npm/spfx-extensions-cli/)

CLI tool to manage SharePoint Framework extensions currently installed on a Site Collection, Site or a List.

> Note: SharePoint Framework extensions RC is out now. It is not recommended to use them in production yet. The functionality of this tool might change after extensions reach GA.

![Working of the spfx-extensions-cli tool](https://github.com/vman/spfx-extensions-cli/raw/master/assets/cli.gif "spfx-cli-extensions")

#### Install:

`npm install spfx-extensions-cli -g`

#### Help

`spfx-ext --help`

#### Authenticate:

`spfx-ext --connect "https://yourtenant.sharepoint.com/sites/team"`

#### Get sitecollection level extensions:

`spfx-ext --site`

#### Get web level extensions:

`spfx-ext --web`

#### Get list level extensions by list title:

`spfx-ext --list "My List"`

#### Add an extension:

> Adding an extension to a site is mainly useful for tenant scoped extensions. Make sure the `.sppkg` file is uploaded/deployed in the app catalog and the extension is available to be added to a site without activating any features. After that, `spfx-ext add` can be used to add the extension to a perticular site collection, web or list.


`spfx-ext add <title> <extensionType> <scope> <clientSideComponentId> --registrationid --registrationType --clientprops`



For help, type

`spfx-ext add --help`

Examples:


`spfx-ext add "SiteCollection App Customizer" ApplicationCustomizer site f5c5285d-0141-42e5-b198-044433cd3d0c`

`spfx-ext add "App Customizer with Props" ApplicationCustomizer web f7b1ca4a-705d-45f6-a072-3803748556a9 --clientProps "{\"Top\":\"Top area\",\"Bottom\":\"Bottom area\"}"`

`spfx-ext add "List CommandSet" ListViewCommandSet.CommandBar list 297808d9-98da-44c7-a697-0605fc4062b7 --listtitle "Documents"`

`spfx-ext add "List CommandSet" ListViewCommandSet.CommandBar web 297808d9-98da-44c7-a697-0605fc4062b7 --registrationId 100 --registrationType List`

#### Remove an extension:

`spfx-ext remove <scope> <id> --listtitle`

For help, type:

`spfx-ext remove --help`

Examples:


`spfx-ext remove web b424419b-af2f-4748-bd76-503fe1bd567a`

`spfx-ext remove site 92b384c7-4a78-4ad1-b6c6-a9c2d85c18b5`

`spfx-ext remove list --listtitle "Documents" 38f4ce5c-e447-4199-9f56-fd9f96370cfd`


>Note: Adding/Removing FieldCustomizers is not implemented in spfx-extensions-cli at this time. 


