# spfx-extensions-cli

[![NPM](https://nodei.co/npm/spfx-extensions-cli.png?compact=true)](https://nodei.co/npm/spfx-extensions-cli/)

CLI tool to view SharePoint Framework extensions currently installed on a Site Collection, Site or a List.

> Note: SharePoint Framework extensions RC is out now. It is not recommended to use them in production yet. The functionality of this tool might change after extensions reach GA.

![Working of the spfx-extensions-cli tool](https://github.com/vman/spfx-extensions-cli/raw/master/assets/cli.gif "spfx-cli-extensions")

#### Install:

`npm install spfx-extensions-cli -g`

#### Help

`spfx-ext --help`

#### Authenticate:

`spfx-ext --connect "https://yourtenant.sharepoint.com/sites/team"`

#### Get sitecollection level extensions:

`spfx-ext --sitecollection`

#### Get web level extensions:

`spfx-ext --web`

#### Get list level extensions by list title:

`spfx-ext --list "My List"`

#### Add an extension to a sitecollection or web:

> Adding an extension to a site is mainly useful for tenant scoped extensions. Make sure the `.sppkg` file is uploaded/deployed in the app catalog and the extension is available to be added to a site without activating any features. After that, `spfx-ext add` can be used to add the extention to a perticular site collection or web.

`spfx-ext add <title> <extensionType> <scope> <clientSideComponentId> --registrationid --registrationType --clientprops`



For help, type

`spfx-ext add --help`

Examples:


`spfx-ext add "My App Customizer" ApplicationCustomizer sitecollection f5c5285d-0141-42e5-b198-044433cd3d0c`

`spfx-ext add "Another App Customizer" ApplicationCustomizer web 412b8279-2e5b-4546-a554-2f3a6ccf801a`

#### Remove an extention from the sitecollection or web:

`spfx-ext remove <scope> <id>`

For help, type:

`spfx-ext remove --help`

Examples:


`spfx-ext remove web b424419b-af2f-4748-bd76-503fe1bd567a`

`spfx-ext remove sitecollection 92b384c7-4a78-4ad1-b6c6-a9c2d85c18b5`


>Note: Adding/Removing FieldCustomizers is not implemented in spfx-extentions-cli at this time. 


