# spfx-extensions-cli (preview)

CLI tool to view SharePoint Framework extensions currently installed on a Site Collection, Site or a List.

> Note: SharePoint Framework extensions are in preview right now. It is not recommended to use them in production yet. The functionality of this tool might change after extensions reach GA.

![Working of the spfx-extensions-cli tool](https://github.com/vman/spfx-extensions-cli/raw/master/assets/cli.gif "spfx-cli-extensions")

### Install:

`npm install spfx-extensions-cli -g`

### Authenticate:

`spfx-ext --connect "https://yourtenant.sharepoint.com/sites/team"`

### Get sitecollection level extensions:

`spfx-ext --sitecollection`

### Get web level extensions:

`spfx-ext --web`

### Get list level extensions by list title:

`spfx-ext --list "My List"`

