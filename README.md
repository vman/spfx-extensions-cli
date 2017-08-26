# spfx-extentions-cli (preview)

CLI tool to view SharePoint Framework extentions currently installed on a Site Collection, Site or a List.

> Note: SharePoint Framework extentions are in preview right now. It is not recommended to use them in production yet. The functionality of this tool might change after extentions reach GA.

![Working of the spfx-extentions-cli tool](https://github.com/vman/spfx-extentions-cli/raw/master/assets/cli.gif "spfx-cli-extentions")

### Install:

`npm install spfx-extentions-cli -g`

### Authenticate:

`spfx-ext --connect "https://yourtenant.sharepoint.com/sites/team"`

### Get sitecollection level extentions:

`spfx-ext --sitecollection`

### Get web level extentions:

`spfx-ext --web`

### Get list level extentions by list title:

`spfx-ext --list "My List"`

