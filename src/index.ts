#!/usr/bin/env node
import * as program from 'Commander';
import { getAuth, IAuthResponse } from 'node-sp-auth';
import * as request from 'request-promise';
const Preferences: any = require('preferences');
import { IUserCustomAction } from './interfaces';
const colors = require('colors/safe');

program
  .version('0.1.0')
  .option('-c, --connect <siteurl>', 'Connect to SharePoint Online', null)
  .option('-w, --web', 'Show extentions at the web level')
  .option('-s, --sitecollection', 'Show extentions at the site collection level')
  .option('-l, --list <list title>', 'Show extentions at the list level')
  .parse(process.argv);

const prefs = new Preferences('vman.sp.extentions.cli', {
  siteUrl: '',
  authHeaders: null
});

if (program.connect) {
  prefs.siteUrl = program.connect;
  getAuth(prefs.siteUrl, {
    ondemand: true,
    electron: require('electron'),
    force: false,
    persist: true
  }).then((authResponse: IAuthResponse) => {

    prefs.authHeaders = authResponse.headers;
    prefs.authHeaders.Accept = 'application/json;odata=nometadata';
  });
}

if (program.web) {
  displayExtentions('web');
}

if (program.sitecollection) {
  displayExtentions('site');
}

if (program.list) {
  displayExtentions(`web/lists/GetByTitle('${program.list}')`);
}

function displayExtentions(scope: string) {
  if (!prefs.siteUrl) {
    console.error('Please use --connect <siteurl');
  }

  request.get({
    url: `${prefs.siteUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location, 'ClientSideExtension')
    &$select=ClientSideComponentId,Title,Location,ClientSideComponentProperties`,
    headers: prefs.authHeaders
  }).then((response: any) => {
    const userCustomActions: IUserCustomAction[] = JSON.parse(response).value;

    console.log(colors.magenta(`${scope} level spfx extentions at ${prefs.siteUrl}:`));
    console.log(colors.yellow('Title, ClientSideComponentId, Location, ClientSideComponentProperties'));

    for (const uca of userCustomActions) {
      console.log(colors.green([uca.Title, uca.ClientSideComponentId, uca.Location, uca.ClientSideComponentProperties].join(', ')));
    }
  }, (error: Error) => {
    console.log(colors.red(error.message));
  });
}