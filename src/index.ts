#!/usr/bin/env node
import * as program from 'Commander';
import { getAuth, IAuthResponse } from 'node-sp-auth';
import * as request from 'request-promise';
const preferences: any = require('preferences');
import { IUserCustomAction } from './interfaces';

program
  .version('0.1.0')
  .option('-c, --connect <siteurl>', 'Connect to SharePoint Online', null)
  .option('-w, --web', 'Show extentions at the web level')
  .option('-s, --sitecollection', 'Show extentions at the site collection level')
  .option('-l, --list <list title>', 'Show extentions at the list level')
  .parse(process.argv);

const prefs = new preferences('vman.sp.extentions.cli', {
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

function displayExtentions(path: string) {
  if (!prefs.siteUrl) {
    console.error('Please use --connect <siteurl');
  }

  request.get({
    url: `${prefs.siteUrl}/_api/${path}/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`,
    headers: prefs.authHeaders
  }).then((response: any) => {
    const userCustomActions: IUserCustomAction[] = JSON.parse(response).value;

    for (const uca of userCustomActions) {
      console.log(`Title: ${uca.Title},
                   ClientSideComponentId: ${uca.ClientSideComponentId},
                   ClientSideComponentProperties: ${uca.ClientSideComponentProperties},
                   Location: ${uca.Location}`);
    }
  }, (error: Error) => {
    console.error(error.message);
  });
}