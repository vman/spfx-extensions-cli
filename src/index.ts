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
  .option('-l, --list <listtitle>', 'Show extentions at the list level')
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
  if (!prefs.siteUrl) {
    console.error('Please use --connect');
  }

  request.get({
    url: `${prefs.siteUrl}/_api/web/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`,
    headers: prefs.authHeaders
  }).then((response) => {
    const userCustomActions: IUserCustomAction[] = JSON.parse(response).value;

    for (const uca of userCustomActions) {
      console.log(`Title: ${uca.Title},
                   ClientSideComponentId: ${uca.ClientSideComponentId},
                   ClientSideComponentProperties: ${uca.ClientSideComponentProperties},
                   Location: ${uca.Location}`);
    }
  });
}

if (program.sitecollection) {
  if (!prefs.siteUrl) {
    console.error('Please use --connect <siteurl');
  }

  request.get({
    url: `${prefs.siteUrl}/_api/site/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`,
    headers: prefs.authHeaders
  }).then((response) => {
    const userCustomActions: IUserCustomAction[] = JSON.parse(response).value;

    for (const uca of userCustomActions) {
      console.log(`Title: ${uca.Title},
                   ClientSideComponentId: ${uca.ClientSideComponentId},
                   ClientSideComponentProperties: ${uca.ClientSideComponentProperties},
                   Location: ${uca.Location}`);
    }
  });
}

if (program.list) {
  if (!prefs.siteUrl) {
    console.error('Please use --connect <siteurl>');
  }

  request.get({
    url: `${prefs.siteUrl}/_api/web/lists/GetByTitle('${program.list}')/UserCustomActions?
    $filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`,
    headers: prefs.authHeaders
  }).then((response) => {
    const userCustomActions: IUserCustomAction[] = JSON.parse(response).value;

    for (const uca of userCustomActions) {
      console.log(`Title: ${uca.Title},
                   ClientSideComponentId: ${uca.ClientSideComponentId},
                   ClientSideComponentProperties: ${uca.ClientSideComponentProperties},
                   Location: ${uca.Location}`);
    }
  });
}